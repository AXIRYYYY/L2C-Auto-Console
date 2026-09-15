"""整本凭证 PDF 按 年月+基础凭证号 拆分。

用途：解析凭证 PDF 每页的「日期:」「凭证号:」文本（兼容全角/半角冒号与暂记凭证格式），
      按 (年月, 基础凭证号) 分组另存为独立 PDF，文件名带页数便于纸质核对。

输入：整本记账凭证 PDF（弹窗选择，或命令行参数传入）
输出：{凭证名}_拆分结果/{YYYY-MM}-{基础凭证号}-{N}页.pdf

用法：
    python split_vouchers_凭证PDF拆分.py                     # 弹窗选择 PDF
    python split_vouchers_凭证PDF拆分.py 记账凭证-202601.pdf  # 命令行直接运行（无弹窗）

依赖：pypdf
对应回放：docs/demo/replay/voucher_split.json
"""

import os
import re
from collections import defaultdict
from pypdf import PdfReader, PdfWriter
import tkinter as tk
from tkinter import filedialog
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

def process_voucher_pdf_final_with_page_count():
    """
    主函数，用于选择、读取、拆分和合并凭证PDF文件。
    【已更新】兼容全角和半角冒号。
    合并后的文件直接保存在输出目录中，并以'YYYY-MM-基础凭证号-X页.pdf'格式命名。
    """
    # 解决在某些打包环境下Tkinter窗口的路径问题
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 项目根目录

    # --- 1. 选择 PDF 文件：命令行参数优先，否则弹窗 ---
    cli_args = [a for a in sys.argv[1:] if a.strip()]
    if cli_args:
        file_path = cli_args[0]
        print(f"[CLI] 使用命令行参数: {file_path}")
    else:
        print("正在打开文件选择窗口...")
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        file_path = filedialog.askopenfilename(
            initialdir=application_path,
            title="请选择需要处理的凭证PDF文件",
            filetypes=[("PDF files", "*.pdf")]
        )

    if not file_path:
        print("操作已取消：用户未选择任何文件。")
        return

    print(f"已选择文件: {file_path}")

    # --- 2. 设置输入和输出路径 ---
    try:
        source_dir = os.path.dirname(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_dir = os.path.join(source_dir, f"{base_name}_拆分结果")
        os.makedirs(output_dir, exist_ok=True)
        print(f"结果将保存在: {os.path.abspath(output_dir)}")
    except Exception as e:
        print(f"错误：无法创建输出目录。{e}")
        return

    # --- 3. 读取PDF并按凭证号分组 ---
    vouchers = defaultdict(list)
    try:
        reader = PdfReader(file_path)
        total_pages = len(reader.pages)
        print(f"文件共有 {total_pages} 页，开始逐页解析...")

        for i, page in enumerate(reader.pages):
            page_num = i + 1
            text = page.extract_text()
            if not text:
                print(f"警告: 第 {page_num} 页无法提取文本，将跳过。")
                continue

            # 使用 [：:] 来同时匹配全角和半角冒号
            date_match = re.search(r"日期\s*[:：]\s*(\d{4}-\d{2}-\d{2})", text)
            voucher_match = re.search(r"凭证号\s*[:：]\s*([\w-]+/\d+)", text)

            if date_match and voucher_match:
                date = date_match.group(1)
                voucher_full = voucher_match.group(1)
                
                # 提取年份和月份
                year_month = date[:7]  # 格式: YYYY-MM
                
                # 提取基础凭证号
                base_voucher_match = re.match(r'([\w-]+)-\d+/\d+', voucher_full)
                if base_voucher_match:
                    base_voucher = base_voucher_match.group(1)
                else:
                    base_voucher = voucher_full

                # 将页面对象按 (年份月份, 基础凭证号) 进行分组
                vouchers[(year_month, base_voucher)].append(page)
            else:
                # 检查是否存在暂记凭证
                temp_voucher_match = re.search(r"凭证号\s*[:：]\s*([\w-]+-\d+/\d+)", text)
                if date_match and temp_voucher_match:
                    date = date_match.group(1)
                    voucher_full = temp_voucher_match.group(1)
                    year_month = date[:7]
                    base_voucher = '-'.join(voucher_full.split('-')[:-1])
                    vouchers[(year_month, base_voucher)].append(page)
                else:
                    print(f"警告: 第 {page_num} 页未能找到标准的日期或凭证号格式，将跳过。")


    except Exception as e:
        print(f"错误：读取或解析PDF文件时发生错误。{e}")
        return

    if not vouchers:
        print("处理完成，但未找到任何可供分类的凭证页。请检查PDF内容格式。")
        return
        
    # --- 4. 写入新的PDF文件 ---
    print("\n开始生成新的PDF文件...")
    file_count = 0
    for (year_month, base_voucher), pages in vouchers.items():
        try:
            # **【修改】** 获取当前凭证的页数
            page_count = len(pages)
            
            # **【修改】** 在文件名中加入页数
            output_filename = os.path.join(output_dir, f"{year_month}-{base_voucher}-{page_count}页.pdf")
            
            # 创建PDF写入器并添加页面
            writer = PdfWriter()
            for page in pages:
                writer.add_page(page)
            
            # 将内容写入文件
            with open(output_filename, "wb") as f:
                writer.write(f)
            
            print(f"  -> 已创建文件: {output_filename}")
            file_count += 1
        except Exception as e:
            print(f"错误：写入文件 {base_voucher}.pdf 时失败。{e}")

    print(f"\n处理完成！共生成 {file_count} 个凭证文件。")
    print(f"所有文件已保存至绝对路径: {os.path.abspath(output_dir)}")


if __name__ == "__main__":
    process_voucher_pdf_final_with_page_count()