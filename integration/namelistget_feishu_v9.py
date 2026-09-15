"""月底资料包匹配整合（飞书报销单 + 附件 + 关联申请单）。

用途：按关键词从源文件夹匹配「报销单文件夹 + 附件 + 关联申请单」，复制归集、
      图片附件转 PDF 后合并为整包 PDF，并提取成本编号生成对账表与错误报告。

输入：namelist.xlsx（A 列关键词，需与本脚本放在同一文件夹）、关键词源文件夹、差旅申请单搜索文件夹、关联申请单 PDF
输出（生成在脚本所在文件夹）：搜索结果/{关键词}/ 资料包、{关键词}.pdf、关键词与PDF文件对应表.xlsx、合并文件报错信息.xlsx

用法：
    python namelistget_feishu_v9.py                                # 全部弹窗选择
    python namelistget_feishu_v9.py 申请单文件夹 关键词源文件夹      # 命令行直接运行（无弹窗）

依赖：openpyxl, pandas, PyMuPDF(fitz), Pillow, pypdf
对应回放：docs/demo/replay/namelist_pack.json
"""

import os
import shutil
import sys
from openpyxl import load_workbook
import time
import pandas as pd
from io import BytesIO
import re

# 导入tkinter用于弹出选择窗口
import tkinter as tk
from tkinter import filedialog

CLI_DIRS = [a for a in sys.argv[1:] if a.strip()]
USE_CLI = len(CLI_DIRS) >= 2

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# --- 核心改动：导入新的、更强大的PDF处理库 ---
try:
    import fitz  # PyMuPDF库的导入名是fitz
    from PIL import Image
except ImportError:
    print("错误：缺少必要的库。请先通过命令提示符安装 PyMuPDF 和 Pillow。")
    print("安装命令: pip install PyMuPDF Pillow")
    if not USE_CLI:
        input("按回车键退出。")
    sys.exit(1)

# pypdf 仍然用于合并，因为它在这方面很稳定
from pypdf import PdfWriter, PdfReader

def merge_folders(source, destination):
    """递归合并两个文件夹，处理文件覆盖和子文件夹"""
    if not os.path.exists(destination): os.makedirs(destination)
    for item in os.listdir(source):
        s, d = os.path.join(source, item), os.path.join(destination, item)
        if os.path.isdir(s): merge_folders(s, d)
        else:
            if os.path.exists(d): print(f"警告：文件 {d} 已存在，将被覆盖")
            shutil.copy2(s, d)

def copy_files_based_on_keywords():
    # --- 1. 初始化设置 ---
    script_dir = os.path.dirname(os.path.abspath(__file__))  # 脚本所在目录（= 数据目录）
    print(f"数据目录: {script_dir}")
    
    print("准备中...")
    if USE_CLI:
        application_search_path = CLI_DIRS[0]
        print(f"[CLI] 申请单搜索文件夹: {application_search_path}")
    else:
        root = tk.Tk()
        root.withdraw()
        application_search_path = filedialog.askdirectory(title="请选择一个文件夹用于搜索【差旅申请单】")

    if not application_search_path:
        print("操作已取消：您没有选择任何文件夹。")
        if not USE_CLI:
            input("按回车键退出。")
        return
        
    print(f"已选定申请单搜索文件夹: {application_search_path}")
    
    excel_file_path = os.path.join(script_dir, "namelist.xlsx")
    result_folder_path = os.path.join(script_dir, "搜索结果")
    if not os.path.exists(result_folder_path): os.makedirs(result_folder_path)
    
    merge_errors = []

    try:
        workbook = load_workbook(excel_file_path)
    except FileNotFoundError:
        print(f"错误：无法找到 {excel_file_path} 文件。"); input("按回车键退出。"); return
        
    sheet = workbook.active
    keywords = [str(cell.value).strip() for cell in sheet['A'] if cell.value is not None and str(cell.value).strip()]
    print(f"从 {excel_file_path} 中获取到了 {len(keywords)} 个关键词: {keywords}")
    if not USE_CLI:
        input("按回车继续")

    # --- 新增配置：选择关键词文件的搜索源 ---
    # 1 = 在脚本当前文件夹搜索 (默认)
    # 2 = 弹窗选择其他文件夹
    KEYWORD_SOURCE_MODE = 2  # <--- 修改这里：1为自动，2为手动

    keyword_search_root = script_dir  # 默认搜索脚本所在目录
    if USE_CLI:
        keyword_search_root = CLI_DIRS[1]
        print(f"[CLI] 关键词搜索源: {keyword_search_root}")
    elif KEYWORD_SOURCE_MODE == 2:
        print("\n模式 2: 准备选择【包含关键词文件】的源文件夹...")
        root = tk.Tk()
        root.withdraw()
        # 注意：这里是选包含“张三-住宿费”这种文件的文件夹，不是选申请单
        selected_path = filedialog.askdirectory(title="请选择要搜索关键词文件的文件夹（源文件所在位置）")
        if selected_path:
            keyword_search_root = selected_path
            print(f"已选定关键词搜索源: {keyword_search_root}")
        else:
            print("未选择文件夹，将默认回退到搜索脚本当前目录。")
    
    # --- 2. 扫描与规划 ---
    print("\n--- 开始第一阶段：扫描（不含申请单存档文件夹）... ---")
    actions_to_take = {kw: {"folders": [], "pdfs": []} for kw in keywords}
    print(f"当前搜索范围: {keyword_search_root}")
    # 将 script_dir 替换为上面定义好的 keyword_search_root
    for root, dirs, files in os.walk(keyword_search_root):
        abs_root = os.path.abspath(root)
        # 修复：只排除“搜索结果”文件夹，不再排除申请单文件夹，防止同目录搜索时被跳过
        if abs_root.startswith(os.path.abspath(result_folder_path)): continue
        for keyword in keywords:
            for dir_name in dirs:
                if f"-{keyword}" in dir_name and dir_name.endswith(keyword):
                    actions_to_take[keyword]["folders"].append(os.path.join(root, dir_name))
            for file in files:
                if f"-{keyword}" in file and file.lower().endswith('.pdf'):
                    actions_to_take[keyword]["pdfs"].append(os.path.join(root, file))
    print("扫描完成。")

    # --- 3. 筛选、执行复制并查找关联单据 ---
    print("\n--- 开始第二阶段：处理文件并查找关联申请单号 ---")
    file_count, folder_count = 0, 0
    keyword_folders_dest, pdf_keyword_map = {}, {}
    linked_docs_info = {}

    for keyword in keywords:
        if actions_to_take[keyword]["pdfs"]:
            print(f"\n处理关键词 '{keyword}'...")
            if actions_to_take[keyword]["folders"]:
                for spath in actions_to_take[keyword]["folders"]:
                    dpath = os.path.join(result_folder_path, os.path.basename(spath))
                    keyword_folders_dest[keyword] = dpath
                    try:
                        if os.path.exists(dpath): merge_folders(spath, dpath)
                        else: shutil.copytree(spath, dpath)
                        folder_count += 1; print(f"  - 已处理文件夹: {spath}")
                    except Exception as e: print(f"  - 处理文件夹 {spath} 时出错：{e}")
            
            for spath in actions_to_take[keyword]["pdfs"]:
                fname = os.path.basename(spath)
                dfolder = keyword_folders_dest.get(keyword, result_folder_path)
                dpath = os.path.join(dfolder, fname)
                try:
                    os.makedirs(os.path.dirname(dpath), exist_ok=True)
                    shutil.copy2(spath, dpath)
                    file_count += 1; print(f"  - 已复制PDF: {spath}")
                    if keyword not in pdf_keyword_map: pdf_keyword_map[keyword] = []
                    pdf_keyword_map[keyword].append(os.path.splitext(fname)[0])
                    
                    with fitz.open(dpath) as doc:
                        text = doc[0].get_text()
                        if "差旅报销单" in text:
                            match = re.search(r'[A-Z]{2,4}CLSQ\d+', text)
                            if match:
                                linked_docs_info[keyword] = match.group(0)
                                print(f"  - 发现关联申请单号: {linked_docs_info[keyword]}")
                except Exception as e: print(f"  - 复制或读取文件 {fname} 时出错: {e}")

    # --- 复制关联申请单并提取成本编号 ---
    print(f"\n--- 开始在您选择的文件夹中搜索、复制并提取信息 ---")
    cost_center_info = {}
    if not os.path.isdir(application_search_path):
        msg = f"查找关联申请单失败：指定的搜索文件夹不存在。"
        print(f"警告: {msg} 路径: {application_search_path}")
        if linked_docs_info:
            for keyword, clsq_num in linked_docs_info.items():
                merge_errors.append({'关键词': keyword, '问题文件': f"关联单号: {clsq_num}", '错误信息': msg})
    elif not linked_docs_info:
        print("未发现需要处理的关联申请单。")
    else:
        for keyword, clsq_num in linked_docs_info.items():
            found = False
            for root, dirs, files in os.walk(application_search_path):
                for file in files:
                    if clsq_num in file and file.lower().endswith('.pdf'):
                        
                        # --- ##### 核心修复：检查附件文件夹是否存在，记录错误但继续复制 ##### ---
                        if keyword in keyword_folders_dest:
                            # 正常情况：存在对应的附件文件夹
                            dfolder = keyword_folders_dest[keyword]
                        else:
                            # 异常情况：附件文件夹缺失
                            dfolder = result_folder_path # 降级：放入根目录
                            msg = "警告：未找到该关键词对应的独立附件文件夹，文件被复制到根目录。"
                            print(f"    - {msg}")
                            # 记录到错误列表，以便后续输出到 Excel
                            merge_errors.append({'关键词': keyword, '问题文件': '附件文件夹缺失', '错误信息': msg})
                        # --- ##### 修复结束 ##### ---

                        source_path = os.path.join(root, file)
                        dest_path = os.path.join(dfolder, file)
                        try:
                            print(f"  - 找到并复制 '{file}' 到 '{os.path.basename(dfolder)}'")
                            shutil.copy2(source_path, dest_path)
                            found = True
                            
                            with fitz.open(dest_path) as doc:
                                pdf_text = "".join([page.get_text() for page in doc])
                                clean_text = " ".join(pdf_text.replace('\n', ' ').replace('\r', ' ').split())
                                start_anchor = "成本编号/名称"
                                end_anchor = "外出开始时间-外出"
                                exclude_text = "成本编号:与项目或客户等相关的成本,请务必填写成本编号,非必要不可选择“无”"
                                cost_info = ""
                                pattern = re.escape(start_anchor) + r'(.*?)' + re.escape(end_anchor)
                                match = re.search(pattern, clean_text)
                                if match:
                                    content_between = match.group(1)
                                    cost_info = content_between.replace(exclude_text, "").strip()
                                if cost_info:
                                    cost_center_info[keyword] = cost_info
                                    print(f"    - 成功提取成本编号: {cost_info}")
                                else:
                                    print(f"    - 警告: 在 {file} 中未能根据规则提取到有效的成本编号。")
                            break 
                        except Exception as e: print(f"  - 错误: 复制或读取 {file} 时失败: {e}")
                        
                if found: break
            if not found:
                msg = f"查找关联申请单失败：未能在指定文件夹中找到单号为 '{clsq_num}' 的PDF。"
                print(f"  - 警告: {msg}")
                merge_errors.append({'关键词': keyword, '问题文件': f"关联单号: {clsq_num}", '错误信息': msg})

    # --- 4. 合并文件为PDF ---
    print("\n--- 开始第三阶段：合并文件夹内容为PDF ---")
    for keyword, folder_path in keyword_folders_dest.items():
        print(f"\n正在处理 '{keyword}' 的合并任务...")
        try:
            main_pdf, linked_pdf, attachments = None, None, []
            clsq_num = linked_docs_info.get(keyword)
            
            for filename in os.listdir(folder_path):
                full_path = os.path.join(folder_path, filename)
                if f"-{keyword}" in filename and not main_pdf: main_pdf = full_path
                elif clsq_num and clsq_num in filename and not linked_pdf: linked_pdf = full_path
                else: attachments.append(full_path)
            
            files_to_merge = []
            if main_pdf: files_to_merge.append(main_pdf)
            else: print(f"  - 警告：未找到主报销单PDF。")
            if linked_pdf: files_to_merge.append(linked_pdf)
            
            processed = {p for p in [main_pdf, linked_pdf] if p}
            for att_path in sorted(attachments):
                if att_path not in processed: files_to_merge.append(att_path)
            
            if not files_to_merge: print(f"  - 文件夹为空，跳过合并。"); continue

            pdf_writer = PdfWriter()
            keyword_errors = [] 
            merge_will_fail = False

            if clsq_num and not linked_pdf:
                msg = f"未能找到关联的申请单PDF({clsq_num})。"
                print(f"  - 失败: {msg}")
                merge_will_fail = True

            for file_path in files_to_merge:
                file_ext = os.path.splitext(file_path)[1].lower()
                try:
                    if file_ext == '.pdf':
                        with open(file_path, 'rb') as f:
                            reader = PdfReader(f); [pdf_writer.add_page(p) for p in reader.pages]
                    elif file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']:
                        with Image.open(file_path) as img:
                            if img.mode == 'RGBA': img = img.convert('RGB')
                            pdf_buffer = BytesIO()
                            img.save(pdf_buffer, 'PDF', resolution=100.0)
                            pdf_buffer.seek(0)
                            with PdfReader(pdf_buffer) as image_pdf_reader:
                                pdf_writer.add_page(image_pdf_reader.pages[0])
                    else: 
                        raise TypeError(f"不支持的文件类型 '{file_ext}'")
                except Exception as e:
                    msg = f"无法处理文件: {e}"
                    print(f"  - 失败: {msg} (文件: {file_path})")
                    keyword_errors.append({'关键词': keyword, '问题文件': file_path, '错误信息': msg})
                    merge_will_fail = True
            
            if merge_will_fail:
                print(f"  - 由于遇到错误，已中止 '{keyword}' 的合并。")
                merge_errors.extend(keyword_errors)
                try:
                    original_name = os.path.basename(folder_path)
                    if not original_name.startswith("合并失败_"):
                        new_name = f"合并失败_{original_name}"
                        os.rename(folder_path, os.path.join(os.path.dirname(folder_path), new_name))
                        print(f"  - 已将文件夹重命名为: {new_name}")
                except Exception as e:
                    merge_errors.append({'关键词': keyword, '问题文件': folder_path, '错误信息': f"重命名文件夹失败: {e}"})
            else:
                output_pdf_path = os.path.join(result_folder_path, f"{keyword}.pdf")
                with open(output_pdf_path, 'wb') as f_out: pdf_writer.write(f_out)
                print(f"  - 成功！已合并 {len(files_to_merge)} 个文件为: {os.path.basename(output_pdf_path)}")

        except Exception as e: 
            merge_errors.append({'关键词': keyword, '问题文件': folder_path, '错误信息': f"合并时发生未知错误: {e}"})

    # --- 5. 创建对应关系表格 ---
    print("\n--- 开始第四阶段：创建最终的状态对应表 ---")
    if pdf_keyword_map:
        try:
            report_data = [{'关键词': k, 'PDF文件名 (不含后缀)': v} for k, l in pdf_keyword_map.items() for v in l]
            df_report = pd.DataFrame(report_data)

            df_report['差旅申请单编号'] = df_report['关键词'].map(linked_docs_info).fillna('')
            df_report['成本编号/名称'] = df_report['关键词'].map(cost_center_info).fillna('')
            
            # --- ##### 核心更新：扩展备注列的筛选规则，包含附件文件夹缺失错误 ##### ---
            clsq_errors = {
                err['关键词']: err['错误信息']
                for err in merge_errors
                if "查找关联申请单失败" in str(err.get('错误信息')) or "附件文件夹" in str(err.get('错误信息'))
            }
            # --- ##### 更新结束 ##### ---
            
            df_report['备注'] = df_report['关键词'].map(clsq_errors).fillna('')

            final_columns = ['关键词', 'PDF文件名 (不含后缀)', '差旅申请单编号', '成本编号/名称', '备注']
            df_report = df_report[final_columns]
            
            df_report = df_report.sort_values(by='关键词', ascending=True)
            
            output_excel_path = os.path.join(result_folder_path, '关键词与PDF文件对应表.xlsx')
            df_report.to_excel(output_excel_path, index=False)
            print(f"已成功创建状态对应表：{os.path.basename(output_excel_path)}")
            
        except Exception as e: print(f"创建状态对应表时出错：{e}")
    else: print("没有可供报告的PDF文件，无需创建对应表格。")
        
    # --- 6. 生成独立的完整错误报告 ---
    print("\n--- 开始第五阶段：生成合并错误报告 ---")
    if merge_errors:
        try:
            error_df = pd.DataFrame(merge_errors)
            error_df = error_df.sort_values(by='关键词', ascending=True)

            error_excel_path = os.path.join(result_folder_path, '合并文件报错信息.xlsx')
            error_df.to_excel(error_excel_path, index=False)
            print(f"发现错误，已生成独立的错误报告: {os.path.basename(error_excel_path)}")
        except Exception as e: print(f"创建合并错误报告Excel时出错: {e}")
    else: print("未发生任何文件合并错误。")

    print(f"\n总共成功处理了 {folder_count} 个文件夹和 {file_count} 个PDF文件。")

if __name__ == "__main__":
    print("开始执行文件搜索、复制与合并任务...")
    copy_files_based_on_keywords()
    print("所有任务执行完毕。")
    if not USE_CLI:
        input("按回车键退出。")