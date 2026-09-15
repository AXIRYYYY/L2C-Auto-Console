"""回放采集器：在隔离工作区中用 sample_data 真实执行核心脚本，采集日志与产物快照。

- 每个目标脚本复制到独立工作区（docs/demo/.replay_work/<slug>/<原目录>/<脚本名>）
  使脚本内部的 script_dir（项目根目录）解析到工作区，避免污染仓库
- tkinter / input() / os.startfile 由 _harness/run_with_patches.py 拦截
- 产物写入 docs/demo/replay/<slug>.json，供 build_console.py 注入控制台

用法：
  python docs/demo/run_replay.py            # 全部目标
  python docs/demo/run_replay.py profit_merge voucher_split   # 指定 slug
"""

import glob
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(os.path.dirname(HERE))
SAMPLE = os.path.join(BASE, "sample_data")
WORK = os.path.join(HERE, ".replay_work")
OUT = os.path.join(HERE, "replay")
HARNESS = os.path.join(HERE, "_harness", "run_with_patches.py")
TIMEOUT = 180


def cp(src, dst_dir, name=None):
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, name or os.path.basename(src))
    shutil.copy2(src, dst)
    return dst


def cp_glob(pattern, dst_dir):
    hits = []
    for f in glob.glob(os.path.join(SAMPLE, pattern)):
        hits.append(cp(f, dst_dir))
    return hits


def cp_tree(src, dst):
    shutil.copytree(src, dst, dirs_exist_ok=True)
    return dst


# 公开的 4 个样本已改为「脚本与数据同层」；其余旧脚本仍是「脚本放子目录、数据在上一级」
SAME_FOLDER_SCRIPTS = {
    "accounting/compare_vouchers_advanced.py",
    "accounting/split_vouchers_凭证PDF拆分.py",
    "accounting/bankStatementSplitterBasedOnJournal.py",
    "integration/namelistget_feishu_v9.py",
}


def stage_script(ws, script_rel):
    if script_rel in SAME_FOLDER_SCRIPTS:
        # 新约定：脚本与数据放在同一层（脚本把「自己所在目录」当数据目录）
        dst = os.path.join(ws, os.path.basename(script_rel))
        shutil.copy2(os.path.join(BASE, script_rel.replace("/", os.sep)), dst)
        return dst
    # 旧约定：脚本放子目录，其「上一级」即工作区（数据在上一级）
    dst = os.path.join(ws, script_rel.replace("/", os.sep))
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(os.path.join(BASE, script_rel.replace("/", os.sep)), dst)
    return dst


def snapshot(root):
    snap = {}
    for r, _, files in os.walk(root):
        for f in files:
            p = os.path.join(r, f)
            try:
                snap[p] = os.path.getsize(p)
            except OSError:
                pass
    return snap


def preview_xlsx(path, max_rows=9, max_cols=12):
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    result = {"sheets": wb.sheetnames}
    for name in wb.sheetnames[:3]:
        rows = []
        for i, row in enumerate(wb[name].iter_rows(values_only=True)):
            if i >= max_rows:
                break
            rows.append(["" if v is None else (round(v, 2) if isinstance(v, float) else v) for v in row[:max_cols]])
        result.setdefault("preview", {})[name] = rows
    wb.close()
    return result


def preview_pdf(path):
    import fitz

    doc = fitz.open(path)
    pages = doc.page_count
    text = ""
    if pages:
        text = " ".join(doc[0].get_text().split())[:160]
    doc.close()
    return {"pages": pages, "first_page_text": text}


def preview_text(path, limit=500):
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            return f.read(limit)
    except OSError:
        return ""


def describe(path, ws, before):
    rel = os.path.relpath(path, ws)
    size = os.path.getsize(path)
    ext = os.path.splitext(path)[1].lower()
    info = {"path": rel.replace("\\", "/"), "bytes": size, "kind": "other", "preview": None}
    try:
        if ext == ".xlsx":
            info["kind"] = "xlsx"
            info["preview"] = preview_xlsx(path)
        elif ext == ".pdf":
            info["kind"] = "pdf"
            info["preview"] = preview_pdf(path)
        elif ext in (".txt", ".log", ".md", ".csv"):
            info["kind"] = "text"
            info["preview"] = preview_text(path)
    except Exception as e:
        info["preview_error"] = f"{type(e).__name__}: {e}"
    return info


def run_process(cmd, cwd):
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.run(
        cmd, cwd=cwd, env=env, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=TIMEOUT,
    )
    log = proc.stdout or ""
    if proc.stderr and proc.stderr.strip():
        log += "\n[stderr]\n" + proc.stderr
    return proc.returncode, log


def setup_invoice_match(ws):
    cp_glob("日常/发票/*.pdf", ws)
    cp(os.path.join(SAMPLE, "日常/全量发票查询导出结果.xlsx"), ws)


def setup_invoice_category(ws):
    cp(os.path.join(SAMPLE, "日常/processed_invoices.xlsx"), ws)


def setup_num_query(ws):
    cp(os.path.join(SAMPLE, "日常/invoicesnumlist.xlsx"), ws)
    cp(os.path.join(SAMPLE, "日常/全量发票查询导出结果.xlsx"), ws)


def setup_approval_rename(ws):
    cp_glob("月底/审批单/*.pdf", ws)


def setup_namelist_pack(ws):
    cp(os.path.join(SAMPLE, "月底/namelist.xlsx"), ws)
    cp_tree(os.path.join(SAMPLE, "月底/附件"), os.path.join(ws, "src"))
    cp_tree(os.path.join(SAMPLE, "月底/申请单"), os.path.join(ws, "app"))


def setup_voucher_split(ws):
    cp(os.path.join(SAMPLE, "归档/记账凭证-202601.pdf"), os.path.join(ws, "in"))


def setup_audit_extract(ws):
    split_dir = os.path.join(WORK, "voucher_split", "in", "记账凭证-202601_拆分结果")
    if not os.path.isdir(split_dir):
        raise RuntimeError("请先执行 voucher_split（audit_extract 依赖它的拆分产物）")
    cp_tree(split_dir, os.path.join(ws, "in", "拆分好的凭证"))
    cp(os.path.join(SAMPLE, "归档/凭证清单.xlsx"), os.path.join(ws, "in"))


def setup_voucher_compare(ws):
    cp(os.path.join(SAMPLE, "归档/凭证清单_修改前.xlsx"), os.path.join(ws, "in"))
    cp(os.path.join(SAMPLE, "归档/凭证清单_修改后.xlsx"), os.path.join(ws, "in"))


def setup_bank_split(ws):
    cp(os.path.join(SAMPLE, "归档/银行流水-202601.pdf"), ws)
    cp(os.path.join(SAMPLE, "归档/日记账202601.xlsx"), ws)


def setup_text_diff(ws):
    cp(os.path.join(SAMPLE, "归档/结算单-修改前.pdf"), os.path.join(ws, "in"))
    cp(os.path.join(SAMPLE, "归档/结算单-修改后.pdf"), os.path.join(ws, "in"))


def setup_profit(ws):
    cp_glob("辅助/利润表*.xlsx", ws)


def setup_cashflow(ws):
    cp_glob("辅助/现金流量表*.xlsx", ws)


def setup_approval_amounts(ws):
    cp_tree(os.path.join(SAMPLE, "月底/审批单"), os.path.join(ws, "审批单"))


def setup_repo_audit(ws):
    skip = {".git", "__pycache__", ".replay_work", "sample_data", "demo"}
    for r, dirs, files in os.walk(BASE):
        dirs[:] = [d for d in dirs if d not in skip]
        for f in files:
            if f.endswith(".py") and "demo" not in os.path.relpath(r, BASE).split(os.sep):
                rel = os.path.relpath(os.path.join(r, f), BASE)
                cp(os.path.join(r, f), os.path.join(ws, os.path.dirname(rel)), os.path.basename(rel))


TARGETS = [
    {
        "slug": "invoice_match",
        "title": "发票核验（主链路）",
        "script": "accounting/invoice.py",
        "stage": "daily",
        "setup": setup_invoice_match,
        "mode": "plain",
        "inputs": ["日常/发票/*.pdf ×5", "日常/全量发票查询导出结果.xlsx"],
        "note": "输出 processed_invoices.xlsx 供分类汇总消费；未匹配发票进入提取失败目录",
    },
    {
        "slug": "invoice_category",
        "title": "发票分类汇总",
        "script": "accounting/summary_by_invoice_category.py",
        "stage": "daily",
        "setup": setup_invoice_category,
        "mode": "plain",
        "inputs": ["日常/processed_invoices.xlsx（中间产物快照）"],
        "note": "读取第 24 列类别与第 16 列金额，输出分类汇总 + 合计",
    },
    {
        "slug": "invoice_enhanced",
        "title": "发票核验（增强版 134）",
        "script": "file_processing/134_invoicetofoldertxtclose.py",
        "stage": "daily",
        "setup": setup_invoice_match,
        "mode": "plain",
        "inputs": ["日常/发票/*.pdf ×5", "日常/全量发票查询导出结果.xlsx"],
        "note": "最新版：序号 + 5 类关键词 + 提取结果目录结构",
    },
    {
        "slug": "invoice_number_query",
        "title": "发票号码批量回捞",
        "script": "accounting/invoice_data_extractor.py",
        "stage": "daily",
        "setup": setup_num_query,
        "mode": "plain",
        "inputs": ["日常/invoicesnumlist.xlsx", "日常/全量发票查询导出结果.xlsx"],
        "note": "按号码长度路由 8 位/20 位列，回填完整发票字段",
    },
    {
        "slug": "approval_rename",
        "title": "审批单标准化重命名",
        "script": "file_processing/pdf_processor_for_approval.py",
        "stage": "monthend",
        "setup": setup_approval_rename,
        "mode": "plain",
        "inputs": ["月底/审批单/*.pdf ×5（无固定文件名）"],
        "note": "输出 单据类型-申请编号-单据编号-申请人-审批状态-序号 标准命名",
    },
    {
        "slug": "namelist_pack",
        "title": "月底资料包匹配整合",
        "script": "integration/namelistget_feishu_v9.py",
        "stage": "monthend",
        "setup": setup_namelist_pack,
        "mode": "harness",
        "dialogs": {
            "dirs": ["{ws}/app", "{ws}/src"],
            "stdin": ["", ""],
        },
        "inputs": ["月底/namelist.xlsx", "月底/附件/（报销单+附件）", "月底/申请单/（关联申请单）"],
        "note": "复制归集 → 关联申请单 → 图片转 PDF → 合并资料包 → 生成对应表与错误报告",
    },
    {
        "slug": "voucher_split",
        "title": "凭证 PDF 拆分",
        "script": "accounting/split_vouchers_凭证PDF拆分.py",
        "stage": "archive",
        "setup": setup_voucher_split,
        "mode": "harness",
        "dialogs": {"files": ["{ws}/in/记账凭证-202601.pdf"]},
        "inputs": ["归档/记账凭证-202601.pdf（3 页 2 张凭证）"],
        "note": "按 年月+基础凭证号 分组，输出 {年月}-{凭证号}-{N}页.pdf",
    },
    {
        "slug": "audit_extract",
        "title": "审计清单提取",
        "script": "accounting/清单提取拆分好的凭证.py",
        "stage": "archive",
        "setup": setup_audit_extract,
        "mode": "harness",
        "dialogs": {"files": ["{ws}/in/凭证清单.xlsx"], "dirs": ["{ws}/in/拆分好的凭证"]},
        "inputs": ["归档/凭证清单.xlsx", "voucher_split 的拆分产物"],
        "note": "文件名前缀 {YYYY-MM}-{凭证类别}-{凭证号zfill(3)} 匹配，输出提取报告",
    },
    {
        "slug": "voucher_compare",
        "title": "凭证深度比对（Jaccard）",
        "script": "accounting/compare_vouchers_advanced.py",
        "stage": "archive",
        "setup": setup_voucher_compare,
        "mode": "harness",
        "dialogs": {"files": ["{ws}/in/凭证清单_修改前.xlsx", "{ws}/in/凭证清单_修改后.xlsx"]},
        "inputs": ["归档/凭证清单_修改前.xlsx", "归档/凭证清单_修改后.xlsx"],
        "note": "覆盖 内容修改/换号/新增/删除 四类差异，附单据数变化标红",
    },
    {
        "slug": "voucher_list_compare",
        "title": "新旧凭证列表整组比对",
        "script": "accounting/新旧凭证列表比较及差异输出.py",
        "stage": "archive",
        "setup": setup_voucher_compare,
        "mode": "harness",
        "dialogs": {
            "stdin": ["{ws}/in/凭证清单_修改前.xlsx", "{ws}/in/凭证清单_修改后.xlsx"],
        },
        "inputs": ["归档/凭证清单_修改前.xlsx", "归档/凭证清单_修改后.xlsx"],
        "note": "整组匹配思路：凭证号整组一致才算匹配，差异行红字标注",
    },
    {
        "slug": "bank_split",
        "title": "银行流水按凭证拆分",
        "script": "accounting/bankStatementSplitterBasedOnJournal.py",
        "stage": "archive",
        "setup": setup_bank_split,
        "mode": "plain",
        "inputs": ["归档/银行流水-202601.pdf", "归档/日记账202601.xlsx"],
        "note": "根目录必须恰好 1 个 PDF + 1 个日记账；行数与页数差异 >2 直接拒绝",
    },
    {
        "slug": "pdf_text_diff",
        "title": "PDF 文本逐页比对",
        "script": "accounting/文字pdf对比脚本.py",
        "stage": "archive",
        "setup": setup_text_diff,
        "mode": "harness",
        "dialogs": {"files": ["{ws}/in/结算单-修改前.pdf", "{ws}/in/结算单-修改后.pdf"]},
        "inputs": ["归档/结算单-修改前.pdf", "归档/结算单-修改后.pdf"],
        "note": "逐页提取文本，无差异/有差异 + unified diff",
    },
    {
        "slug": "profit_merge",
        "title": "利润表多月合并",
        "script": "accounting/利润表生成.py",
        "stage": "aux",
        "setup": setup_profit,
        "mode": "plain",
        "inputs": ["辅助/利润表202601.xlsx", "辅助/利润表202602.xlsx"],
        "note": "按利润表项目外连接，行次取样本文件，输出带时间戳合并表",
    },
    {
        "slug": "cashflow_merge",
        "title": "现金流量表多月合并",
        "script": "accounting/合并现金流量表.py",
        "stage": "aux",
        "setup": setup_cashflow,
        "mode": "plain",
        "inputs": ["辅助/现金流量表202601.xlsx", "辅助/现金流量表202602.xlsx"],
        "note": "先解除合并单元格再提取，输出列宽已设置的合并表",
    },
    {
        "slug": "approval_amounts",
        "title": "审批单金额/日期提取",
        "script": "file_processing/processedfilesinfoalltxt.py",
        "stage": "aux",
        "setup": setup_approval_amounts,
        "mode": "plain",
        "inputs": ["月底/审批单/*.pdf ×5"],
        "note": "README 标注未完成；可验证关键词扩列与日期提取逻辑",
    },
    {
        "slug": "repo_audit",
        "title": "仓库重复文件自检",
        "script": "accounting/analyze_duplicates.py",
        "stage": "aux",
        "setup": setup_repo_audit,
        "mode": "plain",
        "inputs": ["仓库全部 .py 源码副本（59 个）"],
        "note": "按核心名分组 + difflib 相似度，输出 保留/审查/删除 建议报告",
    },
]

SIMULATED = [
    {
        "slug": "clipboard_summary",
        "title": "飞书审批摘要（剪贴板）",
        "script": "integration/get_filename_from_feishu.py",
        "stage": "daily",
        "reason": "依赖人工在飞书页面全选复制 + GUI 热键，无法无头执行；本条目为按代码逻辑生成的示意回放",
        "inputs": ["剪贴板：飞书审批详情页文本（人工复制）"],
        "log": (
            "[监控] 剪贴板轮询中 ... (800ms)\n"
            "[命中] 剪贴板内容变更，长度 1832 > 50，含关键词：审批详情 / 单据编号 / 申请人\n"
            "[解析] 锚点「审批详情」向前回溯 1000 / 向后 2000 字符\n"
            "[提取] 单据类型 = 差旅报销单\n"
            "[提取] 申请编号 = 202603010001\n"
            "[提取] 单据编号 = KZCLBX202603010001\n"
            "[提取] 申请人   = 张三（部门词表命中：研发中心）\n"
            "--------------------------------------------------\n"
            "标准凭证摘要（已复制到剪贴板）：\n"
            "差旅报销单-202603010001-KZCLBX202603010001-张三\n"
            "--------------------------------------------------\n"
            "[提示] 粘贴到畅捷通凭证草稿箱摘要栏即可完成录入"
        ),
    },
    {
        "slug": "ofd_convert",
        "title": "OFD 转 PDF",
        "script": "file_processing/convert_all.py",
        "stage": "daily",
        "reason": "OFD 需真实电子发票样本；本机无 OFD 文件，日志为按代码逻辑生成的示意回放",
        "inputs": ["OFD 电子发票/文档（文件夹）"],
        "log": (
            "--- PyMuPDF 版本: PyMuPDF 1.26.x ---\n"
            "正在扫描文件夹: <工作目录>\n"
            "找到了 3 个 OFD 文件，开始转换...\n\n"
            "--> 正在处理: 26332000000012345601.ofd\n"
            "    转换成功! 已保存为: 26332000000012345601.pdf\n\n"
            "--> 正在处理: 26332000000012345602.ofd\n"
            "    跳过! 文件 '26332000000012345602.pdf' 已存在。\n\n"
            "--> 正在处理: 26332000000012345603.ofd\n"
            "    转换成功! 已保存为: 26332000000012345603.pdf\n\n"
            "请按 Enter 键退出..."
        ),
    },
    {
        "slug": "web_downloader",
        "title": "网银/网页文件下载（实验）",
        "script": "integration/web_file_downloader_v2.py",
        "stage": "experiment",
        "reason": "Selenium 需要真实网页与 Edge 会话，未接入生产链路；日志为示意",
        "inputs": ["人工输入的 https URL"],
        "log": (
            "请输入安全的网页URL（https://开头）: https://example-bank.example.com/statement\n"
            "[Selenium] 启动 Edge 驱动（禁用 SSL 校验，仅测试环境）\n"
            "[Selenium] 打开页面完成，定位「下载」元素 ×3\n"
            "  - 点击第 1 个「下载」-> 已保存 银行流水_202601.pdf\n"
            "  - 点击第 2 个「下载」-> 已保存 银行流水_202602.pdf\n"
            "  - 点击第 3 个「下载」-> 已保存 回单明细.xlsx\n"
            "下载目录：~/Downloads/Secure_Downloads\n"
            "[说明] 实验性功能：规划用于替代畅捷通/网银的人工下载环节"
        ),
    },
    {
        "slug": "intercompany",
        "title": "关联方往来清理（半成品）",
        "script": "accounting/IntercompanyOutstandingCleanupTool.py",
        "stage": "aux",
        "reason": "some_condition_met 未定义、依赖人工填写 column_info.xlsx，README 标注开发中；日志为示意",
        "inputs": ["全部科目余额表.xlsx", "科目辅助明细账.xlsx", "column_info.xlsx（列映射配置）"],
        "log": (
            "已在 D:\\your_data\\column_info.xlsx 创建 column_info.xlsx，请填写文件名称/列名/列号...\n"
            "请按回车键继续 ...\n"
            "请输入全部科目余额表xlsx文件的路径：<已输入>\n"
            "请输入科目辅助明细账xlsx文件的路径：<已输入>\n"
            "[读取] 余额表：header=3，按 column_info 定位「期末余额」列\n"
            "[读取] 明细账：定位 摘要/借方/贷方/项目/客户/供应商/部门/员工 列\n"
            "[中断] 处理到「科目余额是否为0」判定时命中占位逻辑 some_condition_met\n"
            "[说明] 该工具开发中：列映射零配置思路已跑通，判定逻辑待补全"
        ),
    },
]


def sanitize_obj(obj):
    if isinstance(obj, str):
        return obj.replace(WORK, "<工作区>").replace(BASE, "<仓库>")
    if isinstance(obj, list):
        return [sanitize_obj(x) for x in obj]
    if isinstance(obj, dict):
        return {k: sanitize_obj(v) for k, v in obj.items()}
    return obj


def execute_target(target):
    slug = target["slug"]
    ws = os.path.join(WORK, slug)
    if os.path.exists(ws):
        shutil.rmtree(ws)
    os.makedirs(ws)

    started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    record = {
        "slug": slug,
        "title": target["title"],
        "script": target["script"],
        "stage": target["stage"],
        "fidelity": "real",
        "mode": target["mode"],
        "note": target["note"],
        "inputs": target["inputs"],
        "started_at": started,
        "dialogs": [],
        "outputs": [],
        "log": "",
        "exit_code": None,
        "duration_s": None,
    }

    try:
        target["setup"](ws)
        script_abs = stage_script(ws, target["script"])
        before = snapshot(ws)
        t0 = time.time()

        if target["mode"] == "plain":
            cmd = [sys.executable, script_abs]
        else:
            dialogs = target.get("dialogs", {})
            cfg = {
                "script": script_abs,
                "cwd": ws,
                "files": [p.format(ws=ws) for p in dialogs.get("files", [])],
                "dirs": [p.format(ws=ws) for p in dialogs.get("dirs", [])],
                "stdin": [p.format(ws=ws) for p in dialogs.get("stdin", [])],
                "events": os.path.join(ws, ".harness_events.json"),
            }
            cfg_path = os.path.join(ws, ".harness_config.json")
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            cmd = [sys.executable, HARNESS, cfg_path]

        code, log = run_process(cmd, ws)
        record["exit_code"] = code
        record["log"] = log
        record["duration_s"] = round(time.time() - t0, 2)

        events_path = os.path.join(ws, ".harness_events.json")
        if os.path.exists(events_path):
            with open(events_path, encoding="utf-8") as f:
                record["dialogs"] = json.load(f)

        after = snapshot(ws)
        new_files = [p for p in after if p not in before and not os.path.basename(p).startswith(".harness")]
        for path in sorted(new_files):
            record["outputs"].append(describe(path, ws, before))

    except Exception as e:
        import traceback

        record["exit_code"] = -1
        record["log"] = (record.get("log") or "") + "\n[replay-error] " + "".join(
            traceback.format_exception_only(type(e), e)
        )
        traceback.print_exc()

    return record


def build_simulated(spec):
    return {
        "slug": spec["slug"],
        "title": spec["title"],
        "script": spec["script"],
        "stage": spec["stage"],
        "fidelity": "simulated",
        "mode": "simulated",
        "note": spec["reason"],
        "inputs": spec["inputs"],
        "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "dialogs": [],
        "outputs": [],
        "log": spec["log"],
        "exit_code": None,
        "duration_s": None,
    }


def main():
    wanted = set(sys.argv[1:])
    os.makedirs(OUT, exist_ok=True)
    os.makedirs(WORK, exist_ok=True)

    results = []
    for target in TARGETS:
        if wanted and target["slug"] not in wanted:
            continue
        print(f"\n>>> 回放 {target['slug']} - {target['title']}")
        record = sanitize_obj(execute_target(target))
        ok = record["exit_code"] == 0
        print(f"    退出码={record['exit_code']}  新增产物={len(record['outputs'])}  {'OK' if ok else 'FAILED'}")
        with open(os.path.join(OUT, f"{record['slug']}.json"), "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
        results.append(record)

    for spec in SIMULATED:
        if wanted and spec["slug"] not in wanted:
            continue
        record = sanitize_obj(build_simulated(spec))
        with open(os.path.join(OUT, f"{record['slug']}.json"), "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
        results.append(record)
        print(f"\n>>> 示意条目 {record['slug']} - {record['title']}（simulated）")

    real = [r for r in results if r["fidelity"] == "real"]
    failed = [r for r in real if r["exit_code"] != 0]
    print("\n" + "=" * 64)
    print(f"回放完成：实测 {len(real)} 条，其中失败 {len(failed)} 条；示意 {len(results) - len(real)} 条")
    if failed:
        print("失败清单：", ", ".join(r["slug"] for r in failed))
    print("输出目录：", OUT)


if __name__ == "__main__":
    main()
