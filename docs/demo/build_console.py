"""构建单文件控制台与功能地图文档。

输入：
  docs/function_map.json     功能关联档案（人工整理）
  docs/demo/replay/*.json    回放采集结果（run_replay.py 产出）
输出：
  docs/demo/console.html     单文件静态控制台（零依赖，双击即开）
  docs/FUNCTION_MAP.md       人类可读功能地图（由档案生成，保证与页面一致）

用法：python docs/demo/build_console.py
"""

import json
import os
import sys
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(os.path.dirname(HERE))
FUNCTION_MAP = os.path.join(BASE, "docs", "function_map.json")
TEMPLATE = os.path.join(HERE, "console.template.html")
CONSOLE_OUT = os.path.join(HERE, "console.html")
MD_OUT = os.path.join(BASE, "docs", "FUNCTION_MAP.md")
REPLAY_DIR = os.path.join(HERE, "replay")

SCRIPT_REPLAY = {
    "accounting/invoice.py": "invoice_match",
    "accounting/summary_by_invoice_category.py": "invoice_category",
    "file_processing/134_invoicetofoldertxtclose.py": "invoice_enhanced",
    "accounting/invoice_data_extractor.py": "invoice_number_query",
    "file_processing/pdf_processor_for_approval.py": "approval_rename",
    "integration/namelistget_feishu_v9.py": "namelist_pack",
    "accounting/split_vouchers_凭证PDF拆分.py": "voucher_split",
    "accounting/清单提取拆分好的凭证.py": "audit_extract",
    "accounting/compare_vouchers_advanced.py": "voucher_compare",
    "accounting/新旧凭证列表比较及差异输出.py": "voucher_list_compare",
    "accounting/bankStatementSplitterBasedOnJournal.py": "bank_split",
    "accounting/文字pdf对比脚本.py": "pdf_text_diff",
    "accounting/利润表生成.py": "profit_merge",
    "accounting/合并现金流量表.py": "cashflow_merge",
    "file_processing/processedfilesinfoalltxt.py": "approval_amounts",
    "accounting/analyze_duplicates.py": "repo_audit",
    "integration/get_filename_from_feishu.py": "clipboard_summary",
    "file_processing/convert_all.py": "ofd_convert",
    "integration/web_file_downloader_v2.py": "web_downloader",
    "accounting/IntercompanyOutstandingCleanupTool.py": "intercompany",
}

MAP_DEFAULT = [
    "integration/get_filename_from_feishu.py",
    "file_processing/134_invoicetofoldertxtclose.py",
    "accounting/invoice.py",
    "accounting/summary_by_invoice_category.py",
    "accounting/invoice_data_extractor.py",
    "file_processing/convert_all.py",
    "file_processing/pdf_processor_for_approval.py",
    "integration/namelistget_feishu_v9.py",
    "integration/get_dingding_form.py",
    "202602/飞书/get_dingding_form_keyword_line_1_and_2_with_name_from_line2_with_date_in_result.py",
    "accounting/compare_vouchers_advanced.py",
    "accounting/新旧凭证列表比较及差异输出.py",
    "accounting/清单提取拆分好的凭证.py",
    "accounting/split_vouchers_凭证PDF拆分.py",
    "accounting/bankStatementSplitterBasedOnJournal.py",
    "accounting/文字pdf对比脚本.py",
    "accounting/利润表生成.py",
    "accounting/合并现金流量表.py",
    "accounting/analyze_duplicates.py",
    "accounting/IntercompanyOutstandingCleanupTool.py",
    "file_processing/processedfilesinfoalltxt.py",
    "integration/web_file_downloader_v2.py",
]


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_replays():
    replays = {}
    if os.path.isdir(REPLAY_DIR):
        for name in sorted(os.listdir(REPLAY_DIR)):
            if name.endswith(".json"):
                data = load_json(os.path.join(REPLAY_DIR, name))
                replays[data["slug"]] = data
    return replays


def build_audit(replays):
    split = replays.get("voucher_split", {})
    extract = replays.get("audit_extract", {})

    split_files = [o for o in split.get("outputs", []) if o["kind"] == "pdf"]
    extract_files = [o for o in extract.get("outputs", []) if o["path"].startswith("in/拆分好的凭证")]

    report = next((o for o in extract.get("outputs", []) if o["path"].endswith("提取结果报告.xlsx")), None)
    report_preview = (report or {}).get("preview") or {}
    sheet_rows = []
    if report_preview.get("preview"):
        sheet_rows = list(report_preview["preview"].values())[0]

    status_by_voucher, filename_by_voucher = {}, {}
    if sheet_rows:
        header = [str(c) for c in sheet_rows[0]]
        try:
            num_idx = header.index("凭证号")
            status_idx = header.index("提取状态")
            file_idx = header.index("提取到的文件名") if "提取到的文件名" in header else None
            for row in sheet_rows[1:]:
                key = "记-" + str(int(row[num_idx])).zfill(3)
                status_by_voucher[key] = row[status_idx]
                if file_idx is not None:
                    filename_by_voucher[key] = row[file_idx]
        except (ValueError, TypeError, IndexError):
            pass

    vouchers = []
    for o in split_files:
        name = os.path.basename(o["path"])
        parts = name.replace(".pdf", "").split("-")
        if len(parts) < 4:
            continue
        voucher_id = parts[2] + "-" + parts[3]
        vouchers.append({
            "id": voucher_id,
            "pages": (o.get("preview") or {}).get("pages", 0),
            "split_file": name,
            "split_bytes": o["bytes"],
            "extract_status": status_by_voucher.get(voucher_id, "成功提取"),
            "extract_files": filename_by_voucher.get(voucher_id, name),
            "extract_report": "提取结果报告.xlsx",
        })

    report_name = os.path.basename(report["path"]) if report else "提取结果报告.xlsx"
    picked = [o for o in extract_files if o["kind"] == "pdf"]

    steps = [
        {
            "title": "① 整本记账凭证（畅捷通导出）",
            "detail": "月底结账后导出的整本凭证 PDF，按凭证号连续排列（样例 3 页）",
            "files": [{"name": "记账凭证-202601.pdf", "meta": "3 页 · 外部来源"}],
        },
        {
            "title": "② 按 年月+凭证号 拆分",
            "detail": "split_vouchers_凭证PDF拆分.py 逐页解析「日期 / 凭证号」，按凭证归组输出独立文件",
            "files": [{"name": v["split_file"], "meta": str(v["pages"]) + " 页 · " + str(round(v["split_bytes"] / 1024)) + " KB", "voucher": v["id"]} for v in vouchers],
        },
        {
            "title": "③ 凭证清单（决定提取范围）",
            "detail": "审计清单 Excel：凭证类别 / 凭证号 / 凭证日期，作为提取白名单",
            "files": [{"name": "凭证清单.xlsx", "meta": "样例 2 张凭证"}],
        },
        {
            "title": "④ 按清单提取审计凭证",
            "detail": "清单提取拆分好的凭证.py 用文件名前缀 {YYYY-MM}-{类别}-{号zfill(3)} 匹配并复制",
            "files": [{"name": o["path"].split("/")[-1], "meta": str(o["bytes"]) + " B", "voucher": v["id"] if (v := next((x for x in vouchers if os.path.basename(o["path"]).find(x["split_file"].replace(".pdf", "")) >= 0), None)) else None} for o in picked] + [{"name": report_name, "meta": "提取结果报告"}],
        },
        {
            "title": "⑤ 归档",
            "detail": "提取结果连同报告归档进「凭证提取结果_{时间戳}/」，任意凭证 5 分钟可完整溯源",
            "files": [],
        },
    ]
    return {
        "vouchers": vouchers,
        "steps": steps,
        "note": "证据链基于 sample_data 生成假数据、由 run_replay.py 真实执行采集；提取状态与文件名直接来自「提取结果报告.xlsx」的回放快照。",
    }


def sanitize_replays(replays):
    work = os.path.join(HERE, ".replay_work")
    for rp in replays.values():
        rp["log"] = rp.get("log", "").replace(work, "<工作区>").replace(BASE, "<仓库>")
        for ev in rp.get("dialogs", []):
            for key in ("value", "path"):
                if ev.get(key):
                    ev[key] = ev[key].replace(work, "<工作区>").replace(BASE, "<仓库>")
    return replays


def build_data():
    data = load_json(FUNCTION_MAP)
    replays = sanitize_replays(load_replays())
    data["replays"] = replays
    data["scriptReplay"] = SCRIPT_REPLAY
    data["mapDefault"] = MAP_DEFAULT
    data["audit"] = build_audit(replays)
    return data


def render_console(data):
    with open(TEMPLATE, encoding="utf-8") as f:
        template = f.read()
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    marker = "/*__DATA__*/ null"
    if marker not in template:
        raise SystemExit("模板缺少注入标记 /*__DATA__*/ null")
    html = template.replace(marker, payload)
    with open(CONSOLE_OUT, "w", encoding="utf-8") as f:
        f.write(html)
    return os.path.getsize(CONSOLE_OUT)


STATUS_LABEL = {
    "core": "核心脚本", "active": "在用", "wip": "半成品",
    "broken": "无法运行", "legacy": "已被替代", "dead": "已废弃", "experiment": "实验性",
}


def render_markdown(data):
    stages = data["stages"]
    scripts = data["scripts"]
    by_stage = {s["id"]: [] for s in stages}
    for s in scripts:
        by_stage.setdefault(s["stage"], []).append(s)

    replay_slugs = data.get("scriptReplay", {})
    replays = data.get("replays", {})

    lines = []
    lines.append("# 功能关联地图（FUNCTION_MAP）")
    lines.append("")
    lines.append("> 本文件由 `docs/demo/build_console.py` 从 [`docs/function_map.json`](function_map.json) 自动生成，与 `docs/demo/console.html` 数据同源。")
    lines.append("")
    lines.append(f"> 生成时间：{data['meta']['generated_at']} ｜ 脚本总数：**{data['meta']['total_scripts']}** ｜ 核心脚本（未公开）：**{data['meta']['tracked_scripts']}**")
    lines.append("")
    lines.append(f"> 公开范围：{data['meta'].get('public_scope', '')}")
    lines.append("")
    lines.append("## 关系模型")
    lines.append("")
    lines.append(f"{data['meta']['relation_model']}。")
    lines.append("")
    lines.append("## 阶段分布")
    lines.append("")
    lines.append("| 阶段 | 脚本数 | 说明 |")
    lines.append("|------|--------|------|")
    for st in stages:
        lines.append(f"| {st['name']} | {len(by_stage.get(st['id'], []))} | {st['desc']} |")
    lines.append("")

    for st in stages:
        items = sorted(
            by_stage.get(st["id"], []),
            key=lambda s: (0 if s.get("core_chain") else 1, list(STATUS_LABEL).index(s["status"]) if s["status"] in STATUS_LABEL else 9),
        )
        if not items:
            continue
        lines.append(f"## {st['name']}（{len(items)}）")
        lines.append("")
        for s in items:
            tags = [STATUS_LABEL.get(s["status"], s["status"])]
            if s.get("core_chain"):
                tags.append("核心链路")
            slug = replay_slugs.get(s["id"])
            if slug and slug in replays:
                tags.append("实测回放" if replays[slug]["fidelity"] == "real" else "示意回放")
            lines.append(f"### {s['name']}")
            lines.append("")
            lines.append(f"`{s['id']}` ｜ {' · '.join(tags)}")
            lines.append("")
            lines.append(f"{s['purpose']}。")
            lines.append("")
            if s.get("inputs"):
                lines.append(f"- 输入：{'；'.join(s['inputs'])}")
            if s.get("outputs"):
                lines.append(f"- 输出：{'；'.join(s['outputs'])}")
            if s.get("interactive"):
                lines.append(f"- 交互：{'、'.join(s['interactive'])}")
            if s.get("deps"):
                lines.append(f"- 依赖：{', '.join(s['deps'])}")
            if s.get("supersedes"):
                lines.append(f"- 替代：{s['supersedes']}")
            if s.get("superseded_by"):
                lines.append(f"- 被替代：{s['superseded_by']}")
            for n in s.get("notes", []):
                lines.append(f"- 要点：{n}")
            lines.append("")

    lines.append("## 版本谱系")
    lines.append("")
    for lin in data.get("lineages", []):
        lines.append(f"### {lin['family']}")
        lines.append("")
        lines.append(f"{lin['desc']}。")
        lines.append("")
        for m in lin["members"]:
            mark = " ← 当前" if m == lin["current"].split(" + ")[0] else ""
            lines.append(f"- `{m}`{mark}")
        if " + " in lin["current"]:
            lines.append(f"- 当前（另一条路线）：`{lin['current'].split(' + ')[1]}`")
        lines.append("")

    lines.append("## 效能指标（示意口径）")
    lines.append("")
    lines.append("| 指标 | 优化前 | 优化后 | 变化 |")
    lines.append("|------|--------|--------|------|")
    for k in data.get("kpis", []):
        lines.append(f"| {k['name']} | {k['before']} | {k['after']} | {k['improve']} |")
    lines.append("")
    lines.append(f"> {data.get('kpi_note', '')}")
    lines.append("")
    lines.append("## 回放清单")
    lines.append("")
    lines.append("| 回放 | 脚本 | 类型 | 说明 |")
    lines.append("|------|------|------|------|")
    for slug, rp in replays.items():
        kind = "实测" if rp["fidelity"] == "real" else "示意"
        lines.append(f"| {rp['title']} | `{rp['script']}` | {kind} | {rp['note']} |")
    lines.append("")
    lines.append("> 实测 = 在隔离工作区用 `sample_data/` 假数据真实执行采集；示意 = 依赖 GUI/外部环境无法无头执行，按代码逻辑生成，已在页面中标注。")
    lines.append("")

    with open(MD_OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return len(lines)


def main():
    data = build_data()
    size = render_console(data)
    lines = render_markdown(data)
    real = sum(1 for r in data["replays"].values() if r["fidelity"] == "real")
    print(f"console.html 生成完成: {CONSOLE_OUT} ({size/1024:.0f} KB)")
    print(f"FUNCTION_MAP.md 生成完成: {MD_OUT} ({lines} 行)")
    print(f"回放注入: {len(data['replays'])} 条（实测 {real} / 示意 {len(data['replays']) - real}）")


if __name__ == "__main__":
    main()
