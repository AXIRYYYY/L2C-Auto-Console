"""生成 sample_data 假数据（零真实业务信息）。

数据结构按脚本解析规则定制：
- 发票 PDF 文本可通过 pdfplumber 提取，号码/金额与「全量发票查询导出结果.xlsx」对得上
- 审批单 PDF 含 申请人/申请编号/单据编号/审批状态 等字段，可被 pdf_processor_for_approval.py 重命名
- 记账凭证 PDF 含「日期: 」「凭证号: 记-001-1/2」文本，可被 split_vouchers 拆分
- 银行流水 PDF 页数与日记账标签行数对应，可被 bankStatementSplitterBasedOnJournal.py 拆分

用法：python sample_data/generate_sample_data.py
"""

import os
import shutil
import sys

import fitz
from openpyxl import Workbook

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

BASE = os.path.dirname(os.path.abspath(__file__))

CJK = "china-s"

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\Deng.ttf",
    r"C:\Windows\Fonts\meiryo.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
]


def pick_font():
    for cand in FONT_CANDIDATES:
        if os.path.exists(cand):
            return cand
    return None


FONT_FILE = pick_font()


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def wrap_line(line, width=38):
    return [line[i:i + width] for i in range(0, len(line), width)] or [line]


def make_pdf(path, page_lines, font_size=11):
    ensure_dir(os.path.dirname(path))
    doc = fitz.open()
    for lines in page_lines:
        page = doc.new_page()
        y = 72
        for line in lines:
            for segment in wrap_line(line):
                if FONT_FILE:
                    page.insert_text((72, y), segment, fontsize=font_size, fontname="emb", fontfile=FONT_FILE)
                else:
                    page.insert_text((72, y), segment, fontsize=font_size, fontname=CJK)
                y += 24
    try:
        doc.subset_fonts()
    except Exception:
        pass
    doc.save(path)
    doc.close()


def make_xlsx(path, sheets):
    ensure_dir(os.path.dirname(path))
    wb = Workbook()
    first = True
    for name, rows, col_widths in sheets:
        ws = wb.active if first else wb.create_sheet()
        ws.title = name
        first = False
        for row in rows:
            ws.append(row)
        if col_widths:
            for idx, width in enumerate(col_widths, start=1):
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = width
    wb.save(path)


INVOICE_MASTER_HEADER = [
    "序号", "发票号码", "数电票号码", "数电发票号码", "开票日期", "购买方名称",
    "销售方名称", "项目名称", "税率", "税额", "不含税金额", "价税合计金额", "开票人", "备注",
]

INVOICES = [
    {
        "file": "日常/发票/dzfp_26332000000012345601_示例酒店_20260205103000.pdf",
        "number": "26332000000012345601",
        "kind": "数电",
        "date": "2026-02-05",
        "seller": "杭州示例酒店管理有限公司",
        "item": "*住宿服务*住宿费",
        "category": "住宿服务",
        "amount": 1180.00,
        "tax": 66.79,
    },
    {
        "file": "日常/发票/dzfp_26332000000012345602_示例餐饮_20260206120000.pdf",
        "number": "26332000000012345602",
        "kind": "数电",
        "date": "2026-02-06",
        "seller": "杭州样例餐饮有限公司",
        "item": "*餐饮服务*工作餐",
        "category": "餐饮服务",
        "amount": 256.00,
        "tax": 14.49,
    },
    {
        "file": "日常/发票/dzfp_26332000000012345603_示例加油_20260207180000.pdf",
        "number": "26332000000012345603",
        "kind": "数电",
        "date": "2026-02-07",
        "seller": "示例石油销售有限公司",
        "item": "*汽油*车用汽油",
        "category": "汽油",
        "amount": 300.00,
        "tax": 34.51,
    },
    {
        "file": "日常/发票/发票_12345678_通行费.pdf",
        "number": "12345678",
        "code": "87654321",
        "kind": "纸质",
        "date": "2026-02-08",
        "seller": "示例高速公路管理有限公司",
        "item": "通行费",
        "category": "通行费",
        "amount": 15.00,
        "tax": 0.44,
    },
    {
        "file": "日常/发票/dzfp_99999999999999999999_未匹配_20260208090000.pdf",
        "number": "99999999999999999999",
        "kind": "数电",
        "date": "2026-02-08",
        "seller": "示例未知供应商",
        "item": "*办公用品*打印纸",
        "category": "",
        "amount": 88.00,
        "tax": 10.12,
    },
]

APPROVALS = [
    {
        "file": "月底/审批单/审批单01.pdf",
        "type": "差旅报销单",
        "bill": "202603010001",
        "doc": "KZCLBX202603010001",
        "applicant": "张三",
        "dept": "研发中心",
        "status": "已通过",
    },
    {
        "file": "月底/审批单/审批单02.pdf",
        "type": "差旅报销单",
        "bill": "202603020002",
        "doc": "KZCLBX202603020002",
        "applicant": "张三",
        "dept": "研发中心",
        "status": "已通过",
    },
    {
        "file": "月底/审批单/审批单03.pdf",
        "type": "费用报销单",
        "bill": "202603030003",
        "doc": "KZFYBX202603030003",
        "applicant": "李四",
        "dept": "销售管理部",
        "status": "审批中",
    },
    {
        "file": "月底/审批单/审批单04.pdf",
        "type": "采购申请单",
        "bill": "202603040004",
        "doc": "KZCGSQ202603040004",
        "applicant": "王五",
        "dept": "行政部",
        "status": "已通过",
    },
    {
        "file": "月底/审批单/审批单05.pdf",
        "type": "付款申请",
        "bill": "202603050005",
        "doc": "KZFKSQ202603050005",
        "applicant": "赵六",
        "dept": "财务部",
        "status": "已驳回",
    },
]


def gen_invoices():
    for inv in INVOICES:
        if inv["kind"] == "数电":
            lines = [
                "电子发票（普通发票）",
                f"发票号码：{inv['number']}",
                f"开票日期：{inv['date']}",
                f"购买方名称：杭州示例科技有限公司",
                f"销售方名称：{inv['seller']}",
                f"{inv['item']} 1 {inv['amount']:.2f} {inv['tax']:.2f}",
                f"价税合计（小写）￥{inv['amount']:.2f}",
            ]
        else:
            lines = [
                "电子发票（普通发票）",
                f"发票代码:{inv['code']}",
                f"发票号码：{inv['number']}",
                f"开票日期：{inv['date']}",
                f"购买方名称：杭州示例科技有限公司",
                f"销售方名称：{inv['seller']}",
                f"{inv['item']} 1 {inv['amount']:.2f} {inv['tax']:.2f}",
                f"价税合计（小写）￥{inv['amount']:.2f}",
            ]
        make_pdf(os.path.join(BASE, inv["file"]), [lines])


def gen_invoice_master():
    rows = [INVOICE_MASTER_HEADER]
    for idx, inv in enumerate(INVOICES, start=1):
        if inv["category"] == "":
            continue
        if inv["kind"] == "数电":
            rows.append([
                idx, "", inv["number"], inv["number"], inv["date"], "杭州示例科技有限公司",
                inv["seller"], inv["item"], "6%", inv["tax"], round(inv["amount"] - inv["tax"], 2),
                inv["amount"], "示例开票员", "",
            ])
        else:
            rows.append([
                idx, inv["number"], "", "", inv["date"], "杭州示例科技有限公司",
                inv["seller"], inv["item"], "3%", inv["tax"], round(inv["amount"] - inv["tax"], 2),
                inv["amount"], "示例开票员", "",
            ])
    make_xlsx(
        os.path.join(BASE, "日常/全量发票查询导出结果.xlsx"),
        [("发票基础信息", rows, [6, 14, 24, 24, 12, 24, 24, 18, 8, 10, 12, 14, 10, 10])],
    )


def gen_invoices_numlist():
    rows = [
        ["发票号码清单"],
        ["26332000000012345601"],
        ["12345678"],
        ["26332000000012345603"],
    ]
    make_xlsx(os.path.join(BASE, "日常/invoicesnumlist.xlsx"), [("Sheet1", rows, [24])])


def gen_processed_invoices_snapshot():
    header = [
        "序号", "文件名称", "发票号码（提取）", "发票代码（提取）",
        INVOICE_MASTER_HEADER[0], INVOICE_MASTER_HEADER[1], INVOICE_MASTER_HEADER[2],
        INVOICE_MASTER_HEADER[3], INVOICE_MASTER_HEADER[4], INVOICE_MASTER_HEADER[5],
        INVOICE_MASTER_HEADER[6], INVOICE_MASTER_HEADER[7], INVOICE_MASTER_HEADER[8],
        INVOICE_MASTER_HEADER[9], INVOICE_MASTER_HEADER[10], INVOICE_MASTER_HEADER[11],
        INVOICE_MASTER_HEADER[12], INVOICE_MASTER_HEADER[13],
        "", "", "", "", "", "发票类别",
    ]
    rows = [header]
    for idx, inv in enumerate(INVOICES[:4], start=1):
        if inv["kind"] == "数电":
            rows.append([
                idx, "日常/发票/" + os.path.basename(inv["file"]), inv["number"], "",
                idx, "", inv["number"], inv["number"], inv["date"], "杭州示例科技有限公司",
                inv["seller"], inv["item"], "6%", inv["tax"], round(inv["amount"] - inv["tax"], 2),
                inv["amount"], "示例开票员", "",
            ] + [None] * 5 + [inv["category"]])
        else:
            rows.append([
                idx, "日常/发票/" + os.path.basename(inv["file"]), inv["number"], inv["code"],
                idx, inv["number"], "", "", inv["date"], "杭州示例科技有限公司",
                inv["seller"], inv["item"], "3%", inv["tax"], round(inv["amount"] - inv["tax"], 2),
                inv["amount"], "示例开票员", "",
            ] + [None] * 5 + [inv["category"]])
    make_xlsx(
        os.path.join(BASE, "日常/processed_invoices.xlsx"),
        [("Invoice Data", rows, None)],
    )


def gen_approvals():
    for appr in APPROVALS:
        lines = [
            appr["type"],
            f"申请编号:{appr['bill']}",
            f"单据编号 {appr['doc']}",
            f"申请人 {appr['applicant']}（{appr['dept']}）",
            f"审批状态 {appr['status']}",
            f"费用明细：交通费 860.00 元 / 住宿费 1180.00 元",
        ]
        make_pdf(os.path.join(BASE, appr["file"]), [lines])


def gen_namelist_pack():
    keyword = "张三-住宿费"
    main_pdf = os.path.join(BASE, "月底/附件", f"报销单-{keyword}", f"差旅报销单-{keyword}.pdf")
    make_pdf(main_pdf, [[
        "差旅报销单",
        "关联申请单号：KZCLSQ202603010001",
        f"报销人：张三",
        "费用合计：1180.00 元",
        "出差日期：2026-03-02 至 2026-03-04",
    ]])
    make_pdf(
        os.path.join(BASE, "月底/附件", f"报销单-{keyword}", "住宿发票.pdf"),
        [["电子发票（普通发票）", "发票号码：26332000000012345601", "*住宿服务*住宿费 1180.00"]],
    )
    make_pdf(
        os.path.join(BASE, "月底/附件", f"报销单-{keyword}", "行程单.pdf"),
        [["航空运输电子客票行程单", "旅客姓名：张三", "票价：860.00 元"]],
    )
    make_pdf(
        os.path.join(BASE, "月底/申请单", "申请单-KZCLSQ202603010001.pdf"),
        [[
            "差旅申请单",
            "申请单编号 KZCLSQ202603010001",
            "申请人 张三（研发中心）",
            "成本编号/名称",
            "成本编号:与项目或客户等相关的成本,请务必填写成本编号,非必要不可选择“无”",
            "PRJ-2026-018 市场部差旅",
            "外出开始时间-外出 2026-03-02 09:00",
        ]],
    )
    make_xlsx(os.path.join(BASE, "月底/namelist.xlsx"), [("Sheet1", [["关键词"], [keyword]], [20])])


def gen_voucher_pdf():
    pages = []
    specs = [
        ("记-001-1/2", [
            "记账凭证",
            "日期: 2026-01-31  凭证号: 记-001-1/2",
            "摘要：报销差旅费",
            "借：管理费用-差旅费 1,000.00",
            "贷：库存现金 1,000.00",
        ]),
        ("记-001-2/2", [
            "记账凭证",
            "日期: 2026-01-31  凭证号: 记-001-2/2",
            "附单据数：3",
            "制单：示例会计  审核：示例主管",
        ]),
        ("记-002-1/1", [
            "记账凭证",
            "日期: 2026-01-31  凭证号: 记-002-1/1",
            "摘要：采购办公用品",
            "借：管理费用-办公费 500.00",
            "贷：银行存款 500.00",
        ]),
    ]
    for _, lines in specs:
        pages.append(lines)
    make_pdf(os.path.join(BASE, "归档/记账凭证-202601.pdf"), pages)


def gen_audit_manifest():
    rows = [
        ["凭证类别", "凭证号", "凭证日期", "附单据数", "摘要"],
        ["记", 1, "2026-01-31", 3, "报销差旅费"],
        ["记", 2, "2026-01-31", 2, "采购办公用品"],
    ]
    make_xlsx(os.path.join(BASE, "归档/凭证清单.xlsx"), [("Sheet1", rows, [10, 10, 14, 10, 24])])


def gen_voucher_compare_lists():
    base_header = ["凭证类别", "凭证号", "凭证日期", "附单据数", "摘要", "科目编码", "科目名称", "借方金额", "贷方金额", "制单人", "审核人"]

    before = [
        ["记", 1, "2026-01-31", 3, "报销差旅费", "5602", "管理费用", 1000.00, 0, "示例会计", "示例主管"],
        ["记", 1, "2026-01-31", 3, "报销差旅费", "1001", "库存现金", 0, 1000.00, "示例会计", "示例主管"],
        ["记", 2, "2026-01-31", 2, "采购办公用品", "5602", "管理费用", 500.00, 0, "示例会计", "示例主管"],
        ["记", 2, "2026-01-31", 2, "采购办公用品", "1002", "银行存款", 0, 500.00, "示例会计", "示例主管"],
        ["记", 3, "2026-01-31", 1, "支付房租", "5602", "管理费用", 8000.00, 0, "示例会计", "示例主管"],
        ["记", 3, "2026-01-31", 1, "支付房租", "1002", "银行存款", 0, 8000.00, "示例会计", "示例主管"],
        ["记", 6, "2026-01-31", 2, "计提折旧-管理", "5602", "管理费用", 2000.00, 0, "示例会计", "示例主管"],
        ["记", 6, "2026-01-31", 2, "计提折旧-生产", "5101", "制造费用", 3000.00, 0, "示例会计", "示例主管"],
        ["记", 6, "2026-01-31", 2, "计提折旧", "1602", "累计折旧", 0, 5000.00, "示例会计", "示例主管"],
    ]
    after = [
        ["记", 1, "2026-01-31", 3, "报销差旅费", "5602", "管理费用", 1000.00, 0, "示例会计", "示例主管"],
        ["记", 1, "2026-01-31", 3, "报销差旅费", "1001", "库存现金", 0, 1000.00, "示例会计", "示例主管"],
        ["记", 2, "2026-01-31", 2, "采购办公用品", "5602", "管理费用", 650.00, 0, "示例会计", "示例主管"],
        ["记", 2, "2026-01-31", 2, "采购办公用品", "1002", "银行存款", 0, 650.00, "示例会计", "示例主管"],
        ["记", 4, "2026-01-31", 1, "支付房租", "5602", "管理费用", 8000.00, 0, "示例会计", "示例主管"],
        ["记", 4, "2026-01-31", 1, "支付房租", "1002", "银行存款", 0, 8000.00, "示例会计", "示例主管"],
        ["记", 5, "2026-01-31", 1, "支付水电费", "5602", "管理费用", 320.00, 0, "示例会计", "示例主管"],
        ["记", 5, "2026-01-31", 1, "支付水电费", "1002", "银行存款", 0, 320.00, "示例会计", "示例主管"],
        ["记", 6, "2026-01-31", 2, "计提折旧-管理部", "5602", "管理费用", 2000.00, 0, "示例会计", "示例主管"],
        ["记", 6, "2026-01-31", 2, "计提折旧-生产", "5101", "制造费用", 3000.00, 0, "示例会计", "示例主管"],
        ["记", 6, "2026-01-31", 2, "计提折旧", "1602", "累计折旧", 0, 5000.00, "示例会计", "示例主管"],
    ]
    make_xlsx(os.path.join(BASE, "归档/凭证清单_修改前.xlsx"), [("Sheet1", [base_header] + before, None)])
    make_xlsx(os.path.join(BASE, "归档/凭证清单_修改后.xlsx"), [("Sheet1", [base_header] + after, None)])


def gen_bank_set():
    pages = []
    for i, tag in enumerate(["记-001", "记-002", "记-003"], start=1):
        pages.append([
            "浙商银行电子回单",
            f"回单编号：200026010{i}",
            "交易日期：2026-01-31",
            f"金额：{i * 1000}.00",
            "对方户名：示例供应商有限公司",
            f"摘要：{tag}",
        ])
    make_pdf(os.path.join(BASE, "归档/银行流水-202601.pdf"), pages)

    rows = [
        ["银行日记账（示例）"],
        ["账号：0000 0000 0000 0000"],
        [],
        ["标签", "日期", "金额"],
        ["记-001", "2026-01-31", 1000.00],
        ["记-002", "2026-01-31", 2000.00],
        ["记-003", "2026-01-31", 3000.00],
        ["合计", "", 6000.00],
        ["制表人：示例会计", "", ""],
    ]
    make_xlsx(os.path.join(BASE, "归档/日记账202601.xlsx"), [("Sheet1", rows, [14, 14, 12])])


def gen_text_diff_pair():
    old_lines = ["服务结算单", "收款方：示例科技服务有限公司", "金额：1,000.00 元", "用途：技术服务费", "结算日期：2026-01-15"]
    new_lines = ["服务结算单", "收款方：示例科技服务有限公司", "金额：1,200.00 元", "用途：技术服务费", "结算日期：2026-01-15"]
    make_pdf(os.path.join(BASE, "归档/结算单-修改前.pdf"), [old_lines])
    make_pdf(os.path.join(BASE, "归档/结算单-修改后.pdf"), [new_lines])


def gen_profit_statements():
    items = [
        ("一、营业收入", 1, 1280000.00),
        ("减：营业成本", 2, 620000.00),
        ("税金及附加", 3, 5600.00),
        ("销售费用", 4, 98400.00),
        ("管理费用", 5, 152300.00),
        ("研发费用", 6, 118000.00),
        ("财务费用", 7, 2100.00),
        ("二、营业利润", 8, 283600.00),
        ("三、利润总额", 9, 283600.00),
        ("减：所得税费用", 10, 42540.00),
        ("四、净利润", 11, 241060.00),
    ]
    month_data = {
        "2026-01": [1280000.00, 620000.00, 5600.00, 98400.00, 152300.00, 118000.00, 2100.00, 283600.00, 283600.00, 42540.00, 241060.00],
        "2026-02": [1410000.00, 668000.00, 6100.00, 104500.00, 158900.00, 126000.00, 1900.00, 344600.00, 344600.00, 51690.00, 292910.00],
    }
    for month, amounts in month_data.items():
        rows = [
            ["利润表", "", ""],
            ["编制单位：杭州示例科技有限公司", "", "单位：元"],
            ["", "", month],
        ]
        for (name, line_no, _), amount in zip(items, amounts):
            rows.append([name, line_no, amount])
        make_xlsx(os.path.join(BASE, f"辅助/利润表{month.replace('-', '')}.xlsx"), [("利润表", rows, [28, 8, 14])])


def gen_cashflow_statements():
    items = [
        ("一、经营活动产生的现金流量", 1),
        ("销售商品、提供劳务收到的现金", 2),
        ("经营活动现金流入小计", 3),
        ("购买商品、接受劳务支付的现金", 4),
        ("经营活动现金流出小计", 5),
        ("经营活动产生的现金流量净额", 6),
        ("二、投资活动产生的现金流量净额", 7),
        ("三、筹资活动产生的现金流量净额", 8),
        ("四、现金及现金等价物净增加额", 9),
    ]
    month_data = {
        "202601": [None, 1360000.00, 1360000.00, 842000.00, 842000.00, 518000.00, -96000.00, 0.00, 422000.00],
        "202602": [None, 1495000.00, 1495000.00, 918000.00, 918000.00, 577000.00, -88000.00, 0.00, 489000.00],
    }
    for month, amounts in month_data.items():
        rows = [
            ["现金流量表", ""],
            ["编制单位：杭州示例科技有限公司", ""],
            ["", ""],
        ]
        for (name, line_no), amount in zip(items, amounts):
            rows.append([name, None, line_no, None, amount])
        make_xlsx(os.path.join(BASE, f"辅助/现金流量表{month}.xlsx"), [("现金流量表", rows, [34, 8, 14])])


def clean():
    for sub in ["日常", "月底", "归档", "辅助"]:
        path = os.path.join(BASE, sub)
        if os.path.exists(path):
            shutil.rmtree(path)


def main():
    clean()
    gen_invoices()
    gen_invoice_master()
    gen_invoices_numlist()
    gen_processed_invoices_snapshot()
    gen_approvals()
    gen_namelist_pack()
    gen_voucher_pdf()
    gen_audit_manifest()
    gen_voucher_compare_lists()
    gen_bank_set()
    gen_text_diff_pair()
    gen_profit_statements()
    gen_cashflow_statements()
    ensure_dir(BASE)
    print("嵌入字体:", FONT_FILE or "未找到系统 CJK 字体，回退内置 china-s（文本提取可能受限）")
    print("sample_data 生成完成：", BASE)
    for root, dirs, files in os.walk(BASE):
        if files:
            print(" ", os.path.relpath(root, BASE), "->", len(files), "个文件")


if __name__ == "__main__":
    main()
