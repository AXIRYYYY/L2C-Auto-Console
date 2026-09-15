# 功能关联地图（FUNCTION_MAP）

> 本文件由 `docs/demo/build_console.py` 从 [`docs/function_map.json`](function_map.json) 自动生成，与 `docs/demo/console.html` 数据同源。

> 生成时间：2026-09-14 ｜ 脚本总数：**59** ｜ 核心脚本（未公开）：**18**

> 公开范围：公开内容仅含演示层（控制台/文档/回放数据/生成器）；业务脚本实现未公开，回放为真实执行的采集快照

## 关系模型

数据流二部图：脚本 ↔ 中间产物 ↔ 业务阶段（脚本之间无代码级 import，关联全部通过文件/剪贴板约定建立）。

## 阶段分布

| 阶段 | 脚本数 | 说明 |
|------|--------|------|
| 日常 | 8 | 实时：凭证摘要生成、发票核验分类、OFD 格式转换 |
| 月底 | 10 | 批量：审批单重命名、资料包匹配整合、检索归集 |
| 归档 | 12 | 核算：凭证比对、PDF 拆分、审计提取、银行流水拆分 |
| 辅助 | 26 | 支撑：报表合并、文件治理、脱敏、文档渲染 |
| 实验 | 3 | 探索：Selenium 网页自动化，未接入生产链路 |

## 日常（8）

### get_filename_from_feishu

`integration/get_filename_from_feishu.py` ｜ 核心脚本 · 核心链路 · 示意回放

监控剪贴板的飞书审批详情，生成「单据类型-编号-单据编号-申请人」标准文件名并回写剪贴板。

- 输入：飞书审批详情页全选复制文本
- 输出：标准文件名（写入剪贴板，供凭证摘要使用）
- 交互：tkinter GUI、pyperclip
- 依赖：pyperclip
- 要点：滑动窗口：以「审批详情」为锚点向前 1000 / 向后 2000 字符截取，容忍页面结构局部变化
- 要点：申请人识别失败时弹出人工修正框，展示「错误」占位
- 要点：部门词表需按公司实际维护（正则硬编码）

### 134_invoicetofoldertxtclose

`file_processing/134_invoicetofoldertxtclose.py` ｜ 核心脚本 · 核心链路 · 实测回放

发票 PDF 号码匹配全量发票表 + 5 类关键词分类，输出台账与成功/失败归档。

- 输入：待处理发票 PDF 池（递归）；全量发票查询导出结果.xlsx（发票基础信息）
- 输出：提取结果/processed_invoices.xlsx；提取结果/unprocessed_files.xlsx；提取成功/提取失败副本（序号_原名.pdf）
- 依赖：openpyxl, pdfplumber
- 替代：file_processing/invoicetofoldertxtclose_v1.py
- 要点：8 位号码查「发票号码」列、20 位查「数电发票号码」列
- 要点：关键词：住宿服务/餐饮服务/汽油/通行费/停车 → 第 24 列「发票类别」
- 要点：自动排除 提取结果/ 目录避免自我循环

### summary_by_invoice_category

`accounting/summary_by_invoice_category.py` ｜ 核心脚本 · 核心链路 · 实测回放

按发票类别汇总价税合计金额，生成分类汇总表（含合计行）。

- 输入：processed_invoices.xlsx（根目录）
- 输出：summary_by_invoice_category.xlsx（发票类别 × 金额 + 合计）
- 依赖：openpyxl
- 要点：读取第 24 列类别 / 第 16 列价税合计（索引 23/15），字段位置是硬约定

### invoice_data_extractor

`accounting/invoice_data_extractor.py` ｜ 核心脚本 · 核心链路 · 实测回放

按发票号码清单批量从全量发票表中回捞完整发票信息。

- 输入：invoicesnumlist.xlsx（第一列号码）；全量发票查询导出结果.xlsx（发票基础信息）
- 输出：extracted_invoice_result.xlsx（提取数字 + 完整发票字段）
- 依赖：pandas
- 要点：按号码长度路由到「发票号码」或「数电票号码」列
- 要点：未匹配行保留原号并打印提示

### convert_all

`file_processing/convert_all.py` ｜ 核心脚本 · 核心链路 · 示意回放

把根目录 OFD 文件批量转换为 PDF（PyMuPDF，带详细错误诊断）。

- 输入：根目录 *.ofd 文件
- 输出：同名 .pdf（已存在则跳过）
- 交互：input() 回车退出
- 依赖：PyMuPDF
- 替代：file_processing/ofdtopdf.py
- 要点：直接 fitz.open(ofd) → save(pdf)，依赖新版 PyMuPDF 的 OFD 支持
- 要点：打印 traceback 便于定位个别损坏文件

### invoice

`accounting/invoice.py` ｜ 在用 · 实测回放

发票 PDF 提取号码并匹配全量发票表，输出根目录台账（含逐票 TXT 文本快照）。

- 输入：待处理发票 PDF 池（递归）；全量发票查询导出结果.xlsx
- 输出：processed_invoices.xlsx（根目录）；unprocessed_files.xlsx；提取结果/提取成功｜提取失败；同名 .txt 文本快照
- 依赖：openpyxl, pdfplumber
- 替代：archive/invoice_2024-12-13.py
- 要点：会为每张 PDF 落一份 .txt 全文，便于人工核查
- 要点：台账输出在根目录，直接对接 summary_by_invoice_category.py

### invoicetofoldertxtclose_v1

`file_processing/invoicetofoldertxtclose_v1.py` ｜ 已被替代

发票 PDF 匹配全量发票表并分类归档（3 类关键词旧版）。

- 输入：待处理发票 PDF 池；全量发票查询导出结果.xlsx
- 输出：提取结果/processed_invoices.xlsx；提取成功/提取失败副本
- 依赖：openpyxl, pdfplumber
- 被替代：file_processing/134_invoicetofoldertxtclose.py
- 要点：关键词仅 住宿服务/餐饮服务/汽油
- 要点：写死「提取结果」为其下目录名

### ofdtopdf

`file_processing/ofdtopdf.py` ｜ 已废弃

根目录 OFD 转 PDF（easyofd，早期版本）。

- 输入：根目录 *.ofd
- 输出：同名 .pdf
- 依赖：easyofd
- 被替代：file_processing/convert_all.py
- 要点：sys.setdefaultencoding 为 Python 2 写法，运行即 AttributeError
- 要点：已被 convert_all.py（fitz 方案）取代

## 月底（10）

### pdf_processor_for_approval

`file_processing/pdf_processor_for_approval.py` ｜ 核心脚本 · 核心链路 · 实测回放

提取审批单 PDF 的申请人/申请编号/单据编号/审批状态/单据类型，按标准格式批量重命名。

- 输入：项目根目录下的审批单 PDF（递归仅根目录一层）
- 输出：{关键字}-{申请编号}-{单据编号}-{申请人}-{审批状态}-{序号}.pdf
- 依赖：PyMuPDF, pdfplumber
- 替代：202602/飞书/pdf_processor_for_approval.py
- 要点：申请人同名时以出现次数作序号后缀
- 要点：关键字表覆盖 12 种审批类型（差旅/招待/采购/付款/开票…）
- 要点：非法字符清洗后再拼文件名

### namelistget_feishu_v9

`integration/namelistget_feishu_v9.py` ｜ 核心脚本 · 核心链路 · 实测回放

按 namelist 关键词匹配「报销单文件夹 + 附件 + 关联申请单」，复制归集、合并为资料包 PDF 并提取成本编号。

- 输入：namelist.xlsx（A 列关键词，需与脚本同目录）；关键词文件源文件夹（弹窗选择）；差旅申请单搜索文件夹（弹窗选择）
- 输出：搜索结果/{关键词}/ 资料包；搜索结果/{关键词}.pdf 合并件；关键词与PDF文件对应表.xlsx；合并文件报错信息.xlsx
- 交互：tkinter.filedialog、input()
- 依赖：openpyxl, pandas, PyMuPDF, Pillow, pypdf
- 替代：integration/namelistget_with_folder.py
- 要点：识别主报销单「-{关键词}」与关联申请单「XXCLSQ...」，图片附件转 PDF 后合并
- 要点：从申请单提取「成本编号/名称」并回填对应表
- 要点：合并失败时把文件夹重命名为「合并失败_」并记录错误报告（数据可追溯）

### get_dingding_form

`integration/get_dingding_form.py` ｜ 在用

在钉钉导出的 PDF 目录中按 num.xlsx 关键词检索并复制命中文件（系列初版）。

- 输入：num.xlsx（第一列）；D:\test\tt\pdf 下的 PDF
- 输出：搜索结果/ PDF 副本 + 根目录 search_results.xlsx
- 依赖：PyMuPDF, pandas, openpyxl
- 被替代：202602/飞书/get_dingding_form_keyword_line_1_and_2_with_name_from_line2_with_date_in_result.py
- 要点：SHA256 去重避免同文件重复复制
- 要点：路径 D:\test\tt\pdf 为写死的测试路径
- 要点：金额格式化 {:.2f} 后检索

### get_dingding_form 增强版（双关键词+日期列）

`202602/飞书/get_dingding_form_keyword_line_1_and_2_with_name_from_line2_with_date_in_result.py` ｜ 在用

双列关键词 + 文件名筛选 + 结果表附带申请提交日期，钉钉检索系列最终版。

- 输入：num.xlsx（关键词/第二关键词/文件名筛选 三列）；D:\work\凭证 下 PDF
- 输出：搜索结果/ + search_results_{时间戳}.xlsx（含第二关键词/文件名筛选/申请提交日期列）
- 交互：input()
- 依赖：PyMuPDF, pandas, openpyxl
- 替代：integration/get_dingding_form_keyword.py
- 要点：兼容千分位金额与大小写文件名
- 要点：时间戳格式串 %m 重复的小瑕疵（不影响检索）

### namelistget_with_folder

`integration/namelistget_with_folder.py` ｜ 已被替代

按关键词从 D 盘搜索复制文件与整个文件夹（v2.0）。

- 输入：namelist.xlsx；D:/ 全盘文件
- 输出：搜索结果/（文件+文件夹副本）
- 交互：input()
- 依赖：openpyxl
- 替代：integration/namelistget.py
- 被替代：integration/namelistget_feishu_v9.py
- 要点：D:/ 全盘 os.walk，权限失败仅打印跳过
- 要点：新增文件夹整包 copytree

### namelistget

`integration/namelistget.py` ｜ 已被替代

按关键词从 D 盘搜索并复制文件（v1 原型）。

- 输入：namelist.xlsx；D:/ 全盘文件
- 输出：搜索结果/
- 依赖：openpyxl
- 被替代：integration/namelistget_with_folder.py
- 要点：逐关键词全盘遍历，性能差
- 要点：结果目录固定为项目根目录/搜索结果

### namelistget_v9.0（来源可选参数变体）

`202602/飞书/namelistget_current file_forfeishu_v9.0增加来源可选参数.py` ｜ 已被替代

v9.0 变体：新增关键词来源可选参数（自动/弹窗）。

- 输入：namelist.xlsx；关键词源文件夹（KEYWORD_SOURCE_MODE=2 弹窗）
- 输出：搜索结果/、{关键词}.pdf、对应表.xlsx、报错信息.xlsx
- 交互：tkinter.filedialog、input()
- 依赖：openpyxl, pandas, PyMuPDF, Pillow, pypdf
- 替代：integration/namelistget_with_folder.py
- 被替代：integration/namelistget_feishu_v9.py
- 要点：KEYWORD_SOURCE_MODE 1=脚本目录 2=弹窗选择，默认 2
- 要点：主版本已合入该能力，此文件为历史快照

### get_dingding_form_keyword

`integration/get_dingding_form_keyword.py` ｜ 已被替代

数字多格式展开检索（x.xx/整数/x.0）并输出带时间戳结果表。

- 输入：num.xlsx；D:\test\tt\pdf
- 输出：搜索结果/ + search_results_{时间戳}.xlsx
- 交互：input()
- 依赖：PyMuPDF, pandas, openpyxl
- 替代：integration/get_dingding_form.py
- 被替代：202602/飞书/get_dingding_form_keyword_line_1_and_2_with_name_from_line2_with_date_in_result.py
- 要点：数字三形态展开提高命中率
- 要点：SHA256 去重

### get_dingding_form_get_text

`integration/get_dingding_form_get_text.py` ｜ 已被替代

修复「末尾 .0 误伤正则」问题的检索版本。

- 输入：num.xlsx；D:\test\tt\pdf
- 输出：搜索结果/ + search_results_{时间戳}.xlsx
- 交互：input()
- 依赖：PyMuPDF, pandas, openpyxl
- 替代：integration/get_dingding_form.py
- 被替代：202602/飞书/get_dingding_form_keyword_line_1_and_2_with_name_from_line2_with_date_in_result.py
- 要点：关键词取原值不展开，str() 兜底
- 要点：末尾 0 处理不再破坏正则结构

### pdf_processor_for_approval（202602 快照）

`202602/飞书/pdf_processor_for_approval.py` ｜ 已被替代

审批单重命名器的历史快照（脚本目录级扫描）。

- 输入：脚本所在文件夹下审批单 PDF
- 输出：{关键字}-{申请编号}-{单据编号}-{申请人}-{审批状态}-{序号}.pdf
- 依赖：PyMuPDF, pdfplumber
- 被替代：file_processing/pdf_processor_for_approval.py
- 要点：主版本已合入相关适配

## 归档（12）

### compare_vouchers_advanced

`accounting/compare_vouchers_advanced.py` ｜ 核心脚本 · 核心链路 · 实测回放

深度比对新旧凭证清单：编号变化/内容变化/新增/删除四分类，用 Jaccard 相似度识别「换号但内容一致」的凭证。

- 输入：修改前凭证清单 xlsx（弹窗）；修改后凭证清单 xlsx（弹窗）
- 输出：凭证比对结果_{时间戳}/凭证比对分析报告(最终版)_{时间戳}.xlsx（摘要/编号内容均修改/仅内容/仅编号/新增/删除）；源文件复制归档
- 交互：tkinter.filedialog
- 依赖：pandas, openpyxl
- 替代：accounting/voucher_comparison.py
- 要点：凭证内容签名用 frozenset，行序无关
- 要点：Jaccard 阈值：>0.5 判「编号+内容修改」、<0.2 判删除+新增
- 要点：附单据数不一致时自动标红，方便审计核对面单数量

### 清单提取拆分好的凭证

`accounting/清单提取拆分好的凭证.py` ｜ 核心脚本 · 核心链路 · 实测回放

按凭证清单从拆分好的凭证 PDF 池中匹配提取，生成审计资料包与提取报告。

- 输入：拆分后的凭证 PDF 文件夹（弹窗）；凭证清单 Excel（弹窗）
- 输出：凭证提取结果_{时间戳}/（命中 PDF + 提取结果报告.xlsx + 清单副本）
- 交互：tkinter.filedialog
- 依赖：pandas, openpyxl
- 要点：文件名前缀约定：{YYYY-MM}-{凭证类别}-{凭证号zfill(3)}
- 要点：命中状态三态：成功提取/未能找到/数据格式错误

### split_vouchers_凭证PDF拆分

`accounting/split_vouchers_凭证PDF拆分.py` ｜ 核心脚本 · 核心链路 · 实测回放

把整本凭证 PDF 按「年月 + 基础凭证号」逐张拆分为独立 PDF。

- 输入：整本记账凭证 PDF（弹窗）
- 输出：{凭证名}_拆分结果/{YYYY-MM}-{基础凭证号}-{N}页.pdf
- 交互：tkinter.filedialog
- 依赖：pypdf
- 要点：正则识别「日期:」「凭证号:」含全角冒号
- 要点：解析不到标准格式时按暂记凭证模式兜底
- 要点：页数写入文件名，纸质核对时可快速点页

### bankStatementSplitterBasedOnJournal

`accounting/bankStatementSplitterBasedOnJournal.py` ｜ 核心脚本 · 核心链路 · 实测回放

按电子日记账标签把银行流水 PDF 按凭证号逐笔拆分为系列 PDF。

- 输入：脚本同目录唯一的银行流水 PDF；唯一含「日记账」的 xlsx（A 列标签，行 5 起）
- 输出：分割完成/{标签}.pdf
- 依赖：pypdf, openpyxl
- 替代：archive/InputPathDrivenPDFJournalSplitter_2025-01-20.py
- 要点：强约束：脚本同目录必须「恰好一个」PDF 与「恰好一个」日记账 xlsx，否则退出
- 要点：行数 vs 页数差 >2 时判定文件错配直接退出（防御性设计）
- 要点：标签即新文件名（通常是凭证号）

### 新旧凭证列表比较及差异输出

`accounting/新旧凭证列表比较及差异输出.py` ｜ 核心脚本 · 实测回放

按「除类别/号/日期外的全部列」匹配新旧凭证，输出整组匹配结果与未匹配行并红字标注。

- 输入：旧凭证列表 xlsx；新凭证列表 xlsx（控制台输入路径）
- 输出：比较结果_{时间戳}.xlsx（匹配结果/旧证未匹配/新证未匹配/凭证号匹配结果，差异红字）
- 交互：input() 文件路径
- 依赖：pandas, openpyxl
- 替代：archive/新旧凭证列表比较及差异输出_2025-01-24.py
- 要点：按凭证号分组后要求「整组完全匹配」才算匹配
- 要点：与 advanced 版互补：不做相似度，做整组一致性

### 文字pdf对比脚本

`accounting/文字pdf对比脚本.py` ｜ 在用 · 实测回放

两份 PDF 逐页提取文本并比对差异，生成差异报告。

- 输入：旧 PDF；新 PDF（弹窗各选一次）
- 输出：PDF文本比对结果_{时间戳}/文本比对结果报告.xlsx（逐页：无差异/有差异+unified diff）
- 交互：tkinter.filedialog
- 依赖：pdfplumber, pandas, openpyxl
- 要点：空白规范化后先等值比较，不同才输出 unified_diff
- 要点：报告页级差异，附完整 diff 文本

### InputPathDrivenPDFJournalSplitter

`accounting/InputPathDrivenPDFJournalSplitter.py` ｜ 无法运行

同上但改为控制台输入路径驱动的变体。

- 输入：PDF 路径、日记账 xlsx 路径（控制台输入）
- 输出：分割完成/{标签}.pdf
- 交互：input()
- 依赖：PyPDF2, openpyxl
- 被替代：accounting/bankStatementSplitterBasedOnJournal.py
- 要点：第 1 行内容为 D:\work\... 路径 + import 语句粘连
- 要点：逻辑与 bankStatementSplitterBasedOnJournal 相同，仅入口不同

### voucher_comparison

`accounting/voucher_comparison.py` ｜ 已被替代

比对新旧凭证的编号变化/内容变化/新增/删除（简化版，无相似度）。

- 输入：修改前凭证清单；修改后凭证清单
- 输出：凭证比对结果_{时间戳}.xlsx（只修改了编号/内容修改/已删除/新增）
- 交互：tkinter.filedialog
- 依赖：pandas, numpy
- 替代：accounting/voucher_comparison_for凭证编号未修改.py
- 被替代：accounting/compare_vouchers_advanced.py
- 要点：无 Jaccard 相似度，换号识别靠「类别|日期」标识
- 要点：依赖科目编码+摘要排序后逐行 equals

### voucher_comparison（早期版）

`accounting/voucher_comparison_for凭证编号未修改.py` ｜ 已被替代

凭证比对最早版本（针对编号未修改场景）。

- 输入：新旧凭证清单
- 输出：凭证比对结果 xlsx
- 交互：tkinter.filedialog
- 依赖：pandas
- 被替代：accounting/voucher_comparison.py
- 要点：内容指纹逻辑已在早期引入（get_voucher_content_maps）

### 新旧凭证列表比较（2025-01-24 归档）

`archive/新旧凭证列表比较及差异输出_2025-01-24.py` ｜ 已被替代

上述脚本的旧快照。

- 输入：新旧凭证列表
- 输出：比较结果_{时间戳}.xlsx
- 交互：input()
- 依赖：pandas
- 被替代：accounting/新旧凭证列表比较及差异输出.py
- 要点：未匹配行用索引 isin 判断，存在错位风险

### InputPathDrivenPDFJournalSplitter（2025-01 归档）

`archive/InputPathDrivenPDFJournalSplitter_2025-01-20.py` ｜ 已被替代

路径驱动拆分器旧快照（可运行版本）。

- 输入：PDF 路径、日记账 xlsx 路径
- 输出：分割完成/{标签}.pdf
- 交互：input()
- 依赖：PyPDF2, openpyxl
- 被替代：accounting/InputPathDrivenPDFJournalSplitter.py
- 要点：逻辑与当前版一致

### invoice（2024-12-13 归档）

`archive/invoice_2024-12-13.py` ｜ 已被替代

发票匹配最早归档版。

- 输入：发票 PDF + 全量发票查询导出结果.xlsx
- 输出：提取结果/processed_invoices.xlsx + 成功/失败副本
- 依赖：openpyxl, pdfplumber
- 被替代：accounting/invoice.py
- 要点：新版简化列结构并改从项目根目录扫描

## 辅助（26）

### 利润表生成

`accounting/利润表生成.py` ｜ 核心脚本 · 核心链路 · 实测回放

扫描根目录全部利润表 xlsx，按「利润表项目」外连接合并为多月对比表并补行次。

- 输入：项目根目录 包含「利润表」的 *.xlsx（C3 为月份）
- 输出：合并利润表_带行次_{时间戳}.xlsx（行次+项目+各月列）
- 依赖：pandas, openpyxl
- 要点：行次随机取一个样本文件（各月行次一致时无影响）
- 要点：outer merge 容忍项目名称增减
- 要点：日志写入 merge_log.log

### 合并现金流量表

`accounting/合并现金流量表.py` ｜ 核心脚本 · 核心链路 · 实测回放

按文件名中的年月排序，合并多月现金流量表，先解除合并单元格再提取数据。

- 输入：项目根目录 匹配「现金流*(\d{4})(\d{2}).xlsx」的月度表
- 输出：现金流量表合并_{时间戳}.xlsx（行次+项目+各月列，已设列宽）
- 依赖：pandas, openpyxl, xlsxwriter
- 要点：先 openpyxl 解除合并单元格并填充值，另存临时文件再 pandas 提取
- 要点：依赖 xlsxwriter（requirements.txt 未列出，属依赖清单缺口）
- 要点：日志写入 cashflow_merge.log

### render_mermaid

`docs/diagrams/render_mermaid.py` ｜ 核心脚本

读取 docs/diagrams/*.mmd，调用 mermaid.ink 渲染并下载 PNG。

- 输入：docs/diagrams/*.mmd
- 输出：同名 .png（zlib+pako+base64url 编码调用在线渲染）
- 要点：纯标准库实现（json/zlib/base64/urllib）
- 要点：依赖 mermaid.ink 在线服务，请求间隔 3 秒防限流

### analyze_duplicates

`accounting/analyze_duplicates.py` ｜ 在用 · 实测回放

仓库自检工具：扫描全部 .py，按核心名分组 + difflib 相似度，输出保留/审查/删除建议报告。

- 输入：项目根目录全部 .py 源码（排除 .pyc 等）
- 输出：文件整理分析报告.md（重复组相似度 + 处理建议，只读不改文件）
- 依赖：difflib（标准库）
- 要点：能解析备份时间戳/副本标记/版本号后缀三类命名噪声
- 要点：相似度 ≥95% 建议删除、85%+ 建议审查、其余保留
- 要点：明确声明不删移任何文件（安全设计）

### business_trip_report

`accounting/business_trip_report.py` ｜ 在用

出差明细/统计双表生成的示例脚本（张三李四王五假数据）。

- 输入：无（脚本内假数据）
- 输出：员工出差统计表.xlsx（明细记录+天数统计）
- 依赖：pandas
- 要点：自动计算出差天数与占比
- 要点：生成在 CWD

### 133_extract_files

`file_processing/133_extract_files.py` ｜ 在用

把子文件夹内文件全部上移到根目录，重名自动 _1/_2 递增并清理空目录。

- 输入：项目根目录所有子文件夹内文件
- 输出：原地移动后的文件；空目录被删除
- 要点：会扫描包括 .git 在内的所有子目录（危险面较大，需在专用工作目录运行）

### pdf_processor_to_image

`file_processing/pdf_processor_to_image.py` ｜ 在用

根目录 PDF 逐页转 JPG（dpi=600，打印用）。

- 输入：项目根目录 *.pdf
- 输出：PDF 同目录 {名}_{页码}.jpg
- 依赖：PyMuPDF
- 替代：file_processing/PDFtoImage.py
- 要点：注释写 300 DPI，实际 600 DPI（注释与实现不一致）
- 要点：输出名含绝对路径导致 output_folder 参数失效，JPG 落在 PDF 旁边

### pdf_renamer

`file_processing/pdf_renamer.py` ｜ 在用

去掉 PDF 文件名尾部 (1)(2) 并处理重名冲突。

- 输入：项目根目录 *.pdf
- 输出：重命名后的文件；与现有文件冲突时移入 重复文件/
- 要点：正则只匹配半角括号结尾
- 要点：仅处理根目录一层

### delete_duplicate_files_v2

`file_processing/delete_duplicate_files_v2.py` ｜ 在用

SHA256 查重，每组保留修改时间最新者（v2 修正版）。

- 输入：项目根目录及子目录全部文件
- 输出：较旧的重复副本移入回收站
- 依赖：send2trash
- 替代：file_processing/delete_duplicate_files.py
- 要点：相对 v1 唯一改动：排序 reverse=True

### merge_tables

`file_processing/merge_tables.py` ｜ 在用

把根目录全部 xlsx 纵向合并为一张表。

- 输入：项目根目录全部 .xlsx
- 输出：combined_data.xlsx
- 依赖：pandas
- 要点：重跑会把 combined_data.xlsx 自身再次并入（已知缺陷）
- 要点：按列位置堆叠，不校验表头

### new_folder_with_subfolders

`file_processing/new_folder_with_subfolders.py` ｜ 在用

按配置批量创建「本月发票处理」等月度工作目录与占位文件。

- 输入：脚本内 folder_config 配置
- 输出：本月发票处理/（含 流水/农行/支付宝/农行分割/浙商/携程、飞书、报销附件、零碎的资料）+ 两个占位文件
- 要点：占位文件是空文件（不是真 Excel）
- 要点：makedirs exist_ok 幂等，已存在不覆盖

### onlyforcxy

`utils/onlyforcxy.py` ｜ 在用

按文件夹递归合并 PDF 为「{文件夹名}_合并版.pdf」。

- 输入：目标目录及各子目录 PDF（代码内硬编码路径）
- 输出：各文件夹内 {文件夹名}_合并版.pdf（按文件名排序）
- 依赖：pypdf
- 要点：硬编码 D:\work\凭证\<内部目录>\...（因此被 gitignore 排除）

### pythonrc

`utils/pythonrc.py` ｜ 在用

VS Code Python REPL 提示符集成配置。

- 要点：Windows 下仅打印操作提示，不设置 PS1

### 脱敏汇总

`脱敏汇总.py` ｜ 在用

data/ 下四张业务表首行数据脱敏并汇总输出对照明细。

- 输入：data/审批数据总表.xlsx 等四表
- 输出：data/4表脱敏汇总结果.xlsx（汇总总览+对照明细）
- 依赖：openpyxl
- 替代：accounting/脱敏汇总.py
- 要点：姓氏白名单避免把部门名误脱敏
- 要点：公开仓库数据脱敏流程的上游工具

### IntercompanyOutstandingCleanupTool

`accounting/IntercompanyOutstandingCleanupTool.py` ｜ 半成品 · 示意回放

关联方往来清理：余额表+明细账匹配未清项，并可检索对应审批单 PDF。

- 输入：column_info.xlsx（列映射配置，需人工填写）；全部科目余额表 xlsx；科目辅助明细账 xlsx；可选：凭证 PDF 目录
- 输出：处理结果.xlsx（工作表1/2）；新建xlsx文件.xlsx（检索关键词）；搜索结果/ + search_results_{时间戳}.xlsx
- 交互：input()
- 依赖：pandas, openpyxl, PyMuPDF
- 要点：some_condition_met 为占位变量，未实现判断逻辑（README 已标注开发中）
- 要点：零配置列映射思路：column_info.xlsx 声明「文件-列名-列号」
- 要点：金额检索兼容千分位逗号格式

### processedfilesinfoalltxt

`file_processing/processedfilesinfoalltxt.py` ｜ 半成品 · 实测回放

审批单 PDF 金额+日期提取，输出 Excel 并逐票落 TXT 文本（增强版）。

- 输入：项目根目录及子目录中文件名含「审批/提交」的 PDF
- 输出：processed_files.xlsx + 每份 PDF 同名 .txt（项目根目录）
- 依赖：openpyxl, pdfplumber, locale
- 替代：file_processing/processedfilesinfoall.py
- 要点：重复关键词自动新增「关键词_2」列
- 要点：从「水印名」后提取 YYYY-MM-DD HH:MM:SS 日期

### fileinfonum

`file_processing/fileinfonum.py` ｜ 半成品

扫描运行目录文件名并提取其中数字写入 Excel。

- 输入：os.listdir('.') 当前工作目录
- 输出：file_info.xlsx（文件名 + 提取数字）
- 依赖：openpyxl
- 要点：未用 __file__ 定位目录，双击运行扫描目录可能不对

### invoice_data_from_pdf

`accounting/invoice_data_from_pdf.py` ｜ 已被替代

扫描当前目录 PDF，提取审批单金额关键词与日期，生成 processed_files.xlsx（早期版）。

- 输入：当前工作目录及子目录的 PDF（文件名含「审批/提交」）
- 输出：processed_files.xlsx（文件名称+审批编号+各金额关键词列）
- 依赖：openpyxl, pdfplumber
- 被替代：file_processing/processedfilesinfoalltxt.py
- 要点：水印名拆成单字逐个剔除后再抽取
- 要点：关键词：最终优惠价（元）/本单总计/合同总金额/此次报销费用总计金额/金额

### processedfilesinfoall

`file_processing/processedfilesinfoall.py` ｜ 已被替代

从审批单 PDF 批量提取金额字段并汇总为 processed_files.xlsx。

- 输入：项目根目录及子目录中文件名含「审批/提交」的 PDF
- 输出：processed_files.xlsx
- 依赖：openpyxl, pdfplumber
- 替代：accounting/invoice_data_from_pdf.py
- 被替代：file_processing/processedfilesinfoalltxt.py
- 要点：水印逐字删除后去空格换行再匹配
- 要点：金额关键词覆盖合同/报销/优惠价等

### PDFtoImage

`file_processing/PDFtoImage.py` ｜ 已被替代

根目录 PDF 逐页转 PNG（不可批量打印时的替代方案）。

- 输入：项目根目录 *.pdf
- 输出：PDF 同名子文件夹/{pdf名}_{页码}.png
- 依赖：PyMuPDF
- 被替代：file_processing/pdf_processor_to_image.py
- 要点：按 PDF 名建子文件夹
- 要点：页码从 1 开始

### delete_duplicate_files

`file_processing/delete_duplicate_files.py` ｜ 已被替代

SHA256 分块哈希查重，重复文件保留最旧、其余移入回收站（v1）。

- 输入：项目根目录及子目录全部文件
- 输出：重复副本被 send2trash 移入回收站
- 依赖：send2trash
- 被替代：file_processing/delete_duplicate_files_v2.py
- 要点：4096 字节分块计算哈希，避免大文件吃内存
- 要点：实际保留最旧文件（v2 注释承认此策略不理想）

### getpdf_v1

`file_processing/getpdf_v1.py` ｜ 已被替代

按 invoice.xlsx 清单从发票目录匹配复制 PDF（子串匹配原型）。

- 输入：D:\work\凭证\电子发票\提取使用\invoice.xlsx；D:\work\凭证\电子发票
- 输出：匹配 PDF 复制到 提取使用/
- 依赖：pandas
- 被替代：file_processing/getpdf_v2.py
- 要点：文件名数字子串匹配
- 要点：本机无 D:\work，需改造才能运行

### getpdf_v2

`file_processing/getpdf_v2.py` ｜ 已被替代

同上，row[0]→row.iloc[0] 的 pandas 修正版。

- 输入：D:\work\凭证\电子发票\提取使用\invoice.xlsx
- 输出：匹配 PDF 复制到 提取使用/
- 依赖：pandas
- 替代：file_processing/getpdf_v1.py
- 要点：三个 D 盘路径同样硬编码

### 脱敏汇总（旧版）

`accounting/脱敏汇总.py` ｜ 已被替代

四表脱敏汇总旧实现（未适配 data/ 目录）。

- 输入：项目根目录下四个 xlsx（旧位置）
- 输出：项目根目录/4表脱敏汇总结果.xlsx
- 依赖：openpyxl
- 被替代：脱敏汇总.py
- 要点：被根目录新版取代

### 批量新建月份文件夹

`file_processing/批量新建月份文件夹.py` ｜ 已废弃

在根目录批量创建 202401~202412 十二个月份文件夹。

- 输出：202401~202412 文件夹
- 被替代：file_processing/new_folder_with_subfolders.py
- 要点：年份与命名规则全部写死，仅一次性用途

### organize_project

`utils/organize_project.py` ｜ 已废弃

一次性项目整理器：把根目录散落文件归位到 accounting/file_processing/integration/utils/data/archive/docs。

- 输入：根目录散落文件（含 v1/v2 重命名映射表）
- 输出：标准目录结构
- 要点：本次目录结构的缔造者
- 要点：含新老文件名映射，可作为版本谱系的旁证

## 实验（3）

### web_file_downloader_v2

`integration/web_file_downloader_v2.py` ｜ 半成品 · 示意回放

Selenium 打开网页并点击「下载」元素抓取文件。

- 输入：手工输入的 https URL
- 输出：Edge 下载目录 ~/Downloads/Secure_Downloads 下文件
- 交互：input()、selenium
- 依赖：selenium, webdriver-manager
- 替代：integration/web_file_downloader_v1.py
- 要点：XPath 匹配文本「下载」的元素逐个点击
- 要点：未来方向：替代畅捷通/银行网银人工下载

### web_file_downloader_v1

`integration/web_file_downloader_v1.py` ｜ 已被替代

读取剪贴板中的链接，正则批量下载网页文件。

- 输入：剪贴板 http(s) 链接
- 输出：按 Content-Disposition / URL 末段命名的文件（CWD）
- 交互：pyperclip
- 依赖：requests, pyperclip
- 被替代：integration/web_file_downloader_v2.py
- 要点：无去重逻辑

### new（Selenium 驱动测试）

`utils/new.py` ｜ 实验性

Edge 驱动连通性测试片段（自动下载驱动并打开百度）。

- 输出：控制台打印网页标题
- 交互：selenium
- 依赖：selenium, webdriver-manager
- 要点：生产链路无 Selenium 依赖，仅为网页自动化前瞻

## 版本谱系

### namelistget（月底资料包匹配）

从 D 盘全盘搜索 → 文件夹整包复制 → 飞书资料包一条龙（复制+关联+合并+对账表）。

- `integration/namelistget.py`
- `integration/namelistget_with_folder.py`
- `202602/飞书/namelistget_current file_forfeishu_v9.0增加来源可选参数.py`
- `integration/namelistget_feishu_v9.py` ← 当前

### 钉钉表单检索（get_dingding_form）

关键词检索 PDF → 数字多格式展开 → 修复末尾 .0 → 双关键词+文件名筛选+日期列（202602 终版）。

- `integration/get_dingding_form.py`
- `integration/get_dingding_form_get_text.py`
- `integration/get_dingding_form_keyword.py`
- `202602/飞书/get_dingding_form_keyword_line_1_and_2_with_name_from_line2_with_date_in_result.py` ← 当前

### 发票核验（invoice）

归档版 → 根目录台账版（invoice.py）→ 提取结果目录增强版（134：序号+5类关键词）。

- `archive/invoice_2024-12-13.py`
- `accounting/invoice.py`
- `file_processing/invoicetofoldertxtclose_v1.py`
- `file_processing/134_invoicetofoldertxtclose.py` ← 当前

### 凭证比对（voucher comparison）

早期版 → 编号变化识别版 → Jaccard 深度比对 v2.1；另有整组匹配思路的新旧凭证列表比较。

- `accounting/voucher_comparison_for凭证编号未修改.py`
- `accounting/voucher_comparison.py`
- `accounting/compare_vouchers_advanced.py` ← 当前
- `archive/新旧凭证列表比较及差异输出_2025-01-24.py`
- `accounting/新旧凭证列表比较及差异输出.py`
- 当前（另一条路线）：`accounting/新旧凭证列表比较及差异输出.py（两条路线并存）`

### 银行流水拆分（journal splitter）

归档快照 → 路径驱动版（语法损坏）→ 根目录唯一文件约束版。

- `archive/InputPathDrivenPDFJournalSplitter_2025-01-20.py`
- `accounting/InputPathDrivenPDFJournalSplitter.py`
- `accounting/bankStatementSplitterBasedOnJournal.py` ← 当前

### 审批单金额提取（processed files info）

CWD 依赖版 → 工程化版 → TXT+日期增强版（README 标注未完成）。

- `accounting/invoice_data_from_pdf.py`
- `file_processing/processedfilesinfoall.py`
- `file_processing/processedfilesinfoalltxt.py` ← 当前

### OFD 转换（ofd → pdf）

easyofd 早期版（Py2 语法残留）→ PyMuPDF 全量转换版。

- `file_processing/ofdtopdf.py`
- `file_processing/convert_all.py` ← 当前

### 审批单重命名（pdf processor）

202602 快照 → file_processing 适配版（路径解析与关键字表差异）。

- `202602/飞书/pdf_processor_for_approval.py`
- `file_processing/pdf_processor_for_approval.py` ← 当前

### PDF 转图片

PNG 子目录版 → JPG 直出打印版。

- `file_processing/PDFtoImage.py`
- `file_processing/pdf_processor_to_image.py` ← 当前

### 文件查重（delete duplicates）

保留最旧版 → 保留最新修正版。

- `file_processing/delete_duplicate_files.py`
- `file_processing/delete_duplicate_files_v2.py` ← 当前

### Excel 清单匹配复制（getpdf）

子串匹配原型 → pandas iloc 修正版（均依赖 D 盘硬编码路径）。

- `file_processing/getpdf_v1.py`
- `file_processing/getpdf_v2.py` ← 当前

### 网页文件下载（web file downloader）

剪贴板链接 requests 版 → Selenium 点击版（实验性，未接入生产）。

- `integration/web_file_downloader_v1.py`
- `integration/web_file_downloader_v2.py` ← 当前

### 脱敏汇总

根目录旧版 → data/ 目录适配新版。

- `accounting/脱敏汇总.py`
- `脱敏汇总.py` ← 当前

### 月度目录初始化

2024 年份硬编码版（废弃）→ 配置化目录+占位文件版。

- `file_processing/批量新建月份文件夹.py`
- `file_processing/new_folder_with_subfolders.py` ← 当前

## 效能指标（示意口径）

| 指标 | 优化前 | 优化后 | 变化 |
|------|--------|--------|------|
| 单张凭证处理 | 约 8 分钟 | 约 30 秒 | -94% |
| 发票核对与分类 | 约 3 小时/批次 | 约 10 分钟/批次 | -94% |
| 月结周期 | 3 个工作日 | 4 小时 | -94% |
| 审计凭证提取 | 约 2 天 | 约 5 分钟 | -99% |
| PDF 拆分重命名 | 约 1 天 | 约 2 分钟 | -99% |

> 以上为业务口径的估算与推算，未经独立复核，不构成实测基准；前两项给出拆解过程，其余为口径说明。

## 回放清单

| 回放 | 脚本 | 类型 | 说明 |
|------|------|------|------|
| 审批单金额/日期提取 | `file_processing/processedfilesinfoalltxt.py` | 实测 | README 标注未完成；可验证关键词扩列与日期提取逻辑 |
| 审批单标准化重命名 | `file_processing/pdf_processor_for_approval.py` | 实测 | 输出 单据类型-申请编号-单据编号-申请人-审批状态-序号 标准命名 |
| 审计清单提取 | `accounting/清单提取拆分好的凭证.py` | 实测 | 文件名前缀 {YYYY-MM}-{凭证类别}-{凭证号zfill(3)} 匹配，输出提取报告 |
| 银行流水按凭证拆分 | `accounting/bankStatementSplitterBasedOnJournal.py` | 实测 | 根目录必须恰好 1 个 PDF + 1 个日记账；行数与页数差异 >2 直接拒绝 |
| 现金流量表多月合并 | `accounting/合并现金流量表.py` | 实测 | 先解除合并单元格再提取，输出列宽已设置的合并表 |
| 飞书审批摘要（剪贴板） | `integration/get_filename_from_feishu.py` | 示意 | 依赖人工在飞书页面全选复制 + GUI 热键，无法无头执行；本条目为按代码逻辑生成的示意回放 |
| 关联方往来清理（半成品） | `accounting/IntercompanyOutstandingCleanupTool.py` | 示意 | some_condition_met 未定义、依赖人工填写 column_info.xlsx，README 标注开发中；日志为示意 |
| 发票分类汇总 | `accounting/summary_by_invoice_category.py` | 实测 | 读取第 24 列类别与第 16 列金额，输出分类汇总 + 合计 |
| 发票核验（增强版 134） | `file_processing/134_invoicetofoldertxtclose.py` | 实测 | 最新版：序号 + 5 类关键词 + 提取结果目录结构 |
| 发票核验（主链路） | `accounting/invoice.py` | 实测 | 输出 processed_invoices.xlsx 供分类汇总消费；未匹配发票进入提取失败目录 |
| 发票号码批量回捞 | `accounting/invoice_data_extractor.py` | 实测 | 按号码长度路由 8 位/20 位列，回填完整发票字段 |
| 月底资料包匹配整合 | `integration/namelistget_feishu_v9.py` | 实测 | 复制归集 → 关联申请单 → 图片转 PDF → 合并资料包 → 生成对应表与错误报告 |
| OFD 转 PDF | `file_processing/convert_all.py` | 示意 | OFD 需真实电子发票样本；本机无 OFD 文件，日志为按代码逻辑生成的示意回放 |
| PDF 文本逐页比对 | `accounting/文字pdf对比脚本.py` | 实测 | 逐页提取文本，无差异/有差异 + unified diff |
| 利润表多月合并 | `accounting/利润表生成.py` | 实测 | 按利润表项目外连接，行次取样本文件，输出带时间戳合并表 |
| 仓库重复文件自检 | `accounting/analyze_duplicates.py` | 实测 | 按核心名分组 + difflib 相似度，输出 保留/审查/删除 建议报告 |
| 凭证深度比对（Jaccard） | `accounting/compare_vouchers_advanced.py` | 实测 | 覆盖 内容修改/换号/新增/删除 四类差异，附单据数变化标红 |
| 新旧凭证列表整组比对 | `accounting/新旧凭证列表比较及差异输出.py` | 实测 | 整组匹配思路：凭证号整组一致才算匹配，差异行红字标注 |
| 凭证 PDF 拆分 | `accounting/split_vouchers_凭证PDF拆分.py` | 实测 | 按 年月+基础凭证号 分组，输出 {年月}-{凭证号}-{N}页.pdf |
| 网银/网页文件下载（实验） | `integration/web_file_downloader_v2.py` | 示意 | Selenium 需要真实网页与 Edge 会话，未接入生产链路；日志为示意 |

> 实测 = 在隔离工作区用 `sample_data/` 假数据真实执行采集；示意 = 依赖 GUI/外部环境无法无头执行，按代码逻辑生成，已在页面中标注。
