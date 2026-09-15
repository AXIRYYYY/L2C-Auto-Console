"""公开文件敏感信息扫描器。

用法:
  python docs/demo/sensitive_check.py           # 扫描并出报告（退出码 0=通过 / 1=有硬违规）
  python docs/demo/sensitive_check.py --sync    # 依据 public_files.txt 重新生成 public.gitignore

硬违规（必须清零）：本机工作区路径、真实主体黑名单、密钥/口令赋值、手机号、邮箱、身份证号、内网 IP
提示项（人工确认）：其他盘符路径（允许占位符/测试路径/系统字体路径白名单）
"""

import fnmatch
import io
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(os.path.dirname(HERE))
MANIFEST = os.path.join(HERE, "public_files.txt")
PUBLIC_IGNORE = os.path.join(HERE, "public.gitignore")

SKIP_SELF = {"docs/demo/sensitive_check.py"}

SOURCE_OVERRIDES = {
    ".gitignore": "docs/demo/public.gitignore",
}

BLACKLIST = ["千二", "叶佳敏", "qianer", "K2"]

HARD_PATTERNS = [
    ("本机工作区路径", re.compile(r"qianer_auto_python|[A-Za-z]:[\\/]{1,2}(?:coding|Users)[\\/]", re.I)),
    ("真实主体/内部代号黑名单", re.compile("|".join(re.escape(w) for w in BLACKLIST), re.I)),
    ("密钥/口令赋值", re.compile(r"(?:pass" + r"word|sec" + r"ret|tok" + r"en|api[_-]?key|access[_-]?key)\s*[:=]", re.I)),
    ("内部意图词", re.compile(r"DEMO_GUIDE|面试|简历|\b(?:offer|resume|interview)\b", re.I)),
    ("手机号", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")),
    ("邮箱地址", re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")),
    ("身份证号", re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")),
    ("内网 IP", re.compile(r"\b(?:10\.\d{1,3}|172\.(?:1[6-9]|2\d|3[01])|192\.168)\.\d{1,3}\.\d{1,3}\b")),
]

DRIVE_PATTERN = re.compile(r"(?<![A-Za-z])([A-Za-z]):[\\/]")
CODE_TOKEN = re.compile(r"\b[A-Z]{1,4}\d[A-Z0-9]{2,}\b")
CODE_ALLOW = {"SHA256", "SHA1", "MD5SUM", "BASE64", "PDF17", "UTF8", "PYMUPDF2"}
DRIVE_ALLOW = [
    "D:\\test\\tt\\pdf",
    "D:\\your_data",
    "D:\\your_work_folder",
    "D:\\work",
    "D:/",
    "C:\\Windows\\Fonts",
    "/usr/share/fonts",
]

TEXT_EXT = {".md", ".json", ".txt", ".py", ".html", ".gitignore", ".js", ".css"}
TEXT_NAMES = {".gitignore", "LICENSE"}

GLOB_GROUPS = [
    ("docs/demo/replay/", ".json"),
    ("docs/demo/assets/", ".png"),
]


def read_manifest():
    entries = []
    with io.open(MANIFEST, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#"):
                entries.append(line)
    return entries


def is_text(path):
    name = os.path.basename(path)
    ext = os.path.splitext(path)[1].lower()
    return ext in TEXT_EXT or name in TEXT_NAMES


def scan_file(rel):
    source_rel = SOURCE_OVERRIDES.get(rel, rel)
    full = os.path.join(BASE, source_rel.replace("/", os.sep))
    if not os.path.exists(full):
        return [("MISSING", 0, "清单登记但文件不存在：" + source_rel)], []
    if rel in SKIP_SELF or not is_text(path=source_rel):
        return [], []
    text = io.open(full, encoding="utf-8", errors="replace").read()
    hard, warn = [], []
    for lineno, line in enumerate(text.split("\n"), 1):
        for label, pat in HARD_PATTERNS:
            m = pat.search(line)
            if m:
                hard.append((label, lineno, line.strip()[:120]))
        for m in DRIVE_PATTERN.finditer(line):
            start = max(0, m.start() - 26)
            snippet = line[start:m.end() + 26].strip()
            normalized = line.replace("\\\\", "\\")
            allowed = any(a in line or a in normalized for a in DRIVE_ALLOW)
            warn.append(("允许的占位/系统路径" if allowed else "需人工确认的盘符路径", lineno, snippet))
        for m in CODE_TOKEN.finditer(line):
            if m.group(0).upper() in CODE_ALLOW:
                continue
            start = max(0, m.start() - 20)
            warn.append(("公司代码样式 token", lineno, line[start:m.end() + 20].strip()))
    return hard, warn


def build_gitignore_text(entries):
    """由清单生成 public.gitignore 内容（纯函数，便于校验一致性）"""
    lines = [
        "# ============================================================",
        "# 公开仓库白名单：默认忽略一切，仅放行 public_files.txt 清单内的文件",
        "# 由 docs/demo/sensitive_check.py --sync 自动生成，请勿手改",
        "# ============================================================",
        "",
        "*",
        "!*/",
    ]
    body = []
    grouped = set()
    for rel in entries:
        target = None
        for prefix, ext in GLOB_GROUPS:
            if rel.startswith(prefix) and rel.endswith(ext):
                target = prefix + "*" + ext
                break
        if target:
            if target not in grouped:
                grouped.add(target)
                body.append("!" + target)
        else:
            body.append("!/" + rel if "/" not in rel else "!" + rel)
    return "\n".join(lines + body) + "\n"


def sync_gitignore(entries):
    with io.open(PUBLIC_IGNORE, "w", encoding="utf-8") as f:
        f.write(build_gitignore_text(entries))
    print("public.gitignore 已生成：", PUBLIC_IGNORE)
    print("条目数：", len(entries))


def main():
    entries = read_manifest()
    if "--sync" in sys.argv:
        sync_gitignore(entries)
        return

    print("=" * 66)
    print("公开文件敏感扫描 · 清单", len(entries), "个文件")
    print("=" * 66)

    hard_total, warn_total, problems = 0, 0, 0
    for rel in entries:
        hard, warn = scan_file(rel)
        if hard or warn:
            print("\n[" + rel + "]")
        for label, lineno, snippet in hard:
            print("  ✗ 硬违规 [%s] L%d: %s" % (label, lineno, snippet))
            hard_total += 1
        for label, lineno, snippet in warn:
            print("  · 提示  [%s] L%d: %s" % (label, lineno, snippet))
            warn_total += 1
        if any(h[0] == "MISSING" for h in hard):
            problems += 1

    print("\n" + "=" * 66)
    if hard_total == 0:
        print("扫描通过：硬违规 0；提示项 %d（均需人工确认后放行）" % warn_total)
        print("=" * 66)
        return
    print("扫描失败：硬违规 %d，缺失文件 %d" % (hard_total, problems))
    print("=" * 66)
    sys.exit(1)


if __name__ == "__main__":
    main()
