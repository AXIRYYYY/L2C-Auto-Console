"""一键运行公开的代码样本（隔离工作区，仓库根零污染）。

对齐脚本的原始用法：「把脚本复制到数据文件夹再运行」——
脚本会把「自己所在目录」当作数据目录。
本工具按同样布局把脚本与样例数据放进隔离工作区后再运行：

    docs/demo/.sample_run/<样例>/
    ├── xxx.py                   ← 脚本与数据放在同一层
    ├── <样例数据文件>             ← 脚本要读取的数据
    └── src/ app/ ...            ← 需要整目录匹配的样例（资料包整合）

用法:
  python docs/demo/run_sample.py                   # 列出可运行的样例
  python docs/demo/run_sample.py <样例名>           # 运行并保留工作区（便于查看产物）
  python docs/demo/run_sample.py <样例名> --clean   # 运行后删除工作区（一次性）
"""

import os
import shutil
import subprocess
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(os.path.dirname(HERE))
SAMPLE = os.path.join(BASE, "sample_data")
WORK = os.path.join(HERE, ".sample_run")

SAMPLES = {
    "voucher_compare": {
        "title": "凭证比对（frozenset 内容签名 + Jaccard 四分类差异）",
        "script": "accounting/compare_vouchers_advanced.py",
        "files": ["归档/凭证清单_修改前.xlsx", "归档/凭证清单_修改后.xlsx"],
        "trees": [],
        "argv": ["{ws}/凭证清单_修改前.xlsx", "{ws}/凭证清单_修改后.xlsx"],
    },
    "voucher_split": {
        "title": "凭证 PDF 拆分（年月 + 基础凭证号分组，文件名带页数）",
        "script": "accounting/split_vouchers_凭证PDF拆分.py",
        "files": ["归档/记账凭证-202601.pdf"],
        "trees": [],
        "argv": ["{ws}/记账凭证-202601.pdf"],
    },
    "namelist_pack": {
        "title": "月底资料包整合（归集 → 关联申请单 → 合并 PDF → 对账表）",
        "script": "integration/namelistget_feishu_v9.py",
        "files": ["月底/namelist.xlsx"],
        "trees": [("月底/附件", "src"), ("月底/申请单", "app")],
        "argv": ["{ws}/app", "{ws}/src"],
    },
    "bank_split": {
        "title": "银行流水按凭证拆分（唯一文件约束 + 行/页数一致性校验）",
        "script": "accounting/bankStatementSplitterBasedOnJournal.py",
        "files": ["归档/银行流水-202601.pdf", "归档/日记账202601.xlsx"],
        "trees": [],
        "argv": [],
    },
}


def list_samples():
    print("可运行的公开样例：\n")
    for slug, spec in SAMPLES.items():
        print("  %-16s %s" % (slug, spec["title"]))
    print("\n用法：python docs/demo/run_sample.py <样例名> [--clean]")


def snapshot(root):
    found = {}
    for r, _dirs, files in os.walk(root):
        for f in files:
            p = os.path.join(r, f)
            try:
                found[p] = os.path.getsize(p)
            except OSError:
                pass
    return found


def stage(spec, ws):
    # 脚本与样例数据放在同一层（脚本把「所在目录」当数据目录）
    script_dst = os.path.join(ws, os.path.basename(spec["script"]))
    shutil.copy2(os.path.join(BASE, spec["script"].replace("/", os.sep)), script_dst)
    for rel in spec["files"]:
        shutil.copy2(os.path.join(SAMPLE, rel.replace("/", os.sep)), os.path.join(ws, os.path.basename(rel)))
    for src_rel, dst_name in spec["trees"]:
        shutil.copytree(os.path.join(SAMPLE, src_rel.replace("/", os.sep)), os.path.join(ws, dst_name))
    return script_dst


def main():
    args = [a for a in sys.argv[1:] if a.strip()]
    if not args or args[0] in ("--list", "-l", "--help", "-h"):
        list_samples()
        return 0

    slug = args[0]
    clean = "--clean" in args[1:]
    if slug not in SAMPLES:
        print("未知样例：%s\n" % slug)
        list_samples()
        return 1
    if not os.path.isdir(SAMPLE):
        print("找不到样例数据目录：%s\n请先运行：python sample_data/generate_sample_data.py" % SAMPLE)
        return 1

    spec = SAMPLES[slug]
    ws = os.path.join(WORK, slug)
    shutil.rmtree(ws, ignore_errors=True)
    os.makedirs(ws)

    script_dst = stage(spec, ws)
    before = snapshot(ws)
    argv = [a.format(ws=ws.replace(os.sep, "/")) for a in spec["argv"]]

    print("=" * 66)
    print("样例：%s" % spec["title"])
    print("脚本：%s" % spec["script"])
    print("工作区：%s" % ws)
    print("（脚本把「所在目录」当数据目录，这里即工作区）")
    print("=" * 66)

    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.run(
        [sys.executable, script_dst] + argv,
        cwd=ws, env=env, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=300,
    )
    log = proc.stdout or ""
    if proc.stderr and proc.stderr.strip():
        log += "\n[stderr]\n" + proc.stderr
    print(log.rstrip())
    print("=" * 66)
    print("退出码：%d" % proc.returncode)

    after = snapshot(ws)
    new_files = sorted(p for p in after if p not in before)
    if new_files:
        print("\n产物（%d 个）：" % len(new_files))
        for p in new_files:
            print("  %8.1f KB  %s" % (after[p] / 1024, os.path.relpath(p, ws)))
    else:
        print("\n（没有生成新文件）")

    if clean:
        shutil.rmtree(ws, ignore_errors=True)
        print("\n工作区已删除（--clean）")
    else:
        print("\n工作区保留在：%s" % ws)

    return proc.returncode


if __name__ == "__main__":
    sys.exit(main())
