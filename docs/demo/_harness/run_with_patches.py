"""在受控环境中执行目标脚本：拦截 GUI 弹窗 / input() / os.startfile。

用法: python run_with_patches.py <config.json>

config 字段:
  script   要执行的脚本绝对路径（建议为工作区内副本，使 script_dir 指向工作区）
  cwd      执行时的工作目录（脚本的「项目根目录」）
  files    filedialog.askopenfilename 依次返回的绝对路径队列
  dirs     filedialog.askdirectory 依次返回的绝对路径队列
  stdin    input() 依次返回的内容队列
  events   可选，把交互记录写成 JSON 的路径
"""

import builtins
import json
import os
import runpy
import sys
import traceback

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


def install_patches(cfg, events):
    file_queue = list(cfg.get("files", []))
    dir_queue = list(cfg.get("dirs", []))
    stdin_queue = list(cfg.get("stdin", []))

    from tkinter import filedialog, messagebox

    def askopenfilename(**kwargs):
        value = file_queue.pop(0) if file_queue else ""
        events.append({"type": "dialog_file", "title": kwargs.get("title", ""), "value": value})
        print(f"[dialog] 文件选择 -> {value}")
        return value

    def askdirectory(**kwargs):
        value = dir_queue.pop(0) if dir_queue else ""
        events.append({"type": "dialog_dir", "title": kwargs.get("title", ""), "value": value})
        print(f"[dialog] 目录选择 -> {value}")
        return value

    filedialog.askopenfilename = askopenfilename
    filedialog.askdirectory = askdirectory

    def box(kind):
        def handler(*args, **kwargs):
            events.append({"type": kind, "args": [str(a) for a in args]})
            print(f"[{kind}] {' | '.join(str(a) for a in args)}")
            return True
        return handler

    for name in ["showinfo", "showerror", "showwarning", "askyesno", "askokcancel", "askquestion"]:
        setattr(messagebox, name, box(name))

    def fake_input(prompt=""):
        value = stdin_queue.pop(0) if stdin_queue else ""
        events.append({"type": "input", "prompt": str(prompt), "value": value})
        print(f"{prompt}{value}")
        return value

    builtins.input = fake_input

    if hasattr(os, "startfile"):
        def blocked_startfile(path):
            events.append({"type": "startfile", "path": str(path)})
            print(f"[startfile] 已拦截打开 -> {path}")
        os.startfile = blocked_startfile


def main():
    cfg_path = sys.argv[1]
    with open(cfg_path, encoding="utf-8") as f:
        cfg = json.load(f)

    os.chdir(cfg["cwd"])
    # 重置 argv：脚本只看到自己的参数（默认为空），避免把 harness 的配置路径误当业务参数
    sys.argv = [cfg["script"]] + [str(a) for a in cfg.get("argv", [])]
    events = []
    install_patches(cfg, events)

    print("=" * 64)
    print(f"[harness] 执行: {cfg['script']}")
    print(f"[harness] 工作目录: {cfg['cwd']}")
    print("=" * 64)

    code = 0
    try:
        runpy.run_path(cfg["script"], run_name="__main__")
    except SystemExit as e:
        code = int(e.code or 0)
        print(f"[harness] SystemExit({code})")
    except BaseException:
        code = 1
        traceback.print_exc()

    print("=" * 64)
    print(f"[harness] 结束，退出码 = {code}")
    print("=" * 64)

    if cfg.get("events"):
        with open(cfg["events"], "w", encoding="utf-8") as f:
            json.dump(events, f, ensure_ascii=False, indent=2)

    sys.exit(code)


if __name__ == "__main__":
    main()
