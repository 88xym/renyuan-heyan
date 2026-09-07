# -*- coding: utf-8 -*-
"""重建「开始核验.bat」。

bat 文件必须为 GBK(ANSI) 编码 + CRLF 行尾，否则中文 Windows 的
cmd.exe 会乱码报错。本脚本从 UTF-8 源文件 tools/bat_source.bat
重建 开始核验.bat（GBK + CRLF）。

修改 bat 提示文字的正确流程：
  1. 用任意编辑器修改 tools/bat_source.bat（UTF-8）
  2. 运行: python tools/rebuild_bat.py
  3. 生成的 开始核验.bat 即为可双击运行的版本

注意：不要直接用编辑器修改 开始核验.bat，保存时可能破坏 GBK 编码。
"""

from __future__ import annotations

from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SRC = PROJECT_ROOT / "tools" / "bat_source.bat"
DST = PROJECT_ROOT / "开始核验.bat"


def main() -> int:
    if not SRC.exists():
        print(f"[错误] 源文件不存在: {SRC}", file=sys.stderr)
        return 1
    content = SRC.read_text(encoding="utf-8")
    crlf_content = "\r\n".join(content.splitlines()) + "\r\n"
    DST.write_bytes(crlf_content.encode("gbk"))
    print(f"已重建: {DST}")
    print(f"编码: GBK(ANSI) + CRLF 行尾，可在中文 Windows 双击运行。")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
