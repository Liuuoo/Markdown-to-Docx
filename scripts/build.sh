#!/usr/bin/env bash
# md2thesis: 把 Markdown 转成中文学位论文风格的 docx(纯 Python, 不用 pandoc)。
#
# 用法:
#   build.sh <input.md> [-o out.docx] [--title T] [--author A] [--pdf]
set -euo pipefail

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
BUILD_PY="$SKILL_DIR/scripts/build_docx.py"

INPUT=""; OUT=""; WANT_PDF=0; TITLE=""; AUTHOR=""
while [[ $# -gt 0 ]]; do
  case "$1" in
    --pdf) WANT_PDF=1; shift;;
    -o) OUT="$2"; shift 2;;
    --title) TITLE="$2"; shift 2;;
    --author) AUTHOR="$2"; shift 2;;
    *) INPUT="$1"; shift;;
  esac
done

[[ -z "$INPUT" ]] && { echo "错误: 需提供输入 .md 文件" >&2; exit 1; }
[[ ! -f "$INPUT" ]] && { echo "错误: 找不到文件 $INPUT" >&2; exit 1; }
python3 -c "import docx" 2>/dev/null || { echo "缺少 python-docx。安装: pip3 install --user python-docx" >&2; exit 1; }

[[ -z "$OUT" ]] && OUT="${INPUT%.*}.docx"

ARGS=("$INPUT" -o "$OUT")
[[ -n "$TITLE" ]] && ARGS+=(--title "$TITLE")
[[ -n "$AUTHOR" ]] && ARGS+=(--author "$AUTHOR")

echo "[1/2] 解析 Markdown 并组装 docx..."
python3 "$BUILD_PY" "${ARGS[@]}"

if [[ "$WANT_PDF" -eq 1 ]]; then
  echo "[2/2] 转 PDF..."
  PDF="${OUT%.*}.pdf"
  SOFFICE="$(command -v soffice || true)"
  [[ -z "$SOFFICE" && -x /Applications/LibreOffice.app/Contents/MacOS/soffice ]] && SOFFICE=/Applications/LibreOffice.app/Contents/MacOS/soffice
  if [[ -n "$SOFFICE" ]]; then
    "$SOFFICE" --headless --convert-to pdf --outdir "$(dirname "$OUT")" "$OUT" >/dev/null
    echo "完成 -> $PDF"
  else
    echo "无法转 PDF: 未装 LibreOffice。安装: brew install --cask libreoffice" >&2
    exit 1
  fi
fi
