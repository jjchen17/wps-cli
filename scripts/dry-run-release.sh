#!/usr/bin/env bash
# wps-cli 发布前 dry-run —— 全部本地，不上传任何远程
#
# 用法：
#   bash scripts/dry-run-release.sh
#
# 验证：
#   1. PyPI 上 wps-cli 包名是否被占用
#   2. 干净 venv 里 build + twine check + pip install + wps version 全跑通
#   3. wheel METADATA 显示正确

set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_ROOT"

PYPI_MIRROR="${PYPI_MIRROR:-https://pypi.tuna.tsinghua.edu.cn/simple}"
TRUSTED="pypi.tuna.tsinghua.edu.cn"
PKG_NAME="wps-cli"

step() { echo -e "\n\033[1;36m[ $* ]\033[0m"; }
fail() { echo -e "\033[1;31mFAIL: $*\033[0m" >&2; exit 1; }

# 0. 包名占用实时检查
step "0/7  PyPI 包名占用检查"
HTTP=$(curl -s -o /dev/null -w "%{http_code}" "https://pypi.org/pypi/${PKG_NAME}/json" || echo "000")
case "$HTTP" in
  404) echo "  OK  ${PKG_NAME} 未被占用，可发布";;
  200) echo "  WARN  ${PKG_NAME} 已存在，确认是你自己的包后再继续";;
  *)   echo "  WARN  PyPI 返回 ${HTTP}（可能网络问题），手动到浏览器确认";;
esac

# 1. 清理旧产物
step "1/7  清理 dist/ build/ *.egg-info"
rm -rf dist/ build/ src/*.egg-info

# 2. 干净 venv（与系统 Python 隔离，模拟用户首次安装）
step "2/7  创建干净 venv"
VENV_DIR="$(mktemp -d)/venv"
python -m venv "$VENV_DIR"
if [[ -f "$VENV_DIR/Scripts/python.exe" ]]; then
  PY="$VENV_DIR/Scripts/python.exe"
  PIP="$VENV_DIR/Scripts/pip.exe"
  WPS_BIN="$VENV_DIR/Scripts/wps.exe"
else
  PY="$VENV_DIR/bin/python"
  PIP="$VENV_DIR/bin/pip"
  WPS_BIN="$VENV_DIR/bin/wps"
fi
"$PIP" install -q --upgrade pip -i "$PYPI_MIRROR" --trusted-host "$TRUSTED"
"$PIP" install -q build twine -i "$PYPI_MIRROR" --trusted-host "$TRUSTED"

# 3. python -m build
step "3/7  python -m build"
"$PY" -m build
ls -la dist/

# 4. twine check —— 校验 README 在 PyPI 上能渲染
step "4/7  twine check"
"$PY" -m twine check dist/* || fail "twine check 失败 —— 通常是 README 渲染问题"

# 5. 干净 venv 装 wheel
step "5/7  pip install dist/*.whl 在干净 venv"
WHL=$(ls dist/*.whl | head -1)
"$PIP" install "$WHL" -i "$PYPI_MIRROR" --trusted-host "$TRUSTED"

# 6. wps version + wps doctor
step "6/7  wps version & wps --help"
"$WPS_BIN" version || fail "wps version 失败"
"$WPS_BIN" --help > /dev/null || fail "wps --help 失败"
"$WPS_BIN" writer --help > /dev/null || fail "wps writer --help 失败"
"$WPS_BIN" calc --help > /dev/null || fail "wps calc --help 失败"
echo "  ----- doctor --report 输出 -----"
"$WPS_BIN" doctor --report 2>&1 | head -15 || true
echo "  --------------------------------"

# 7. wheel 元数据 sanity check
step "7/7  wheel 元数据"
"$PY" -c "
import zipfile, re, sys
whl = sys.argv[1]
with zipfile.ZipFile(whl) as z:
    meta_name = next(n for n in z.namelist() if n.endswith('METADATA'))
    meta = z.read(meta_name).decode('utf-8')
for line in meta.splitlines():
    if re.match(r'^(Name|Version|License|Requires-Python|Requires-Dist|Summary):', line):
        print('  ', line)
" "$WHL"

echo -e "\n\033[1;32mALL CHECKS PASSED.\033[0m"
echo ""
echo "下一步——TestPyPI 试发（建议先做这一步）："
echo "  git tag v0.1.0a1-test && git push origin v0.1.0a1-test"
echo ""
echo "正式发 PyPI："
echo "  git tag v0.1.0a1 && git push origin v0.1.0a1"
