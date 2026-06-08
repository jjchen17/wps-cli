# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""WPS COM 注册表诊断与修复模块

解决 Issue #8: WPS Office 12.x 新版安装后 LocalServer32 缺失导致 COM 不可用。

提供以下能力：
- 读取 Windows 注册表检查 ProgID → CLSID → LocalServer32 链
- 检测 Python/pywin32 位数是否与 WPS 位数匹配
- 探测 WPS 安装路径和 ksomgr.exe 位置
- 通过 ksomgr 自动修复 COM 注册
- 生成结构化诊断报告
"""

from __future__ import annotations

import logging
import os
import platform
import shutil
import struct
import subprocess
import sys
from dataclasses import dataclass, field

_logger = logging.getLogger("wps_cli.com_diagnostics")

try:
    import winreg
except ImportError:
    winreg = None  # type: ignore[assignment]


# ══════════════════════════════════════════════════════════════════
# 数据模型
# ══════════════════════════════════════════════════════════════════


@dataclass(frozen=True)
class ProgIDRegInfo:
    """单个 ProgID 的注册表信息"""

    prog_id: str
    clsid: str | None = None
    local_server32: str | None = None
    local_server_exists: bool = False
    registry_view: str = "native"
    error: str = ""


@dataclass(frozen=True)
class ComDiagnosticResult:
    """整体 COM 诊断结果"""

    python_version: str = ""
    python_bits: int = 64
    pywin32_version: str = ""
    platform_name: str = ""
    is_admin: bool = False
    wps_install_path: str | None = None
    wps_bits: int | None = None
    wps_version: str | None = None
    writer_results: list[ProgIDRegInfo] = field(default_factory=list)
    calc_results: list[ProgIDRegInfo] = field(default_factory=list)
    impress_results: list[ProgIDRegInfo] = field(default_factory=list)
    bitness_match: bool = True
    has_working_progid: bool = False
    can_auto_fix: bool = False
    ksomgr_path: str | None = None
    issues: list[str] = field(default_factory=list)
    fix_suggestions: list[str] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════
# 基础工具
# ══════════════════════════════════════════════════════════════════


def _is_windows() -> bool:
    return sys.platform == "win32"


def get_python_bitness() -> int:
    """获取当前 Python 解释器的位数（32 或 64）"""
    return struct.calcsize("P") * 8


def get_pywin32_version() -> str:
    """获取 pywin32 版本号"""
    try:
        import importlib.metadata as md

        return md.version("pywin32")
    except Exception:
        return "未安装"


def is_running_as_admin() -> bool:
    """检查是否以管理员权限运行"""
    if not _is_windows():
        return False
    try:
        import ctypes

        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════
# 注册表检测
# ══════════════════════════════════════════════════════════════════


def _read_reg_default_value(key_path: str, sam_flags: int = 0) -> str | None:
    """读取 HKCR 下指定键的默认值"""
    if winreg is None:
        return None
    access = winreg.KEY_READ | sam_flags
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, access) as key:
            value, _ = winreg.QueryValueEx(key, "")
            return str(value) if value else None
    except (FileNotFoundError, OSError):
        return None


def _check_file_exists(path: str | None) -> bool:
    """检查 LocalServer32 指向的文件是否存在（处理命令行参数）"""
    if not path:
        return False
    exe_path = path.strip().strip('"')
    if exe_path.startswith('"'):
        end = exe_path.find('"', 1)
        exe_path = exe_path[1:end] if end > 0 else exe_path[1:]
    else:
        space = exe_path.find(" ")
        if space > 0:
            exe_path = exe_path[:space]
    return os.path.isfile(exe_path)


def check_progid_registry(prog_id: str) -> list[ProgIDRegInfo]:
    """全面检查 ProgID 在注册表中的状态

    检查范围：原生视图 + WOW6432Node (32bit) + 64bit 显式视图
    """
    if winreg is None:
        return [ProgIDRegInfo(prog_id=prog_id, error="非 Windows 平台")]

    results: list[ProgIDRegInfo] = []
    checked_clsids: set[str] = set()

    view_configs: list[tuple[str, int, str]] = [
        ("native", 0, prog_id),
        ("32bit", winreg.KEY_WOW64_32KEY, f"WOW6432Node\\{prog_id}"),
        ("64bit", winreg.KEY_WOW64_64KEY, prog_id),
    ]

    for view_name, sam_flags, reg_prog_path in view_configs:
        if view_name == "64bit" and get_python_bitness() == 64:
            continue
        if view_name == "32bit" and get_python_bitness() == 32:
            continue

        clsid = _read_reg_default_value(f"{reg_prog_path}\\CLSID", sam_flags)
        if clsid and clsid in checked_clsids:
            continue
        if clsid:
            checked_clsids.add(clsid)

        if not clsid:
            results.append(ProgIDRegInfo(
                prog_id=prog_id, registry_view=view_name,
                error=f"注册表键不存在: HKCR\\{reg_prog_path}\\CLSID",
            ))
            continue

        ls_path = f"CLSID\\{clsid}\\LocalServer32"
        local_server = _read_reg_default_value(ls_path, sam_flags)
        if not local_server:
            ls_wow_path = f"WOW6432Node\\CLSID\\{clsid}\\LocalServer32"
            local_server = _read_reg_default_value(ls_wow_path, sam_flags)

        exists = _check_file_exists(local_server)
        error = ""
        if not local_server:
            error = f"CLSID {clsid} 缺少 LocalServer32 键"
        elif not exists:
            error = "LocalServer32 指向的路径不存在"

        results.append(ProgIDRegInfo(
            prog_id=prog_id, clsid=clsid,
            local_server32=local_server, local_server_exists=exists,
            registry_view=view_name, error=error,
        ))

    if not results:
        results.append(ProgIDRegInfo(prog_id=prog_id, error="所有注册表视图均未找到此 ProgID"))

    return results


# ══════════════════════════════════════════════════════════════════
# WPS 安装检测
# ══════════════════════════════════════════════════════════════════


def detect_wps_install_path(app_type: str = "writer") -> str | None:
    """检测 WPS 安装路径（注册表 → 常见目录 → PATH）"""
    if not _is_windows():
        return None

    exe_names: dict[str, str] = {"writer": "wps.exe", "calc": "et.exe", "impress": "wpp.exe"}
    exe_name = exe_names.get(app_type, "wps.exe")

    # 方法 1: 注册表
    if winreg is not None:
        reg_keys = [
            r"SOFTWARE\Kingsoft\WPS Office",
            r"SOFTWARE\WOW6432Node\Kingsoft\WPS Office",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe",
        ]
        for reg_key in reg_keys:
            for access_flag in (0, winreg.KEY_WOW64_32KEY):
                try:
                    with winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE, reg_key, 0, winreg.KEY_READ | access_flag,
                    ) as key:
                        install_path, _ = winreg.QueryValueEx(key, "")
                        if install_path:
                            candidate = os.path.join(str(install_path).rstrip("\\"), exe_name)
                            if os.path.isfile(candidate):
                                return candidate
                except (FileNotFoundError, OSError):
                    continue

    # 方法 2: 常见目录
    known_dirs = [
        os.environ.get("PROGRAMFILES", r"C:\Program Files"),
        os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
    ]
    for base_dir in known_dirs:
        if not base_dir:
            continue
        base_wps = os.path.join(base_dir, "Kingsoft", "WPS Office")
        if os.path.isdir(base_wps):
            for root, _dirs, files in os.walk(base_wps):
                if exe_name in files:
                    return os.path.join(root, exe_name)

    # 方法 3: PATH
    which = shutil.which(exe_name)
    if which:
        return which

    return None


def detect_wps_bitness(wps_exe_path: str | None) -> int | None:
    """检测 WPS 可执行文件的位数（PE 头解析）"""
    if not wps_exe_path or not os.path.isfile(wps_exe_path) or not _is_windows():
        return None

    # 路径启发式
    path_lower = wps_exe_path.lower()
    if "program files (x86)" in path_lower:
        return 32
    if "program files" in path_lower:
        return 64

    # PE 头解析
    try:
        with open(wps_exe_path, "rb") as f:
            f.seek(0x3C)
            pe_offset_bytes = f.read(4)
            if len(pe_offset_bytes) < 4:
                return None
            pe_offset = struct.unpack("<I", pe_offset_bytes)[0]
            f.seek(pe_offset)
            pe_sig = f.read(4)
            if pe_sig != b"PE\x00\x00":
                return None
            coff = f.read(20)
            if len(coff) < 20:
                return None
            machine = struct.unpack("<H", coff[0:2])[0]
            if machine in (0x8664, 0xAA64):
                return 64
            if machine == 0x014C:
                return 32
    except (OSError, struct.error, IndexError):
        pass

    return None


def find_ksomgr() -> str | None:
    """查找 ksomgr.exe 的路径"""
    known_paths = [
        r"C:\Program Files\Kingsoft\WPS Office\ksomgr.exe",
        r"C:\Program Files (x86)\Kingsoft\WPS Office\ksomgr.exe",
    ]
    for known_path in known_paths:
        if os.path.isfile(known_path):
            return known_path

    for app_type in ("writer", "calc", "impress"):
        wps_path = detect_wps_install_path(app_type)
        if wps_path:
            wps_dir = os.path.dirname(wps_path)
            for candidate_name in ("ksomgr.exe", "ksomisc.exe", "ksolaunch.exe"):
                candidate = os.path.join(wps_dir, candidate_name)
                if os.path.isfile(candidate):
                    return candidate
                parent_candidate = os.path.join(os.path.dirname(wps_dir), candidate_name)
                if os.path.isfile(parent_candidate):
                    return parent_candidate
    return None


# ══════════════════════════════════════════════════════════════════
# COM 连接探测
# ══════════════════════════════════════════════════════════════════


def probe_progids(app_type: str) -> list[str]:
    """探测指定应用类型的所有候选 ProgID，返回可用的列表"""
    from wps_cli.consts import COM_PROGID_CANDIDATES

    candidates = COM_PROGID_CANDIDATES.get(app_type, [])
    working: list[str] = []

    for prog_id in candidates:
        try:
            import win32com.client

            app = win32com.client.Dispatch(prog_id)
            try:
                _ = app.Version
                working.append(prog_id)
            finally:
                try:
                    app.Quit()
                except Exception:
                    pass
        except Exception:
            continue

    return working


# ══════════════════════════════════════════════════════════════════
# 自动修复
# ══════════════════════════════════════════════════════════════════


def run_ksomgr_register(ksomgr_path: str | None = None) -> tuple[bool, str]:
    """运行 ksomgr -regserver 重新注册 WPS COM 组件"""
    if not _is_windows():
        return False, "非 Windows 平台"

    if ksomgr_path is None:
        ksomgr_path = find_ksomgr()

    if not ksomgr_path or not os.path.isfile(ksomgr_path):
        return False, f"未找到注册修复工具: {ksomgr_path or '(null)'}"

    cmd = [ksomgr_path, "-regserver"]
    _logger.info("执行: %s", " ".join(cmd))

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        output = result.stdout.strip() or result.stderr.strip()
        if result.returncode == 0:
            return True, output or "注册成功"
        return False, f"返回码 {result.returncode}: {output}"
    except subprocess.TimeoutExpired:
        return False, "执行超时（30秒）"
    except OSError as exc:
        return False, f"操作系统错误: {exc}"


def attempt_com_fix(app_type: str | None = None) -> tuple[bool, list[str]]:
    """尝试修复 COM 注册问题

    修复步骤：
    1. 以管理员身份运行 ksomgr -regserver
    2. 重新探测 ProgID

    Returns:
        (是否成功, 修复日志列表)
    """
    logs: list[str] = []

    if not _is_windows():
        logs.append("非 Windows 平台，无需修复")
        return False, logs

    if not is_running_as_admin():
        logs.append("⚠ 建议以管理员身份运行 wps doctor --fix")
        logs.append("  否则可能无权限写入注册表")

    # Step 1: 尝试 ksomgr/ksomisc/ksolaunch
    ksomgr_path = find_ksomgr()
    if ksomgr_path:
        logs.append(f"找到注册工具: {ksomgr_path}")
        success, output = run_ksomgr_register(ksomgr_path)
        logs.append(f"结果: {output}")
        if success:
            working = probe_progids(app_type or "writer")
            if working:
                logs.append(f"修复成功！可用 ProgID: {', '.join(working)}")
                return True, logs
            logs.append("注册工具执行成功但 COM 仍不可用，请重启后重试")
            return False, logs
        logs.append("注册工具执行失败")
    else:
        logs.append("未找到 ksomgr.exe/ksomisc.exe，无法自动修复")

    # Step 2: 手动修复建议
    wps_path = detect_wps_install_path(app_type or "writer")
    if wps_path:
        logs.append(f"\n手动修复步骤（WPS 位于: {wps_path}）:")
        logs.append(f"  1. 以管理员身份运行: \"{wps_path}\" /regserver")
        wps_dir = os.path.dirname(wps_path)
        for exe in ("ksolaunch.exe", "ksomisc.exe"):
            p = os.path.join(wps_dir, exe)
            if os.path.isfile(p):
                logs.append(f"  2. 或运行: \"{p}\" -regserver")
        logs.append("  3. 以管理员身份打开一次 WPS Office（触发自注册）")
        logs.append("  4. 如果以上均无效，请重新安装 WPS Office")
        logs.append(f"  5. 确保 Python/pywin32 位数 ({get_python_bitness()}位) 与 WPS 位数一致")

    return False, logs


# ══════════════════════════════════════════════════════════════════
# 综合诊断报告
# ══════════════════════════════════════════════════════════════════


def run_full_diagnostics() -> ComDiagnosticResult:
    """运行完整的 COM 注册诊断"""
    result = ComDiagnosticResult()
    result.python_version = sys.version.split()[0]
    result.python_bits = get_python_bitness()
    result.platform_name = platform.platform()
    result.pywin32_version = get_pywin32_version()
    result.is_admin = is_running_as_admin()

    if not _is_windows():
        result.issues.append("非 Windows 平台，COM 自动化不可用")
        return result

    result.wps_install_path = detect_wps_install_path("writer")
    if result.wps_install_path:
        result.wps_bits = detect_wps_bitness(result.wps_install_path)
    else:
        result.issues.append("未检测到 WPS 安装路径")

    if result.wps_bits and result.python_bits != result.wps_bits:
        result.bitness_match = False
        result.issues.append(
            f"位数不匹配: Python {result.python_bits}位 vs WPS {result.wps_bits}位"
        )

    from wps_cli.consts import COM_PROGID_CANDIDATES

    for app_type, attr in [("writer", "writer_results"), ("calc", "calc_results"), ("impress", "impress_results")]:
        all_infos: list[ProgIDRegInfo] = []
        for prog_id in COM_PROGID_CANDIDATES.get(app_type, []):
            infos = check_progid_registry(prog_id)
            all_infos.extend(infos)
            for info in infos:
                if info.local_server_exists and not info.error:
                    result.has_working_progid = True
        setattr(result, attr, all_infos)

    result.ksomgr_path = find_ksomgr()
    result.can_auto_fix = result.ksomgr_path is not None

    if not result.has_working_progid:
        result.fix_suggestions.append("运行 'wps doctor --fix' 自动修复 COM 注册")
        result.fix_suggestions.append("以管理员身份运行一次 WPS Office（触发自注册）")
        if result.ksomgr_path:
            result.fix_suggestions.append(f"手动运行: \"{result.ksomgr_path}\" -regserver")

    if not result.bitness_match:
        result.fix_suggestions.append("安装与 WPS 位数匹配的 Python 和 pywin32")

    return result
