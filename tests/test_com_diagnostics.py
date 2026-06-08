# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""COM 诊断模块测试"""

from __future__ import annotations

from wps_cli.exceptions import WpsNotFoundError
from wps_cli.services.com_diagnostics import (
    ComDiagnosticResult,
    ProgIDRegInfo,
    check_progid_registry,
    detect_wps_bitness,
    find_ksomgr,
    get_python_bitness,
    get_pywin32_version,
)

# ══════════════════════════════════════════════════════════════════
# TestGetPythonBitness
# ══════════════════════════════════════════════════════════════════


class TestGetPythonBitness:
    """测试 Python 位数检测"""

    def test_python_bitness_returns_32_or_64(self):
        """返回位数必须是 32 或 64"""
        bits = get_python_bitness()
        assert bits in (32, 64), f"expected 32 or 64, got {bits}"


# ══════════════════════════════════════════════════════════════════
# TestGetPywin32Version
# ══════════════════════════════════════════════════════════════════


class TestGetPywin32Version:
    """测试 pywin32 版本获取"""

    def test_pywin32_version_returns_string(self):
        """返回字符串类型"""
        version = get_pywin32_version()
        assert isinstance(version, str)


# ══════════════════════════════════════════════════════════════════
# TestProgIDRegInfo
# ══════════════════════════════════════════════════════════════════


class TestProgIDRegInfo:
    """测试 ProgIDRegInfo 数据类"""

    def test_create_info(self):
        """创建 ProgIDRegInfo 对象"""
        info = ProgIDRegInfo(prog_id="WPS.Application")
        assert info.prog_id == "WPS.Application"
        assert info.clsid is None
        assert info.local_server32 is None
        assert info.local_server_exists is False
        assert info.registry_view == "native"
        assert info.error == ""

    def test_create_info_with_error(self):
        """带错误信息的创建"""
        info = ProgIDRegInfo(prog_id="X.ProgID", error="注册表键不存在")
        assert info.error == "注册表键不存在"

    def test_info_with_local_server(self):
        """带 LocalServer32 的创建"""
        info = ProgIDRegInfo(
            prog_id="WPS.Application",
            clsid="{GUID}",
            local_server32=r"C:\WPS\wps.exe",
            local_server_exists=True,
        )
        assert info.clsid == "{GUID}"
        assert info.local_server32 == r"C:\WPS\wps.exe"
        assert info.local_server_exists is True

    def test_info_with_registry_view(self):
        """指定 registry_view 的创建"""
        info = ProgIDRegInfo(prog_id="WPS.Application", registry_view="32bit")
        assert info.registry_view == "32bit"


# ══════════════════════════════════════════════════════════════════
# TestComDiagnosticResult
# ══════════════════════════════════════════════════════════════════


class TestComDiagnosticResult:
    """测试 ComDiagnosticResult 数据类"""

    def test_create_result(self):
        """创建空的诊断结果（全默认值）"""
        r = ComDiagnosticResult()
        assert r.python_bits == 64
        assert r.wps_install_path is None
        assert r.wps_bits is None
        assert r.writer_results == []
        assert r.calc_results == []
        assert r.impress_results == []
        assert r.issues == []
        assert r.fix_suggestions == []
        assert r.has_working_progid is False
        assert r.can_auto_fix is False
        assert r.bitness_match is True

    def test_create_result_with_issues(self):
        """创建带问题的诊断结果"""
        r = ComDiagnosticResult(
            python_bits=32,
            wps_install_path=r"C:\WPS\wps.exe",
            wps_bits=64,
            bitness_match=False,
            issues=["位数不匹配"],
            fix_suggestions=["安装匹配的 Python"],
        )
        assert r.python_bits == 32
        assert r.wps_bits == 64
        assert r.bitness_match is False
        assert len(r.issues) == 1
        assert len(r.fix_suggestions) == 1

    def test_add_issue_and_fix(self):
        """可以通过列表操作添加问题和建议"""
        r = ComDiagnosticResult()
        r.issues.append("测试问题")
        r.fix_suggestions.append("测试建议")
        assert "测试问题" in r.issues
        assert "测试建议" in r.fix_suggestions

    def test_result_with_writer_infos(self):
        """结果包含 writer 的 ProgIDRegInfo 列表"""
        info = ProgIDRegInfo(prog_id="WPS.Application")
        r = ComDiagnosticResult(writer_results=[info])
        assert len(r.writer_results) == 1
        assert r.writer_results[0].prog_id == "WPS.Application"


# ══════════════════════════════════════════════════════════════════
# TestWpsNotFoundErrorEnhanced
# ══════════════════════════════════════════════════════════════════


class TestWpsNotFoundErrorEnhanced:
    """测试增强后的 WpsNotFoundError"""

    def test_error_with_progids(self):
        """构造带 tried_progids 的错误"""
        e = WpsNotFoundError("writer", tried_progids=["A", "B"])
        assert "已尝试的 ProgID: A, B" in str(e)
        assert e.exit_code == 10

    def test_error_with_registry_info(self):
        """带 reg_info 的错误消息包含注册表诊断信息"""
        e = WpsNotFoundError("calc", registry_info="CLSID 缺少 LocalServer32")
        assert "注册表诊断" in str(e)
        assert "CLSID 缺少 LocalServer32" in str(e)

    def test_error_with_fix_hint(self):
        """带 fix_hint 的错误消息包含修复建议"""
        e = WpsNotFoundError("impress", fix_hint="请重新安装 WPS")
        assert "修复建议" in str(e)
        assert "请重新安装 WPS" in str(e)

    def test_error_contains_debug_info(self):
        """没有额外参数时，错误消息包含默认修复建议"""
        e = WpsNotFoundError("writer")
        assert "wps doctor --fix" in str(e)
        assert "以管理员身份" in str(e)

    def test_error_with_all_params(self):
        """同时传入 tried_progids、registry_info 和 fix_hint"""
        e = WpsNotFoundError(
            "writer",
            tried_progids=["Kwps.Application", "WPS.Application"],
            registry_info="缺少 LocalServer32 键",
            fix_hint="运行 ksomgr -regserver",
        )
        msg = str(e)
        assert "Kwps.Application" in msg
        assert "缺少 LocalServer32 键" in msg
        assert "ksomgr -regserver" in msg

    def test_error_suggestion(self):
        """suggestion 字段指向修复命令"""
        e = WpsNotFoundError("writer")
        assert "wps doctor --fix" in e.suggestion

    def test_error_context_contains_progids(self):
        """context 中包含 tried_progids"""
        e = WpsNotFoundError("writer", tried_progids=["A"])
        assert "tried_progids" in e.context
        assert e.context["tried_progids"] == ["A"]


# ══════════════════════════════════════════════════════════════════
# TestCheckProgidRegistryNonWindows
# ══════════════════════════════════════════════════════════════════


class TestCheckProgidRegistryNonWindows:
    """测试非 Windows 平台的注册表检查（使用 monkeypatch）"""

    def test_returns_error_on_non_windows(self, monkeypatch):
        """非 Windows 平台返回 error 信息"""
        import wps_cli.services.com_diagnostics as mod

        monkeypatch.setattr(mod, "winreg", None)
        results = check_progid_registry("WPS.Application")
        assert len(results) == 1
        assert "非 Windows" in results[0].error

    def test_returns_list(self, monkeypatch):
        """返回 list 类型"""
        import wps_cli.services.com_diagnostics as mod

        monkeypatch.setattr(mod, "winreg", None)
        results = check_progid_registry("WPS.Application")
        assert isinstance(results, list)
        assert all(isinstance(r, ProgIDRegInfo) for r in results)


# ══════════════════════════════════════════════════════════════════
# TestDetectWpsBitness
# ══════════════════════════════════════════════════════════════════


class TestDetectWpsBitness:
    """测试 WPS 位数检测（纯逻辑，无需真实文件）"""

    def test_none_path_returns_none(self):
        """None 输入返回 None"""
        assert detect_wps_bitness(None) is None

    def test_nonexistent_path_returns_none(self):
        """不存在的路径返回 None"""
        result = detect_wps_bitness(r"C:\no\such\path\wps.exe")
        assert result is None


# ══════════════════════════════════════════════════════════════════
# TestFindKsomgrNotExist
# ══════════════════════════════════════════════════════════════════


class TestFindKsomgrNotExist:
    """测试无 ksomgr 时的行为（mock 掉安装路径检测）"""

    def test_returns_none_when_no_ksomgr(self, monkeypatch):
        """找不到 WPS 安装路径且无 ksomgr 文件时返回 None"""
        monkeypatch.setattr(
            "wps_cli.services.com_diagnostics.detect_wps_install_path",
            lambda app_type=None: None,
        )
        monkeypatch.setattr(
            "wps_cli.services.com_diagnostics.os.path.isfile",
            lambda path: False,
        )
        assert find_ksomgr() is None

    def test_returns_none_when_wps_found_but_no_ksomgr(self, monkeypatch):
        """找到 WPS 但目录下无 ksomgr 系列文件时返回 None"""
        monkeypatch.setattr(
            "wps_cli.services.com_diagnostics.detect_wps_install_path",
            lambda app_type=None: r"C:\WPS\wps.exe",
        )
        monkeypatch.setattr(
            "wps_cli.services.com_diagnostics.os.path.isfile",
            lambda path: False,
        )
        assert find_ksomgr() is None
