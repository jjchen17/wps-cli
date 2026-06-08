# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""WPS COM 后端实现 -- 多 ProgID 回退 + 增强诊断

WPS 12.x 不再注册 K 前缀 ProgID（如 KWPS.Application），
本模块依次尝试多个候选 ProgID，失败时自动运行注册表诊断并提供修复建议。
"""

from __future__ import annotations

import logging
from typing import Any

from wps_cli.backends.base import ComBackend
from wps_cli.exceptions import WpsNotFoundError
from wps_cli.services.com_diagnostics import (
    check_progid_registry,
    detect_wps_bitness,
    detect_wps_install_path,
    find_ksomgr,
    get_python_bitness,
)

_logger = logging.getLogger("wps_cli.backends.wps_com")


class WpsComBackend(ComBackend):
    """通过 win32com 驱动 WPS Office 桌面端

    支持多候选 ProgID 回退:
    WPS 12.x 不再注册 K 前缀 ProgID (如 KWPS.Application),
    此时自动回退到非 K 前缀 ProgID (如 WPS.Application)。
    """

    # ── 候选 ProgID 列表（按优先级排列）──────────────────────────
    CANDIDATE_PROGIDS: dict[str, list[str]] = {
        "writer": [
            "KWPS.Application",
            "WPS.Application",
            "Kingsoft.WPS.Application",
        ],
        "calc": [
            "KET.Application",
            "ET.Application",
            "Kingsoft.ET.Application",
        ],
        "impress": [
            "KWPP.Application",
            "WPP.Application",
            "Kingsoft.WPP.Application",
        ],
    }

    # ── 向后兼容的 APP_MAP（指向首选 ProgID）─────────────────────
    APP_MAP: dict[str, str] = {
        "writer": "KWPS.Application",
        "calc": "KET.Application",
        "impress": "KWPP.Application",
    }

    def connect(self, app_type: str) -> object:
        """连接到 WPS 应用实例（多 ProgID 回退 + 增强诊断）

        依次尝试 CANDIDATE_PROGIDS 中所有候选 ProgID，
        对每个 ProgID 调用 win32com.client.Dispatch 并验证 app.Version，
        成功后调用 harden() 加固并返回应用对象。
        全部失败后运行注册表诊断，抛出带详细信息的 WpsNotFoundError。

        Args:
            app_type: 应用类型 (writer/calc/impress)

        Returns:
            COM 应用对象

        Raises:
            ValueError: 不支持的应用类型
            WpsNotFoundError: 所有候选 ProgID 均连接失败（含注册表诊断信息）
        """
        primary_prog_id = self.APP_MAP.get(app_type)
        if not primary_prog_id:
            raise ValueError(
                f"不支持的应用类型: {app_type}, 可选: {list(self.APP_MAP.keys())}"
            )

        try:
            import win32com.client
        except ImportError as exc:
            raise WpsNotFoundError(
                app_type,
                fix_hint="pywin32 未安装。请运行: pip install pywin32",
            ) from exc

        candidates = self.CANDIDATE_PROGIDS.get(app_type, [primary_prog_id])
        last_error: Exception | None = None

        for prog_id in candidates:
            try:
                app = win32com.client.Dispatch(prog_id)
                # 验证连接真正可用
                try:
                    _ = app.Version
                except AttributeError:
                    # Dispatch 成功但对象无 Version 属性，不可用
                    try:
                        app.Quit()
                    except Exception:
                        pass
                    continue
                # 连接成功，加固并返回
                self.harden(app)
                if prog_id != primary_prog_id:
                    _logger.info(
                        "主 ProgID %s 不可用，使用替代: %s",
                        primary_prog_id,
                        prog_id,
                    )
                return app
            except Exception as exc:
                last_error = exc
                continue

        # ── 所有 ProgID 失败 → 运行注册表诊断 ──────────────────
        registry_lines: list[str] = []
        for prog_id in candidates:
            infos = check_progid_registry(prog_id)
            for info in infos:
                status = "OK" if not info.error else f"FAIL: {info.error}"
                registry_lines.append(
                    f"  {info.prog_id} [{info.registry_view}] -> {status}"
                )
                if info.clsid:
                    registry_lines.append(f"    CLSID: {info.clsid}")
                if info.local_server32:
                    registry_lines.append(
                        f"    LocalServer32: {info.local_server32}"
                    )
        registry_info = "\n".join(registry_lines)

        fix_lines: list[str] = []
        fix_lines.append("1. 运行 'wps doctor --fix' 自动检测并修复 COM 注册")
        fix_lines.append("2. 以管理员身份运行一次 WPS Office（触发自注册）")

        # 位数检查
        python_bits = get_python_bitness()
        wps_path = detect_wps_install_path(app_type)
        if wps_path:
            wps_bits = detect_wps_bitness(wps_path)
            if wps_bits and python_bits != wps_bits:
                fix_lines.append(
                    f"3. 位数不匹配: Python {python_bits}位 vs WPS {wps_bits}位，"
                    f"请安装与 WPS 位数一致的 Python 和 pywin32"
                )
            else:
                fix_lines.append(
                    f"3. 检查 Python/pywin32 位数 ({python_bits}位) 是否与 WPS 位数一致"
                )
        else:
            fix_lines.append(
                "3. 未检测到 WPS 安装路径，请确认 WPS Office 已安装"
            )

        # ksomgr 检测
        ksomgr = find_ksomgr()
        if ksomgr:
            fix_lines.append(
                f'4. 手动运行: "{ksomgr}" -regserver'
            )
        else:
            fix_lines.append(
                "4. 如果以上均无效，请重新安装 WPS Office"
            )

        raise WpsNotFoundError(
            app_type,
            tried_progids=candidates,
            registry_info=registry_info,
            fix_hint="\n".join(fix_lines),
        ) from last_error

    def disconnect(self, app: Any) -> None:
        """断开并关闭 WPS 应用实例"""
        try:
            import pythoncom

            com_error = pythoncom.com_error
        except ImportError:
            com_error = Exception
        try:
            app.Quit()
        except com_error:
            pass

    def is_alive(self, app: Any) -> bool:
        """检查应用实例是否存活"""
        try:
            import pythoncom

            com_error = pythoncom.com_error
        except ImportError:
            com_error = Exception
        try:
            _ = app.Name
            return True
        except com_error:
            return False
        except AttributeError:
            return False

    def get_version(self, app: Any) -> str:
        """获取 WPS 版本号"""
        try:
            return str(app.Version)
        except AttributeError:
            return "unknown"
        except Exception:
            return "unknown"
