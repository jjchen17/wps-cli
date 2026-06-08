"""WPS Office 全量 CLI 工具 — 主入口"""

from __future__ import annotations

import sys

import typer

from wps_cli import __version__
from wps_cli.cli import calc, export, impress, install_cmd, mcp_cmd, pdf, resident, writer

app = typer.Typer(
    name="wps",
    help="WPS Office 全量 CLI 工具 — 通过 COM 自动化驱动 WPS 桌面端",
    no_args_is_help=True,
)

app.add_typer(writer.app, name="writer", help="Word 文档操作")
app.add_typer(calc.app, name="calc", help="Excel 电子表格操作")
app.add_typer(impress.app, name="impress", help="PPT 演示文稿操作")
app.add_typer(pdf.app, name="pdf", help="PDF 文档操作")
app.add_typer(export.app, name="export", help="格式转换与导出")
app.add_typer(resident.app, name="resident", help="驻留模式管理")
app.add_typer(mcp_cmd.app, name="mcp", help="MCP 服务器管理")
app.add_typer(install_cmd.app, name="install", help="AI 工具集成安装")


@app.command()
def version() -> None:
    """输出版本信息"""
    typer.echo(f"wps-cli {__version__}")
    typer.echo(f"Python {sys.version.split()[0]}")


def _detect_wps_versions() -> dict[str, str]:
    """探测三个 WPS 组件的版本，使用多候选 ProgID 自动回退"""
    from wps_cli.consts import COM_PROGID_CANDIDATES

    result: dict[str, str] = {}
    try:
        import pythoncom
        import win32com.client

        com_error = pythoncom.com_error
    except ImportError:
        return {"Writer": "pywin32 未安装", "Calc": "pywin32 未安装", "Impress": "pywin32 未安装"}

    for name, prog_ids in [
        ("Writer", COM_PROGID_CANDIDATES["writer"]),
        ("Calc", COM_PROGID_CANDIDATES["calc"]),
        ("Impress", COM_PROGID_CANDIDATES["impress"]),
    ]:
        wps_app = None
        used_prog_id: str | None = None
        for prog_id in prog_ids:
            try:
                wps_app = win32com.client.Dispatch(prog_id)
                used_prog_id = prog_id
                break
            except com_error:
                continue
            except Exception:
                continue

        if wps_app is None:
            result[name] = "未检测到（COM 错误）"
            continue

        try:
            ver = str(wps_app.Version)
            if used_prog_id and used_prog_id != prog_ids[0]:
                result[name] = f"{ver} (使用 {used_prog_id})"
            else:
                result[name] = ver
        except AttributeError:
            result[name] = "已连接但无法读取版本号"
        except Exception as exc:
            result[name] = f"未检测到（{type(exc).__name__}）"
        finally:
            if wps_app is not None:
                try:
                    wps_app.Quit()
                except Exception:
                    pass
    return result


def _check_bitness_mismatch() -> dict:
    """检测 Python 与 WPS 的位数是否匹配"""
    from wps_cli.services.com_diagnostics import detect_wps_bitness, detect_wps_install_path

    python_bits = 64 if sys.maxsize > 2**32 else 32
    result: dict = {
        "python_bits": python_bits,
        "wps_bits": None,
        "mismatch": False,
        "wps_path": None,
    }

    wps_path = detect_wps_install_path("writer")
    result["wps_path"] = wps_path
    if wps_path:
        wps_bits = detect_wps_bitness(wps_path)
        result["wps_bits"] = wps_bits
        if wps_bits and wps_bits != python_bits:
            result["mismatch"] = True

    return result


def _print_doctor_text() -> None:
    """人类友好的 doctor 输出"""
    py_bits = 64 if sys.maxsize > 2**32 else 32

    typer.echo(f"Python: {sys.version.split()[0]}")
    typer.echo(f"Python 位数: {py_bits}-bit")
    typer.echo(f"平台: {sys.platform}")

    if sys.platform != "win32":
        typer.echo("错误: 仅支持 Windows", err=True)
        raise typer.Exit(1)

    try:
        import win32com.client  # noqa: F401
    except ImportError as exc:
        typer.echo(f"错误: pywin32 未安装 — {exc}", err=True)
        raise typer.Exit(1) from exc
    typer.echo("pywin32: 已安装")

    # WPS 安装路径和位数
    from wps_cli.services.com_diagnostics import detect_wps_bitness, detect_wps_install_path

    wps_path = detect_wps_install_path("writer")
    if wps_path:
        typer.echo(f"WPS 安装路径: {wps_path}")
        wps_bits = detect_wps_bitness(wps_path)
        if wps_bits:
            typer.echo(f"WPS 位数: {wps_bits}-bit")
            if wps_bits != py_bits:
                typer.echo(
                    f"⚠ 位数不匹配: Python {py_bits}位 vs WPS {wps_bits}位", err=True,
                )
                typer.echo("  请安装匹配位数的 Python 或 pywin32", err=True)
        else:
            typer.echo("WPS 位数: 无法检测")
    else:
        typer.echo("WPS 安装路径: 未检测到")

    # WPS 版本检测
    for name, ver in _detect_wps_versions().items():
        typer.echo(f"WPS {name}: {ver}")

    # 注册表诊断（检查前2个候选 ProgID）
    from wps_cli.consts import COM_PROGID_CANDIDATES
    from wps_cli.services.com_diagnostics import check_progid_registry

    typer.echo("")
    typer.echo("注册表诊断 (前2个候选 ProgID):")
    for _app_name, app_type in [("Writer", "writer"), ("Calc", "calc"), ("Impress", "impress")]:
        candidates = COM_PROGID_CANDIDATES.get(app_type, [])
        for prog_id in candidates[:2]:
            results = check_progid_registry(prog_id)
            for r in results:
                if r.error:
                    typer.echo(f"  {prog_id}: {r.error}")
                elif r.local_server_exists:
                    typer.echo(f"  {prog_id}: OK (CLSID {r.clsid})")
                else:
                    typer.echo(f"  {prog_id}: LocalServer32 路径不存在 ({r.local_server32})")

    # ksomgr 检测
    from wps_cli.services.com_diagnostics import find_ksomgr

    ksomgr = find_ksomgr()
    if ksomgr:
        typer.echo(f"\n注册修复工具: {ksomgr}")
        typer.echo("  运行 'wps doctor --fix' 自动修复 COM 注册问题（需管理员权限）")
    else:
        typer.echo("\n未找到 ksomgr.exe，无法自动修复")

    typer.echo("\n诊断完成")


def _run_doctor_fix() -> None:
    """自动检测并修复 COM 注册问题"""
    from wps_cli.services.com_diagnostics import attempt_com_fix

    typer.echo("正在诊断并修复 COM 注册问题...\n")
    success, logs = attempt_com_fix()
    for line in logs:
        typer.echo(line)
    if success:
        typer.echo("\n修复成功！请重新运行 'wps doctor' 验证。")


def _print_doctor_report() -> None:
    """脱敏的 markdown 报告，便于粘贴到 GitHub Issue

    报告**不包含**：文件路径、文件名、用户名、计算机名、IP、单元格内容。
    """
    import platform

    py = sys.version.split()[0]
    impl = platform.python_implementation()
    bits = "64-bit" if sys.maxsize > 2**32 else "32-bit"
    plat = sys.platform
    win_release = platform.release() if plat == "win32" else ""
    win_build = ""
    if plat == "win32":
        try:
            win_build = platform.version().split(".")[-1]
        except Exception:
            win_build = ""

    pywin32_ver: str = "未安装"
    try:
        import importlib.metadata as md

        pywin32_ver = md.version("pywin32")
    except Exception:
        pass

    versions = _detect_wps_versions() if plat == "win32" else {}

    lines = [
        "### Environment Report (`wps doctor --report`)",
        "",
        f"- wps-cli: {__version__}",
        f"- Python: {py} ({impl}, {bits})",
        f"- Platform: {plat}"
        + (f" (Windows {win_release}, Build {win_build})" if win_release else ""),
        f"- pywin32: {pywin32_ver}",
    ]
    if versions:
        for name, ver in versions.items():
            lines.append(f"- WPS {name}: {ver}")
    else:
        lines.append("- WPS: 非 Windows 平台或未检测")

    # COM 注册表状态摘要
    if plat == "win32":
        from wps_cli.consts import COM_PROGID_CANDIDATES
        from wps_cli.services.com_diagnostics import check_progid_registry

        lines.append("")
        lines.append("### COM Registry Status")
        lines.append("")
        for app_name, app_type in [("Writer", "writer"), ("Calc", "calc"), ("Impress", "impress")]:
            candidates = COM_PROGID_CANDIDATES.get(app_type, [])
            has_ok = False
            for prog_id in candidates:
                results = check_progid_registry(prog_id)
                for r in results:
                    if r.local_server_exists and not r.error:
                        lines.append(f"- {app_name}: {prog_id} OK (CLSID `{r.clsid}`)")
                        has_ok = True
                        break
                if has_ok:
                    break
            if not has_ok:
                lines.append(f"- {app_name}: 无可用 ProgID")

    lines.extend(
        [
            "",
            "<!-- 此报告未包含文件名、文件路径、单元格内容，可放心粘贴 -->",
            "",
            "### What I was doing",
            "",
            "<!-- 请在此粘贴你执行的命令和期望/实际行为 -->",
        ]
    )
    typer.echo("\n".join(lines))


@app.command()
def doctor(
    report: bool = typer.Option(
        False,
        "--report",
        help="输出可粘贴到 GitHub Issue 的脱敏 markdown 报告",
    ),
    fix: bool = typer.Option(
        False,
        "--fix",
        help="自动检测并修复 COM 注册问题（需管理员权限）",
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose",
        "-v",
        help="输出详细诊断信息",
    ),
) -> None:
    """诊断环境

    默认人类友好输出。``--report`` 输出脱敏 markdown，便于反馈 bug。
    ``--fix`` 自动修复 COM 注册问题。``--verbose`` 输出详细诊断。
    """
    if fix:
        _run_doctor_fix()
    elif report:
        _print_doctor_report()
    else:
        _print_doctor_text()
        if verbose:
            bitness = _check_bitness_mismatch()
            typer.echo(f"\n详细诊断: Python {bitness['python_bits']}位"
                       f" | WPS {bitness.get('wps_bits', '?')}位"
                       f" | 不匹配: {bitness['mismatch']}")


if __name__ == "__main__":
    app()
