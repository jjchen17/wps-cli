"""WPS Office 全量 CLI 工具 — 主入口"""

from __future__ import annotations

import sys

import typer

from wps_cli import __version__
from wps_cli.cli import calc, export, impress, pdf, writer

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


@app.command()
def version() -> None:
    """输出版本信息"""
    typer.echo(f"wps-cli {__version__}")
    typer.echo(f"Python {sys.version.split()[0]}")


def _detect_wps_versions() -> dict[str, str]:
    """探测三个 WPS 组件的版本"""
    result: dict[str, str] = {}
    try:
        import pythoncom
        import win32com.client

        com_error = pythoncom.com_error
    except ImportError:
        return {"Writer": "pywin32 未安装", "Calc": "pywin32 未安装", "Impress": "pywin32 未安装"}

    for name, prog_id in [
        ("Writer", "KWPS.Application"),
        ("Calc", "KET.Application"),
        ("Impress", "KWPP.Application"),
    ]:
        wps_app = None
        try:
            wps_app = win32com.client.Dispatch(prog_id)
            result[name] = str(wps_app.Version)
        except com_error:
            result[name] = "未检测到（COM 错误）"
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


def _print_doctor_text() -> None:
    """人类友好的 doctor 输出"""
    typer.echo(f"Python: {sys.version.split()[0]}")
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

    for name, ver in _detect_wps_versions().items():
        typer.echo(f"WPS {name}: {ver}")
    typer.echo("诊断完成")


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
) -> None:
    """诊断环境

    默认人类友好输出。``--report`` 输出脱敏 markdown，便于反馈 bug。
    """
    if report:
        _print_doctor_report()
    else:
        _print_doctor_text()


if __name__ == "__main__":
    app()
