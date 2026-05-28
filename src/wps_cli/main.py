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


@app.command()
def doctor() -> None:
    """诊断环境"""
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

    try:
        import pythoncom

        com_error = pythoncom.com_error
    except ImportError:
        com_error = Exception

    for name, prog_id in [
        ("Writer", "KWPS.Application"),
        ("Calc", "KET.Application"),
        ("Impress", "KWPP.Application"),
    ]:
        wps_app = None
        try:
            wps_app = win32com.client.Dispatch(prog_id)
            ver = wps_app.Version
            typer.echo(f"WPS {name}: {ver}")
        except com_error as exc:
            typer.echo(f"WPS {name}: 未检测到 (COM 错误: {exc})")
        except AttributeError:
            typer.echo(f"WPS {name}: 已连接但无法读取版本号")
        except Exception as exc:
            typer.echo(f"WPS {name}: 未检测到 ({type(exc).__name__})")
        finally:
            if wps_app is not None:
                try:
                    wps_app.Quit()
                except Exception:
                    pass

    typer.echo("诊断完成")


if __name__ == "__main__":
    app()
