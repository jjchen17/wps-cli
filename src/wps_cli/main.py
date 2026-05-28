"""WPS Office 全量 CLI 工具 — 主入口"""

import sys

import typer

from wps_cli import __version__
from wps_cli.cli import calc, export, impress, pdf, writer


app = typer.Typer(
    name="wps",
    help="WPS Office 全量 CLI 工具 — 通过 COM 自动化驱动 WPS 桌面端",
    no_args_is_help=True,
)

# 注册子命令
app.add_typer(writer.app, name="writer", help="Word 文档操作")
app.add_typer(calc.app, name="calc", help="Excel 电子表格操作")
app.add_typer(impress.app, name="impress", help="PPT 演示文稿操作")
app.add_typer(pdf.app, name="pdf", help="PDF 文档操作")
app.add_typer(export.app, name="export", help="格式转换与导出")


@app.command()
def version():
    """输出版本信息"""
    typer.echo(f"wps-cli {__version__}")
    typer.echo(f"Python {sys.version}")


@app.command()
def doctor():
    """诊断环境"""
    typer.echo(f"Python: {sys.version}")
    typer.echo(f"平台: {sys.platform}")

    if sys.platform != "win32":
        typer.echo("错误: 仅支持 Windows")
        raise typer.Exit(1)

    try:
        import win32com.client  # noqa: F401
        typer.echo("pywin32: 已安装")
    except ImportError:
        typer.echo("错误: pywin32 未安装")
        raise typer.Exit(1)

    for name, prog_id in [
        ("Writer", "KWPS.Application"),
        ("Calc", "KET.Application"),
        ("Impress", "KWPP.Application"),
    ]:
        app = None
        try:
            app = win32com.client.Dispatch(prog_id)
            ver = app.Version
            typer.echo(f"WPS {name}: {ver}")
        except Exception:
            typer.echo(f"WPS {name}: 未检测到")
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass

    typer.echo("诊断完成")


if __name__ == "__main__":
    app()
