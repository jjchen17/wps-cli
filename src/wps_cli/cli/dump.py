"""Dump 往返序列化 CLI 命令

设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""

from __future__ import annotations

import typer

from wps_cli.cli.common import handle_error, make_get_service, success
from wps_cli.consts import IMPRESS_INPUT_EXTENSIONS, WRITER_INPUT_EXTENSIONS
from wps_cli.services.dump_service import (
    DumpService,
    batch_replay_to_file,
    dump_impress_to_file,
    dump_writer_to_file,
)
from wps_cli.utils.path_utils import ensure_safe_input_path, ensure_safe_output_path

app = typer.Typer(help="文档序列化与重放")

_get_service = make_get_service(DumpService)


def _safe_writer_input(file: str):
    return ensure_safe_input_path(file, allowed_extensions=WRITER_INPUT_EXTENSIONS)


def _safe_impress_input(file: str):
    return ensure_safe_input_path(file, allowed_extensions=IMPRESS_INPUT_EXTENSIONS)


@app.command()
def writer(
    file: str = typer.Argument(..., help="Word 文档路径 (.docx)"),
    path_filter: str = typer.Argument(
        "",
        help="路径过滤表达式，如 /section[1]/table[2]（可选）",
    ),
    output: str = typer.Option(
        ..., "--output", "-o", help="输出 JSON 文件路径"
    ),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """将 Word 文档序列化为 batch JSON 指令文件

    示例:
        wps dump writer template.docx -o blueprint.json
        wps dump writer template.docx "/section[1]/table[2]" -o table.json
    """
    cmd = "dump.writer"
    try:
        input_path = _safe_writer_input(file)
        out_path = ensure_safe_output_path(output)
        svc = _get_service()
        count = dump_writer_to_file(svc.manager, input_path, out_path, path_filter)
        success(
            {"commands": count, "output": str(out_path)},
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command()
def impress(
    file: str = typer.Argument(..., help="PPT 演示文稿路径 (.pptx)"),
    path_filter: str = typer.Argument(
        "",
        help="路径过滤表达式（可选）",
    ),
    output: str = typer.Option(
        ..., "--output", "-o", help="输出 JSON 文件路径"
    ),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """将 PPT 演示文稿序列化为 batch JSON 指令文件

    示例:
        wps dump impress deck.pptx -o deck.json
    """
    cmd = "dump.impress"
    try:
        input_path = _safe_impress_input(file)
        out_path = ensure_safe_output_path(output)
        svc = _get_service()
        count = dump_impress_to_file(svc.manager, input_path, out_path, path_filter)
        success(
            {"commands": count, "output": str(out_path)},
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)


@app.command("batch")
def batch_replay(
    input_json: str = typer.Option(
        ..., "--input", "-i", help="输入的 JSON 指令文件"
    ),
    output: str = typer.Option(
        ..., "--output", "-o", help="输出文档路径"
    ),
    app_type: str = typer.Option(
        "writer", "--type", "-t", help="文档类型: writer/impress"
    ),
    json_output: bool = typer.Option(False, "--json", "-j", help="JSON 输出"),
):
    """将 batch JSON 指令文件回放为新文档

    示例:
        wps dump batch --input blueprint.json --output generated.docx
        wps dump batch --input deck.json --output new.pptx --type impress
    """
    cmd = "dump.batch"
    try:
        in_path = ensure_safe_input_path(input_json, allowed_extensions=frozenset({".json"}))
        out_path = ensure_safe_output_path(output)
        svc = _get_service()
        executed = batch_replay_to_file(svc.manager, app_type, in_path, out_path)
        success(
            {"executed": executed, "output": str(out_path)},
            command=cmd,
            json_mode=json_output,
        )
    except Exception as e:
        handle_error(e, command=cmd, json_mode=json_output)
