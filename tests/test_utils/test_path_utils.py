"""路径工具测试 — 重点覆盖路径遍历防护"""

import pytest

from wps_cli.exceptions import FileNotFoundErrorCli, ValidationError
from wps_cli.utils.path_utils import (
    ensure_safe_glob,
    ensure_safe_input_path,
    ensure_safe_output_path,
    redact_path,
    sanitize_filename,
    validate_path,
)


class TestValidatePath:
    def test_relative_path_ok(self):
        p = validate_path(".")
        assert p.is_absolute()

    def test_unc_path_rejected(self):
        with pytest.raises(ValidationError, match="UNC"):
            validate_path("\\\\server\\share\\file.txt")

    def test_unc_forward_slash_rejected(self):
        with pytest.raises(ValidationError, match="UNC"):
            validate_path("//server/share/file.txt")

    def test_allowed_dirs_whitelist(self, tmp_path):
        sub = tmp_path / "ok"
        sub.mkdir()
        f = sub / "x.txt"
        f.write_text("x")
        # 在白名单内
        result = validate_path(str(f), allowed_dirs=[str(tmp_path)])
        assert result == f.resolve()
        # 不在白名单内
        with pytest.raises(ValidationError, match="不在允许"):
            validate_path(str(f), allowed_dirs=[str(tmp_path / "other")])


class TestEnsureSafeInputPath:
    def test_existing_file_ok(self, tmp_path):
        f = tmp_path / "x.txt"
        f.write_text("ok")
        p = ensure_safe_input_path(str(f))
        assert p.exists()

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundErrorCli):
            ensure_safe_input_path(str(tmp_path / "missing.txt"))

    def test_directory_rejected(self, tmp_path):
        with pytest.raises(ValidationError, match="不是文件"):
            ensure_safe_input_path(str(tmp_path))


class TestEnsureSafeOutputPath:
    def test_creates_parent(self, tmp_path):
        target = tmp_path / "deep" / "nest" / "x.txt"
        p = ensure_safe_output_path(str(target))
        assert p.parent.exists()


class TestEnsureSafeGlob:
    """C-3：glob 路径边界限制"""

    def test_absolute_path_rejected(self, tmp_path):
        with pytest.raises(ValidationError, match="绝对路径"):
            ensure_safe_glob("C:\\Windows\\*.exe")

    def test_unc_rejected(self):
        with pytest.raises(ValidationError, match="UNC"):
            ensure_safe_glob("\\\\server\\share\\*")

    def test_empty_rejected(self):
        with pytest.raises(ValidationError):
            ensure_safe_glob("")

    def test_relative_pattern_works(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        (tmp_path / "a.docx").write_text("x")
        (tmp_path / "b.docx").write_text("x")
        (tmp_path / "c.txt").write_text("x")
        results = ensure_safe_glob("*.docx")
        names = sorted(p.name for p in results)
        assert names == ["a.docx", "b.docx"]

    def test_recursive_pattern_constrained_to_cwd(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        sub = tmp_path / "sub"
        sub.mkdir()
        (sub / "x.docx").write_text("x")
        results = ensure_safe_glob("**/*.docx")
        assert len(results) == 1
        assert results[0].name == "x.docx"


class TestSanitizeFilename:
    def test_clean(self):
        assert sanitize_filename("hello") == "hello"

    def test_replaces_illegal(self):
        assert sanitize_filename("file<name>.txt") == "file_name_.txt"

    def test_rejects_empty(self):
        with pytest.raises(ValidationError):
            sanitize_filename("")


class TestRedactPath:
    def test_windows_path_redacted(self):
        text = "open failed: C:\\Users\\alice\\Documents\\secret.docx"
        result = redact_path(text)
        assert "alice" not in result
        assert "<path>" in result

    def test_empty_string(self):
        assert redact_path("") == ""
