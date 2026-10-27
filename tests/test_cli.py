from __future__ import annotations

import json
import subprocess
import sys

from officecheck.cli import main


def write_spec(tmp_path, checks):
    p = tmp_path / "spec.json"
    p.write_text(json.dumps({"checks": checks}), encoding="utf-8")
    return str(p)


def test_cli_pass_prints_table(files, tmp_path, capsys):
    spec = write_spec(tmp_path, [{"id": "pages", "type": "pdf.page_count", "equals": 2}])
    assert main([str(files["pdf_good"]), spec]) == 0
    out = capsys.readouterr().out
    assert "pages" in out and "PASS" in out and "score=1.000" in out


def test_cli_fail_exit_1(files, tmp_path, capsys):
    spec = write_spec(tmp_path, [{"type": "pdf.page_count", "equals": 2}, {"type": "pdf.is_text_searchable"}])
    assert main([str(files["pdf_bad"]), spec]) == 1
    assert "FAIL" in capsys.readouterr().out


def test_cli_json(files, tmp_path, capsys):
    spec = write_spec(tmp_path, [{"type": "pdf.page_count", "equals": 2}])
    assert main([str(files["pdf_good"]), spec, "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["pass"] is True


def test_cli_invalid_spec_exit_2(files, tmp_path, capsys):
    spec = write_spec(tmp_path, [{"type": "pdf.page_count"}])
    assert main([str(files["pdf_good"]), spec]) == 2
    assert "invalid spec" in capsys.readouterr().err


def test_cli_entry_point_module(files, tmp_path):
    spec = write_spec(tmp_path, [{"type": "any.file_opens", "format": "xlsx"}])
    proc = subprocess.run([sys.executable, "-m", "officecheck", str(files["xlsx_good"]), spec],
                          capture_output=True, text=True)
    assert proc.returncode == 0, proc.stdout + proc.stderr


def test_print_schema(capsys):
    assert main(["--print-schema"]) == 0
    assert "docx.has_toc" in capsys.readouterr().out
