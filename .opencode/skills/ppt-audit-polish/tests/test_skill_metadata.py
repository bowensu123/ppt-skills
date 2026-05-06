from __future__ import annotations

import os
import platform
import subprocess
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[4]
SKILL_ROOT = Path(__file__).resolve().parents[1]
INSTALL_SCRIPT = REPO_ROOT / "scripts" / "install_ppt_audit_polish.sh"


def test_skill_frontmatter_mentions_audit_and_polish() -> None:
    text = (SKILL_ROOT / "SKILL.md").read_text(encoding="utf-8")
    assert "name: ppt-audit-polish" in text
    assert "polish" in text.lower()
    assert "audit" in text.lower()


@pytest.mark.skipif(platform.system() == "Windows", reason="bash install script not directly executable on Windows")
def test_install_script_copies_skill_tree(tmp_path: Path) -> None:
    env = os.environ.copy()
    env["OPENCODE_SKILL_TARGET_ROOT"] = str(tmp_path)
    result = subprocess.run(
        [str(INSTALL_SCRIPT)],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        env=env,
    )

    assert result.returncode == 0, result.stderr
    installed_root = tmp_path / "ppt-audit-polish"
    assert (installed_root / "SKILL.md").exists()
    assert not (installed_root / ".venv").exists()
    assert not (installed_root / ".pytest_cache").exists()
    assert not any(installed_root.rglob("__pycache__"))
    assert not any(installed_root.rglob("*.egg-info"))


def test_install_script_resolves_source_relative_to_its_own_location() -> None:
    text = INSTALL_SCRIPT.read_text(encoding="utf-8")
    assert "BASH_SOURCE" in text
    assert "SOURCE_DIR='/Users/subowen/Documents/New project/.opencode/skills/ppt-audit-polish'" not in text
