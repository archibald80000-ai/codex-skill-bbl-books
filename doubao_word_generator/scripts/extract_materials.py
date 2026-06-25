from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PARENT_SCRIPTS = ROOT / "scripts"
if str(PARENT_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(PARENT_SCRIPTS))

from inspect_inputs import InputItem, inspect_inputs  # type: ignore


def extract_zip(zip_path: str | Path, target_dir: str | Path) -> Path:
    target = Path(target_dir)
    target.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(target)
    return target


def copy_files(files: list[Path], target_dir: str | Path) -> list[Path]:
    target = Path(target_dir)
    target.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    for src in files:
        dest = target / src.name
        shutil.copy2(src, dest)
        copied.append(dest)
    return copied


def load_materials(inputs: list[str]) -> list[InputItem]:
    return inspect_inputs(inputs)

