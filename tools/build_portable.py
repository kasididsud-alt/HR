from __future__ import annotations

import platform
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


APP_NAME = "SalaryCalc"
PORTABLE_NAME = "SalaryCalc_Portable"

ROOT = Path(__file__).resolve().parents[1]
DIST_DIR = ROOT / "dist"
PORTABLE_DIR = DIST_DIR / PORTABLE_NAME
ZIP_PATH = DIST_DIR / f"{PORTABLE_NAME}.zip"
ICON_PATH = ROOT / "assets" / "salary_calc.ico"


def _run(cmd: list[str]) -> None:
    print(">", " ".join(cmd))
    subprocess.run(cmd, cwd=ROOT, check=True)


def _copy_tree(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def _write_readme() -> None:
    if platform.system() == "Darwin":
        text = (
            "Salary & Fee Calculator portable build for macOS\n\n"
            "1. Double-click SalaryCalc.app\n"
            "2. Keep the master and assets folders next to SalaryCalc.app\n"
            "3. If you move the app, move the whole SalaryCalc_Portable folder together\n"
            "4. If macOS blocks the app, right-click SalaryCalc.app and choose Open\n"
        )
    else:
        text = (
            "Salary & Fee Calculator portable build\n\n"
            "1. Run SalaryCalc.exe\n"
            "2. Keep the master and assets folders next to the exe\n"
            "3. If you move the app, move the whole folder together\n"
        )
    (PORTABLE_DIR / "README_PORTABLE.txt").write_text(text, encoding="utf-8")


def _zip_portable() -> None:
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in PORTABLE_DIR.rglob("*"):
            zf.write(path, path.relative_to(DIST_DIR))


def main() -> None:
    DIST_DIR.mkdir(parents=True, exist_ok=True)

    _run([sys.executable, str(ROOT / "tools" / "generate_app_icon.py")])
    system = platform.system()
    pyinstaller_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--name",
        APP_NAME,
    ]

    if system == "Darwin":
        pyinstaller_cmd += ["--windowed", "app.py"]
        built_app = DIST_DIR / f"{APP_NAME}.app"
    else:
        pyinstaller_cmd += [
            "--onefile",
            "--windowed",
            "--icon",
            str(ICON_PATH),
            "app.py",
        ]
        built_app = DIST_DIR / f"{APP_NAME}.exe"

    _run(pyinstaller_cmd)

    if PORTABLE_DIR.exists():
        shutil.rmtree(PORTABLE_DIR)
    PORTABLE_DIR.mkdir(parents=True, exist_ok=True)

    if built_app.is_dir():
        _copy_tree(built_app, PORTABLE_DIR / built_app.name)
    else:
        shutil.copy2(built_app, PORTABLE_DIR / built_app.name)

    _copy_tree(ROOT / "master", PORTABLE_DIR / "master")
    _copy_tree(ROOT / "assets", PORTABLE_DIR / "assets")
    _write_readme()
    _zip_portable()

    print(f"Built: {PORTABLE_DIR}")
    print(f"Zip:   {ZIP_PATH}")


if __name__ == "__main__":
    main()
