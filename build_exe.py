import os
import subprocess
import sys
import shutil
import zipfile
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = BASE_DIR / "db_table_backup_export.py"
CONFIG_PATH = BASE_DIR / "db_backup_config.json"
DIST_DIR = BASE_DIR / "dist"
DIST_ZIP_PATH = BASE_DIR / "coupang inventory data export.zip"


def prepare_dist_directory():
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    dist_config = DIST_DIR / CONFIG_PATH.name
    if dist_config.exists():
        dist_config.unlink()


def zip_dist_directory():
    if DIST_ZIP_PATH.exists():
        DIST_ZIP_PATH.unlink()

    with zipfile.ZipFile(DIST_ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in DIST_DIR.rglob("*"):
            if path.is_file():
                zf.write(path, arcname=path.relative_to(DIST_DIR))


def main():
    env = os.environ.copy()
    env["PYTHONNOUSERSITE"] = "1"
    user_base = BASE_DIR / ".pyuserbase"
    user_base.mkdir(parents=True, exist_ok=True)
    env["PYTHONUSERBASE"] = str(user_base)

    pyinstaller_args = [
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--name",
        "CoupangInventoryExport",
        "--hidden-import",
        "pymysql",
        "--add-data",
        f"{BASE_DIR / 'db_backup_config.json'};.",
        str(SCRIPT_PATH),
    ]

    launcher = (
        "import site, sys; "
        "site.USER_SITE=''; "
        "site.getusersitepackages=lambda: ''; "
        "import PyInstaller.__main__; "
        f"sys.argv={pyinstaller_args!r}; "
        "PyInstaller.__main__.run()"
    )

    prepare_dist_directory()

    completed = subprocess.run(
        [sys.executable, "-s", "-c", launcher],
        cwd=BASE_DIR,
        env=env,
        check=False,
    )
    if completed.returncode != 0:
        sys.exit(completed.returncode)

    shutil.copy2(CONFIG_PATH, DIST_DIR / CONFIG_PATH.name)
    zip_dist_directory()

    print(f"EXE created under: {DIST_DIR}")
    print(f"Config copied to: {DIST_DIR / CONFIG_PATH.name}")
    print(f"ZIP package created: {DIST_ZIP_PATH}")
    print(f"Build files under: {BASE_DIR / 'build'}")


if __name__ == "__main__":
    main()
