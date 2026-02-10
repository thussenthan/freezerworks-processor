#!/usr/bin/env python3
import argparse
import os
import subprocess
import sys
from pathlib import Path


def run(cmd):
    print("Running:", " ".join(cmd))
    subprocess.run(cmd, check=True)


def build_windows():
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise RuntimeError("LOCALAPPDATA is not set. Run this on Windows.")

    runtime_tmp = Path(local_app_data) / "Freezerworks Processor" / "runtime"
    runtime_tmp.mkdir(parents=True, exist_ok=True)

    cmd = [
        "pyinstaller",
        "--clean",
        "--onefile",
        "--windowed",
        "--runtime-tmpdir",
        str(runtime_tmp),
        "--add-data",
        "freezerworks.pennstatehealth.net.crt;.",
        "freezerworks_processor.py",
    ]
    run(cmd)


def build_macos():
    cmd = [
        "pyinstaller",
        "--clean",
        "--onefile",
        "--windowed",
        "--add-data",
        "freezerworks.pennstatehealth.net.cer:.",
        "freezerworks_processor.py",
    ]
    run(cmd)


def main():
    parser = argparse.ArgumentParser(description="Build Freezerworks Processor executable")
    parser.add_argument(
        "--target",
        choices=["windows", "macos", "auto"],
        default="auto",
        help="Build target (default: auto from current OS)",
    )
    args = parser.parse_args()

    target = args.target
    if target == "auto":
        if sys.platform.startswith("win"):
            target = "windows"
        elif sys.platform == "darwin":
            target = "macos"
        else:
            raise RuntimeError("Unsupported OS for auto target. Use --target explicitly.")

    if target == "windows":
        build_windows()
    elif target == "macos":
        build_macos()


if __name__ == "__main__":
    main()
