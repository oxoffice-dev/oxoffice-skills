"""Install the oxoffice release these skills are written for.

    python3 install.py              # install once, then exit 0 when present
    python3 install.py --wheel F    # install a local wheel instead (development)
    python3 install.py --where      # print the bin folder and exit

Downloads the Python wheel for this platform from the GitHub release named in
`../oxoffice.json`, checks it against the release's SHA256SUMS, and installs it
in its own virtual environment under $OXOFFICE_HOME (default
~/.local/share/oxoffice-skills)/<version>. The environment's bin folder holds
the `oxoffice` command and a `python` that imports `oxoffice`. A `current`
link points at the installed version.

Exit codes: 0 installed, 2 no build available here (offline, unknown platform),
1 anything else. Standard library only, Python 3.10+.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import venv
from pathlib import Path

PLUGIN = Path(__file__).resolve().parent.parent
PIN = json.loads((PLUGIN / "oxoffice.json").read_text(encoding="utf-8"))
VERSION = PIN["version"]
REPO = PIN["repository"]
HOME = Path(os.environ.get("OXOFFICE_HOME", Path.home() / ".local" / "share" / "oxoffice-skills"))
WINDOWS = os.name == "nt"


def bin_dir(root: Path) -> Path:
    return root / ("Scripts" if WINDOWS else "bin")


def wheel_fragment() -> str | None:
    """The part of a wheel file name that names this platform."""
    system, machine = platform.system(), platform.machine().lower()
    if system == "Linux" and machine in ("x86_64", "amd64"):
        return "manylinux_2_28_x86_64"
    if system == "Linux" and machine in ("aarch64", "arm64"):
        return "manylinux_2_28_aarch64"
    if system == "Darwin" and machine == "arm64":
        return "arm64"
    if system == "Darwin" and machine == "x86_64":
        return "x86_64"
    if system == "Windows" and machine in ("amd64", "x86_64"):
        return "win_amd64"
    return None


def fetch(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "oxoffice-skills-installer"})
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read()


def download_wheel(target: Path) -> Path:
    fragment = wheel_fragment()
    if fragment is None:
        raise LookupError(f"no oxoffice build for {platform.system()} {platform.machine()}")
    release = json.loads(fetch(f"https://api.github.com/repos/{REPO}/releases/tags/v{VERSION}"))
    assets = {asset["name"]: asset["browser_download_url"] for asset in release["assets"]}
    wheels = [name for name in assets if name.endswith(".whl") and fragment in name]
    if not wheels:
        raise LookupError(f"release v{VERSION} has no wheel for {fragment}")
    name = sorted(wheels)[0]
    sums = {}
    for line in fetch(assets["SHA256SUMS"]).decode().splitlines():
        digest, _, file_name = line.partition("  ")
        sums[file_name.strip().lstrip("*")] = digest
    data = fetch(assets[name])
    if hashlib.sha256(data).hexdigest() != sums.get(name):
        raise ValueError(f"{name} does not match SHA256SUMS, refusing to install it")
    path = target / name
    path.write_bytes(data)
    return path


def point_current(root: Path) -> None:
    current = HOME / "current"
    if current.is_symlink() or current.is_file():
        current.unlink()
    elif current.is_dir():
        shutil.rmtree(current) if not WINDOWS else os.rmdir(current)
    if WINDOWS:
        subprocess.run(["cmd", "/c", "mklink", "/J", str(current), str(root)], check=True, capture_output=True)
    else:
        current.symlink_to(root, target_is_directory=True)


def installed(root: Path) -> bool:
    exe = bin_dir(root) / ("oxoffice.exe" if WINDOWS else "oxoffice")
    if not exe.exists():
        return False
    result = subprocess.run([str(exe), "--version"], capture_output=True, text=True)
    return result.returncode == 0 and VERSION in result.stdout


def install(wheel: Path | None) -> int:
    root = HOME / VERSION
    if wheel is None and installed(root):
        point_current(root)
        return 0
    HOME.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        try:
            wheel = wheel or download_wheel(Path(tmp))
        except (OSError, LookupError) as error:
            print(f"install: no oxoffice {VERSION} build available here: {error}", file=sys.stderr)
            return 2
        if root.exists():
            shutil.rmtree(root)
        venv.EnvBuilder(with_pip=True).create(root)
        python = bin_dir(root) / ("python.exe" if WINDOWS else "python")
        subprocess.run(
            [str(python), "-m", "pip", "install", "--quiet", "--disable-pip-version-check", str(wheel)],
            check=True,
        )
    point_current(root)
    print(f"install: oxoffice {VERSION} in {bin_dir(root)}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--wheel", type=Path, help="install this local wheel instead of the release")
    parser.add_argument("--where", action="store_true", help="print the bin folder and exit")
    args = parser.parse_args()
    if args.where:
        print(bin_dir(HOME / "current"))
        return 0
    try:
        return install(args.wheel)
    except (subprocess.CalledProcessError, ValueError) as error:
        print(f"install: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
