"""Run every way of installing oxoffice that the skills document (references/install.md, the same in both skills),
as written, and check that each gives the version the skills ask for, their floor (`oxoffice>=X` in their Setup) or
a later one: `ox --version` and, where the way defines it, `oxpy` importing oxoffice.docx, oxoffice.pptx and oxoffice.xlsx.

    python tests/check_install.py plugin/skills

Each bash block that defines `ox` (all ways but "Already installed") runs under `bash -euo pipefail` with a fresh
HOME and no XDG_CACHE_HOME, with the checks appended; a way renamed or gone fails the run. On Windows the block gets
the substitutions install.md gives for Git Bash (`python` for `python3`, `Scripts` for `bin`, `/python` for
`/bin/python`), and a way whose tool is missing here (conda on some runners) is reported as skipped, not passed.
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SKILLS = Path(sys.argv[1]).resolve()
WINDOWS = sys.platform == "win32"
# The full path: on Windows a bare "bash" would find System32's WSL bash before Git Bash on PATH.
BASH = shutil.which("bash")


def floor():
    found = {v for f in SKILLS.glob("*/SKILL.md") for v in re.findall(r"oxoffice>=(\d+\.\d+\.\d+)", f.read_text())}
    assert len(found) == 1, f"the skills ask for {sorted(found) or 'no version'}, expected one floor"
    return found.pop()


def ways(text):
    """(title, code) of each bash block that installs oxoffice and defines `ox`: every way but "Already installed".
    A block that names another version fails its check, it is not skipped."""
    title, out = "", []
    heading = r"^[ \t]*(?:(\d\. \*\*[^*]+\*\*|\*\*[^*]+\*\*).*?$)"
    fence = r"^[ \t]*```(\w+)\n(.*?)^[ \t]*```"
    for line_title, lang, code in re.findall(f"{heading}|{fence}", text, re.S | re.M):
        if line_title:
            title = line_title.strip("*0123456789. ")
        elif title != "Already installed" and lang == "bash" and "ox()" in code:
            out.append((title, code))
    return out


def for_windows(title, code):
    code = code.replace("python3 ", "python ")
    if "/bin/python" in code and "conda" in title.lower():
        return code.replace("/bin/python", "/python")
    return code.replace("/bin/", "/Scripts/")


# The ways install.md documents, by their titles: a way renamed or no longer found fails the run.
EXPECTED = {"uv", "pip in a virtual environment", "pixi", "conda"}


def check(version, has_oxpy):
    """Lines that fail unless `ox` (and `oxpy`) give `version` or a later one. Python on Windows ends a piped line
    with CR, and /dev/stderr does not exist there when the output is captured: each answer goes through a variable."""
    strip = '| tr -d "\\r")"; echo "$v"'
    at_least = f'[ "$(printf "%s\\n" "{version}" "$v" | sort -V | head -n 1)" = "{version}" ]'
    lines = [f'v="$(ox --version {strip}', '[ "${v#oxoffice }" != "$v" ]', 'v="${v#oxoffice }"', at_least]
    if has_oxpy:
        code = "import oxoffice, oxoffice.docx, oxoffice.pptx, oxoffice.xlsx; print(oxoffice.__version__)"
        lines += [f"v=\"$(oxpy -c '{code}' {strip}", at_least]
    return "\n".join(lines) + "\n"


def run(title, script, env):
    print(f"--- {title}", flush=True)
    proc = subprocess.run([BASH, "-euo", "pipefail", "-c", script], env=env, text=True, capture_output=True)
    print((proc.stdout + proc.stderr).strip()[-3000:], flush=True)
    return proc.returncode == 0


def main():
    version = floor()
    text = (SKILLS / "docx" / "references" / "install.md").read_text()
    for name in ("pptx", "xlsx"):
        assert text == (SKILLS / name / "references" / "install.md").read_text(), f"install.md of docx and {name} differ"
    found = ways(text)
    assert {title for title, _ in found} == EXPECTED, (
        f"install.md ways: {sorted(t for t, _ in found)}, expected {sorted(EXPECTED)}")
    results = {}
    with tempfile.TemporaryDirectory() as home:
        # A home of its own and no XDG_CACHE_HOME: the ways take the `$HOME/.cache` fallback a user gets.
        env = {k: v for k, v in os.environ.items() if k != "XDG_CACHE_HOME"}
        env.update(HOME=home)
        for title, code in found:
            needs = "conda" if "conda create" in code else "pixi" if "pixi exec" in code else (
                "uvx" if "uvx" in code else None)
            if needs and not shutil.which(needs):
                results[title] = f"skipped: no {needs} on this runner"
                print(f"::warning::{title}: not checked, no {needs} on this runner", flush=True)
                continue
            script = (for_windows(title, code) if WINDOWS else code) + "\n" + check(version, "oxpy()" in code)
            results[title] = "ok" if run(title, script, env) else "FAIL"
    print(f"\noxoffice {version}, {sys.platform}:")
    for title, result in results.items():
        print(f"  {title}: {result}")
    if not results or any(r == "FAIL" for r in results.values()):
        sys.exit(1)


if __name__ == "__main__":
    main()
