"""Run every way of installing oxoffice that the skills document (references/install.md, the same in both skills),
as written, and check that each gives the version the skills pin: `ox --version` and, where the way defines it,
`oxpy` importing oxoffice.docx and oxoffice.pptx at that version.

    python tests/check_install.py plugin/skills

Each bash block that defines `ox` (all ways but "Already installed") runs under `bash -euo pipefail` with a fresh
HOME and no XDG_CACHE_HOME, with the checks appended; a way renamed or gone fails the run. On Windows the block gets
the substitutions install.md gives for Git Bash (`python` for `python3`, `Scripts` for `bin`, `/python` for
`/bin/python`), the shell installer is replaced by the PowerShell one, and a way whose tool is missing here (conda
on some runners) is reported as skipped, not passed.
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


def pinned():
    found = {v for f in SKILLS.glob("*/SKILL.md") for v in re.findall(r"oxoffice==(\d+\.\d+\.\d+)", f.read_text())}
    assert len(found) == 1, f"the skills pin {sorted(found) or 'no version'}, expected one"
    return found.pop()


def ways(text):
    """(title, language, code) of each block that installs oxoffice and defines `ox`, or the PowerShell installer:
    every way but "Already installed". A block that names another version fails its check, it is not skipped."""
    title, out = "", []
    heading = r"^[ \t]*(?:(\d\. \*\*[^*]+\*\*|\*\*[^*]+\*\*).*?$)"
    fence = r"^[ \t]*```(\w+)\n(.*?)^[ \t]*```"
    for line_title, lang, code in re.findall(f"{heading}|{fence}", text, re.S | re.M):
        if line_title:
            title = line_title.strip("*0123456789. ")
        elif title != "Already installed" and (lang == "bash" and "ox()" in code or lang == "powershell"):
            out.append((title, lang, code))
    return out


def for_windows(title, code):
    code = code.replace("python3 ", "python ")
    if "/bin/python" in code and "conda" in title.lower():
        return code.replace("/bin/python", "/python")
    return code.replace("/bin/", "/Scripts/")


# The ways install.md documents, by their titles: a way renamed or no longer found fails the run.
EXPECTED = {"uv", "pip in a virtual environment", "pixi", "conda", "The CLI alone", "Windows"}


def check(version, has_oxpy):
    """Lines that fail unless `ox` (and `oxpy`) give `version`. Python on Windows ends a piped line with CR."""
    lines = [f'ox --version | tr -d "\\r" | tee /dev/stderr | grep -qx "oxoffice {version}"']
    if has_oxpy:
        lines.append("oxpy -c 'import oxoffice, oxoffice.docx, oxoffice.pptx; print(oxoffice.__version__)'"
                     f' | tr -d "\\r" | tee /dev/stderr | grep -qx "{version}"')
    return "\n".join(lines) + "\n"


def git_bash_ox(text):
    """The `ox` definition install.md gives in Git Bash for the PowerShell way."""
    found = re.findall(r"then,\s+in\s+Git\s+Bash,\s+`(ox\(\) \{.*?\})`", text, re.S)
    assert len(found) == 1, "install.md: no single Git Bash `ox` line after the PowerShell way"
    return found[0]


def run(title, script, env):
    print(f"--- {title}", flush=True)
    proc = subprocess.run([BASH, "-euo", "pipefail", "-c", script], env=env, text=True, capture_output=True)
    print((proc.stdout + proc.stderr).strip()[-3000:], flush=True)
    return proc.returncode == 0


def main():
    version = pinned()
    text = (SKILLS / "docx" / "references" / "install.md").read_text()
    other = (SKILLS / "pptx" / "references" / "install.md").read_text()
    assert text == other, "the two skills' install.md differ"
    found = ways(text)
    assert {title for title, _, _ in found} == EXPECTED, (
        f"install.md ways: {sorted(t for t, _, _ in found)}, expected {sorted(EXPECTED)}")
    results = {}
    with tempfile.TemporaryDirectory() as home:
        # A home of its own and no XDG_CACHE_HOME: the ways take the `$HOME/.cache` fallback a user gets.
        env = {k: v for k, v in os.environ.items() if k != "XDG_CACHE_HOME"}
        env.update(HOME=home)
        for title, lang, code in found:
            needs = "conda" if "conda create" in code else "pixi" if "pixi exec" in code else (
                "uvx" if "uvx" in code else None)
            if lang == "powershell":
                if not WINDOWS:
                    continue
                env_ps = dict(env, LOCALAPPDATA=home)
                print(f"--- {title} (PowerShell)", flush=True)
                proc = subprocess.run(["pwsh", "-NoProfile", "-Command", code], env=env_ps, text=True,
                                      capture_output=True)
                print((proc.stdout + proc.stderr).strip()[-3000:], flush=True)
                ok = proc.returncode == 0 and run(f"{title} (Git Bash)", git_bash_ox(text) + "\n"
                                                  + check(version, False), env_ps)
                results[f"{title} (PowerShell)"] = "ok" if ok else "FAIL"
                continue
            if WINDOWS and "install.sh" in code:
                continue        # the PowerShell installer stands for it on Windows
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
