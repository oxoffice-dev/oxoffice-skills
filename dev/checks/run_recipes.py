"""Run every ```python and ```bash block of each skill's recipes.md in order, in a fresh folder per file, the way
rdocx-skills' tests/test_docs_snippets.py did: the folder starts with report.docx, deck.pptx and Aileron-Regular.ttf;
python blocks run with "$O/python" -c, bash blocks with eval, both under bash -euo pipefail.

    python3 run_recipes.py [docx|pptx ...]
"""
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

CHECK = Path(__file__).resolve().parent
SKILLS = Path(__file__).resolve().parents[2] / "plugin" / "skills"
HOME = CHECK.parent / "skills-home"
BIN = HOME / "current" / "bin"
FIXTURES = CHECK / "fixtures"

names = sys.argv[1:] or ["docx", "pptx"]
failed = 0
for name in names:
    recipes = SKILLS / name / "references" / "recipes.md"
    blocks = re.findall(r"```(python|bash)\n(.*?)```", recipes.read_text(), re.S)
    work = CHECK / f"run-{name}"
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir()
    for f in ("report.docx", "deck.pptx", "Aileron-Regular.ttf"):
        shutil.copy(FIXTURES / f, work / f)
    skill = SKILLS / name
    env = {k: v for k, v in os.environ.items() if k not in ("R", "PYTHONPATH", "VIRTUAL_ENV")}
    env.update(O=BIN.as_posix(), SKILL=skill.as_posix(), OXOFFICE_HOME=str(HOME),
               PYTHONPATH=str(skill / "scripts"), PYTHONDONTWRITEBYTECODE="1")
    for k, (lang, code) in enumerate(blocks):
        env["RECIPE"] = code
        script = '"$O/python" -c "$RECIPE"' if lang == "python" else 'eval "$RECIPE"'
        t0 = time.time()
        res = subprocess.run(["bash", "-euo", "pipefail", "-c", script], cwd=work, env=env, capture_output=True,
                             text=True, timeout=300)
        status = "ok" if res.returncode == 0 else f"FAIL {res.returncode}"
        print(f"{name}-{k:02d}-{lang}: {status} ({time.time() - t0:.1f} s)")
        (work / f"_block{k:02d}.log").write_text(code + "\n----- stdout\n" + res.stdout + "\n----- stderr\n" + res.stderr)
        if res.returncode:
            failed += 1
            print("  " + (res.stdout[-1500:] + res.stderr[-2500:]).replace("\n", "\n  "))
    print(f"{name}: {len(blocks)} blocks")
print("failed:", failed)
sys.exit(1 if failed else 0)
