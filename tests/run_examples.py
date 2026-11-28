"""Run every block of each skill's references/examples.md in one fresh folder per skill that starts with the files of
make_fixtures.py. Every block is ```python or ```bash, indented or not; any other fence (another language, none, or
~~~) fails the run, as does a file with no block. Python blocks run as `oxpy -c BLOCK`, bash blocks with `eval`, both
under `bash -euo pipefail`, with what the skills' Setup defines: the shell functions `ox` (the oxoffice command) and
`oxpy` (a Python that imports oxoffice), SKILL (the skill's folder) and PYTHONPATH (its scripts). Each block's code
and output go to WORK/<skill>/_blockNN.log.

    python tests/run_examples.py --skills SKILLS_DIR --fixtures DIR --work DIR [--jobs N] [docx] [pptx] [xlsx]

With --jobs 1 (the default) the blocks run in order. With more, up to N blocks run at once: a block starts once every
earlier block that names one of its files (report.docx, "no-break-{i}.docx"...) is done. A missed link almost always
makes a block fail, so when anything fails, or a block changed one of the input files, the skill runs again in order
in a fresh folder, and only that verdict counts. A link the names cannot show (a file with another extension, a name
built by code) could still let a block pass on a file a later block wrote: the in-order run (the default) has the
last word.

The environment variables OX and PY give the commands behind `ox` and `oxpy` (default: the `oxoffice` on PATH
and this Python).
"""

import argparse
import concurrent.futures
import hashlib
import os
import re
import shlex
import shutil
import signal
import subprocess
import sys
import textwrap
import time
from pathlib import Path

FIXTURES = ("report.docx", "deck.pptx", "budget.xlsx", "Aileron-Regular.ttf", "photo.png")
TIMEOUT = 600   # seconds a block may take
# A file a block reads or writes, as the block spells it: a name with one of these extensions, with or without
# f-string fields or a glob. The module names are not files.
FILE = re.compile(r"[\w.{}*\-]+\.(?:docx|dotx|docm|pptx|potx|pptm|xlsx|xlsm|xltx|xltm|pdf|png|jpg|svg|html|md|json|txt|ttf|xml|csv|tsv)\b")
MODULES = {"oxoffice.docx", "oxoffice.pptx", "oxoffice.xlsx"}


def files_named(code):
    """The names of the files a block names, f-string fields and globs as written, but the inputs."""
    return set(FILE.findall(code)) - MODULES - set(FIXTURES)


def pattern(name):
    """A name as a regular expression: an f-string field or a glob matches any text."""
    return re.compile("".join(".*" if part in ("*",) or part.startswith("{") else re.escape(part)
                              for part in re.split(r"(\{[^}]*\}|\*)", name)) + r"\Z")


def related(a, b):
    return a == b or bool(pattern(a).match(b)) or bool(pattern(b).match(a))


def links(blocks):
    """For each block, the earlier blocks that name one of its files."""
    names = [files_named(code) for _, code in blocks]
    return [[j for j in range(i) if any(related(a, b) for a in names[i] for b in names[j])] for i in range(len(blocks))]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def prepare(work, fixtures):
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    for f in FIXTURES:
        shutil.copy(fixtures / f, work / f)


def run_block(name, k, lang, code, work, env, setup, out=print):
    """Run one block in `work`; return its status line and whether it passed."""
    env = dict(env, BLOCK=code)
    script = setup + ('oxpy -c "$BLOCK"' if lang == "python" else 'eval "$BLOCK"')
    t0 = time.monotonic()
    # a session of its own, so that a timeout stops the block's oxoffice or Python too, not only bash
    proc = subprocess.Popen(["bash", "-euo", "pipefail", "-c", script], cwd=work, env=env, text=True,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)
    try:
        stdout, stderr = proc.communicate(timeout=TIMEOUT)
        status = "ok" if proc.returncode == 0 else f"FAIL {proc.returncode}"
    except subprocess.TimeoutExpired:
        os.killpg(proc.pid, signal.SIGKILL)
        stdout, stderr = proc.communicate()
        status = f"FAIL timeout after {TIMEOUT} s"
    (work / f"_block{k:02d}.log").write_text(f"{code}\n----- stdout\n{stdout}\n----- stderr\n{stderr}")
    line = f"{name}-{k:02d}-{lang}: {status} ({time.monotonic() - t0:.1f} s)"
    if status != "ok":
        line += "\n  " + (stdout[-1500:] + stderr[-2500:]).replace("\n", "\n  ")
    return line, status == "ok"


def in_order(name, blocks, work, env, setup):
    failed = 0
    for k, (lang, code) in enumerate(blocks):
        line, ok = run_block(name, k, lang, code, work, env, setup)
        print(line, flush=True)
        failed += not ok
    return failed


def at_once(name, blocks, work, env, setup, jobs):
    """Run the blocks `jobs` at a time, each once the earlier blocks that name one of its files are done. Return
    the number of blocks that failed or did not run because a block they wait for failed."""
    after = links(blocks)
    state = {}   # block -> True (passed) / False (failed or not run)
    running = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=jobs) as pool:
        while len(state) < len(blocks):
            for k, (lang, code) in enumerate(blocks):
                if k in state or k in running:
                    continue
                if any(state.get(j) is False for j in after[k]):
                    state[k] = False
                    print(f"{name}-{k:02d}-{lang}: not run, a block it waits for failed", flush=True)
                elif all(state.get(j) for j in after[k]) and len(running) < jobs:
                    running[k] = pool.submit(run_block, name, k, lang, code, work, env, setup)
            if not running:
                continue
            done, _ = concurrent.futures.wait(running.values(), return_when=concurrent.futures.FIRST_COMPLETED)
            for k in [k for k, f in running.items() if f in done]:
                line, state[k] = running.pop(k).result()
                print(line, flush=True)
    return sum(not ok for ok in state.values())


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--skills", type=Path, required=True, help="the plugin's skills folder")
    ap.add_argument("--fixtures", type=Path, required=True)
    ap.add_argument("--work", type=Path, required=True)
    ap.add_argument("--jobs", type=int, default=1, help="blocks run at once (default 1: in order)")
    ap.add_argument("names", nargs="*", default=["docx", "pptx", "xlsx"])
    args = ap.parse_args()

    failed = 0
    for name in args.names:
        skill = (args.skills / name).resolve()
        text = (skill / "references" / "examples.md").read_text().replace("\r\n", "\n")
        fenced = re.findall(r"^([ \t]*)```(\S*)\n(.*?)^\1```[ \t]*$", text, re.S | re.M)
        blocks = [(lang, textwrap.dedent(code)) for _, lang, code in fenced]
        others = sorted({lang or "(no language)" for lang, _ in blocks if lang not in ("python", "bash")})
        others += ["~~~"] if re.search(r"^[ \t]*~~~", text, re.M) else []
        if len(re.findall(r"^[ \t]*```", text, re.M)) != 2 * len(fenced):   # a fence closed at another indent
            others += ["unpaired fence"]
        if not blocks or others:
            print(f"{name}: {len(blocks)} blocks, fences that do not run: {others or 'none'}")
            failed += 1
            continue
        work = args.work / name
        env = {k: v for k, v in os.environ.items() if k not in ("PYTHONPATH", "VIRTUAL_ENV")}
        env.update(SKILL=skill.as_posix(), PYTHONPATH=str(skill / "scripts"), PYTHONDONTWRITEBYTECODE="1")
        setup = (f'ox() {{ {os.environ.get("OX", "oxoffice")} "$@"; }}; '
                 f'oxpy() {{ {os.environ.get("PY", shlex.quote(sys.executable))} "$@"; }}\n')
        prepare(work, args.fixtures)
        t0 = time.monotonic()
        if args.jobs > 1:
            inputs = {f: digest(work / f) for f in FIXTURES}
            missed = at_once(name, blocks, work, env, setup, args.jobs)
            changed = [f for f in FIXTURES if not (work / f).exists() or digest(work / f) != inputs[f]]
            if missed or changed:
                why = f"{missed} failed" if missed else f"{', '.join(changed)} changed by a block"
                print(f"{name}: {why} with --jobs {args.jobs}, running the skill again in order", flush=True)
                prepare(work, args.fixtures)
                missed = in_order(name, blocks, work, env, setup)
                if not missed:
                    print(f"{name}: passes in order: a link between blocks was missed with --jobs, see links()")
        else:
            missed = in_order(name, blocks, work, env, setup)
        failed += missed
        print(f"{name}: {len(blocks)} blocks in {time.monotonic() - t0:.0f} s")
    print(f"failed: {failed}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
