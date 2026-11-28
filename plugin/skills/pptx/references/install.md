# Running oxoffice

oxoffice is a prerequisite of the skills, 0.10.0 or later (they are written and checked with 0.10.0). They use the
one installed, else get it through a package manager. Every channel carries the same build and the `oxoffice`
command; the Python channels also carry the `oxoffice` package (`oxoffice.docx`, `oxoffice.pptx`, `oxoffice.xlsx`). SKILL.md's Setup
tries them in the order below and stops at the first that works. Each way defines two shell functions, `ox` (the
command) and `oxpy` (a Python that imports oxoffice), to define again at the start of every shell command:
functions behave the same in bash and zsh, and a path with spaces stays one word inside them. Each function holds
its whole path, so it does not depend on a variable set beside it, and `oxpy` sets `PYTHONUTF8=1` so that files and
pipes are read and written in UTF-8 on Windows too.

1. **Already installed**, only when `oxoffice --version` and `oxoffice.__version__` both say 0.10.0 or later
   (never upgrade or downgrade the user's install):

   ```bash
   ox() { oxoffice "$@"; }; oxpy() { PYTHONUTF8=1 python3 "$@"; }
   ```

2. **uv**: nothing installed, uv caches the wheel; `--no-project` keeps the user's project out of it, and
   `--python '>=3.10'` skips an older Python (uv fetches one if needed):

   ```bash
   ox() { uvx --python '>=3.10' --from 'oxoffice>=0.10.0' oxoffice "$@"; }
   oxpy() { PYTHONUTF8=1 uv run --no-project --python '>=3.10' --with 'oxoffice>=0.10.0' python "$@"; }
   ```

3. **pip in a virtual environment** of its own, outside the user's folder (the Python must be 3.10 or later):

   ```bash
   V="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-venv"
   [ -x "$V/bin/oxoffice" ] || { python3 -m venv "$V" && "$V/bin/pip" install -q 'oxoffice>=0.10.0'; }
   ox() { "${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-venv/bin/oxoffice" "$@"; }
   oxpy() { PYTHONUTF8=1 "${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-venv/bin/python" "$@"; }
   ```

4. **pixi**, from the `hadim/oxoffice` channel on prefix.dev with Python from conda-forge (`python -m oxoffice`
   rather than the `oxoffice` launcher, which breaks when the cache path holds a space):

   ```bash
   ox() { pixi exec -c https://prefix.dev/hadim/oxoffice -c conda-forge -s 'oxoffice>=0.10.0' python -m oxoffice "$@"; }
   oxpy() { PYTHONUTF8=1 pixi exec -c https://prefix.dev/hadim/oxoffice -c conda-forge -s 'oxoffice>=0.10.0' python "$@"; }
   ```

   **conda**, in an environment of its own:

   ```bash
   C="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-conda"
   [ -x "$C/bin/python" ] || conda create -y -q -p "$C" --override-channels \
       -c https://repo.prefix.dev/hadim/oxoffice -c conda-forge 'oxoffice>=0.10.0'
   ox() { "${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-conda/bin/python" -m oxoffice "$@"; }
   oxpy() { PYTHONUTF8=1 "${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-conda/bin/python" "$@"; }
   ```

5. **None works**: tell the user in one line that oxoffice is missing, with its official install methods
   ([install instructions](https://github.com/oxoffice-dev/oxoffice#-install)), and do the task with the built-in skill. Installing it another
   way is the user's call.

**Windows**, in Git Bash: `python` (or `py -3`) for `python3`; `Scripts` for `bin` in 3 (`pip`, `oxoffice` and
`python`); `/python` for `/bin/python` in 4 (conda), the `[ -x ... ]` check included.

The wheels on PyPI carry attestations, the conda packages are published through prefix.dev's trusted
publishing. A file downloaded by hand from a
[release](https://github.com/oxoffice-dev/oxoffice/releases) is checked with:

```bash
sha256sum -c SHA256SUMS --ignore-missing
gh attestation verify <file> --repo oxoffice-dev/oxoffice
```

The references cite what 0.10.0 has. A later version keeps it as long as oxoffice's API stays stable, but may
differ: when a cited command or name fails, check it with `ox <docx|pptx|xlsx> <command> --help`, a Python name in the
type stubs, and give `ox --version` with any report of a bug, which may come from that version.
