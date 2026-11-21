# Running oxoffice

The skills are written for **oxoffice 0.1.0**. Every channel carries the same build and the `oxoffice` command;
the Python channels also carry the `oxoffice` package (`oxoffice.docx`, `oxoffice.pptx`). SKILL.md's Setup tries
them in the order below and stops at the first that works. Each way defines two shell functions, `ox` (the
command) and `oxpy` (a Python that imports oxoffice), to define again at the start of every shell command:
functions behave the same in bash and zsh, and a path with spaces stays one word inside them.

1. **Already installed**, only when `oxoffice --version` and `oxoffice.__version__` both say 0.1.0 (never upgrade
   or downgrade the user's install):

   ```bash
   ox() { oxoffice "$@"; }; oxpy() { python3 "$@"; }
   ```

2. **uv**: nothing installed, uv caches the wheel; `--no-project` keeps the user's project out of it, and
   `--python '>=3.10'` skips an older Python (uv fetches one if needed):

   ```bash
   ox() { uvx --python '>=3.10' --from oxoffice==0.1.0 oxoffice "$@"; }
   oxpy() { uv run --no-project --python '>=3.10' --with oxoffice==0.1.0 python "$@"; }
   ```

3. **pip in a virtual environment** of its own, outside the user's folder (the Python must be 3.10 or later):

   ```bash
   V="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-0.1.0"
   [ -x "$V/bin/oxoffice" ] || { python3 -m venv "$V" && "$V/bin/pip" install -q oxoffice==0.1.0; }
   ox() { "$V/bin/oxoffice" "$@"; }; oxpy() { "$V/bin/python" "$@"; }
   ```

4. **pixi**, from the `hadim/oxoffice` channel on prefix.dev with Python from conda-forge (`python -m oxoffice`
   rather than the `oxoffice` launcher, which breaks when the cache path holds a space):

   ```bash
   ox() { pixi exec -c https://prefix.dev/hadim/oxoffice -c conda-forge -s oxoffice==0.1.0 python -m oxoffice "$@"; }
   oxpy() { pixi exec -c https://prefix.dev/hadim/oxoffice -c conda-forge -s oxoffice==0.1.0 python "$@"; }
   ```

   **conda**, in an environment of its own:

   ```bash
   C="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-0.1.0-conda"
   [ -x "$C/bin/python" ] || conda create -y -q -p "$C" --override-channels \
       -c https://repo.prefix.dev/hadim/oxoffice -c conda-forge oxoffice=0.1.0
   ox() { "$C/bin/python" -m oxoffice "$@"; }; oxpy() { "$C/bin/python" "$@"; }
   ```

5. **The CLI alone**, no Python (the steps that need `oxpy` fall back to the built-in skill), installed in a folder
   of its own so that an `oxoffice` the user installed is left alone:

   ```bash
   H="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-0.1.0-cli"
   [ -x "$H/bin/oxoffice" ] || curl -fsSL https://github.com/oxoffice-dev/oxoffice/releases/download/v0.1.0/install.sh \
       | OXOFFICE_VERSION=0.1.0 OXOFFICE_HOME="$H" OXOFFICE_NO_PATH_UPDATE=1 sh
   ox() { "$H/bin/oxoffice" "$@"; }
   ```

6. **Nothing works**: tell the user in one line, and do the task with the built-in skill.

**Windows**, in Git Bash: `python` (or `py -3`) for `python3`; `$V/Scripts/oxoffice` and `$V/Scripts/python` in 3;
`$C/python` for `$C/bin/python` in 4 (conda). Way 5 runs in PowerShell:

```powershell
$env:OXOFFICE_VERSION = "0.1.0"; $env:OXOFFICE_NO_PATH_UPDATE = "1"
$env:OXOFFICE_HOME = "$env:LOCALAPPDATA\oxoffice-0.1.0-cli"
irm https://github.com/oxoffice-dev/oxoffice/releases/download/v0.1.0/install.ps1 | iex
```

then, in Git Bash, `ox() { "$LOCALAPPDATA/oxoffice-0.1.0-cli/bin/oxoffice.exe" "$@"; }`.

The shell installers check the archive against the release's `SHA256SUMS`, and its build provenance when the
GitHub CLI is signed in. The wheels on PyPI carry attestations, the conda packages are published through
prefix.dev's trusted publishing. A file downloaded by hand from the
[release](https://github.com/oxoffice-dev/oxoffice/releases/tag/v0.1.0) is checked with:

```bash
sha256sum -c SHA256SUMS --ignore-missing
gh attestation verify <file> --repo oxoffice-dev/oxoffice
```

Another version than 0.1.0 may differ from what the references cite: check a command with
`ox <docx|pptx> <command> --help`, a Python name in the type stubs.
