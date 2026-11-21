# Running oxoffice

The skills are written for **oxoffice 0.1.0**. Every channel carries the same build and the `oxoffice` command;
the Python channels also carry the `oxoffice` package (`oxoffice.docx`, `oxoffice.pptx`). SKILL.md's Setup tries
them in the order below and stops at the first that works. Each way ends with two shell variables, `OX` (the
command) and `PY` (a Python that imports oxoffice), to set again at the start of every shell command.

| Way | `OX` and `PY` | Notes |
|---|---|---|
| 1. Already installed | `OX=oxoffice; PY=python3` | only when `oxoffice --version` and `oxoffice.__version__` both say 0.1.0; never upgrade or downgrade the user's install |
| 2. uv | `OX="uvx --from oxoffice==0.1.0 oxoffice"; PY="uv run --no-project --with oxoffice==0.1.0 python"` | nothing installed: uv caches the wheel; `--no-project` keeps the user's project out of it |
| 3. pip in a virtual environment | `V="${XDG_CACHE_HOME:-$HOME/.cache}/oxoffice-0.1.0"`, then `python3 -m venv "$V" && "$V/bin/pip" install oxoffice==0.1.0` once; `OX="$V/bin/oxoffice"; PY="$V/bin/python"` | a folder of its own, outside the user's folder |
| 4. pixi | `X="pixi exec -c https://prefix.dev/hadim/oxoffice -c conda-forge -s oxoffice==0.1.0"; OX="$X oxoffice"; PY="$X python"` | the conda packages of the `hadim/oxoffice` channel on prefix.dev, Python from conda-forge |
| 4. conda | `conda create -y -p "$V" -c https://repo.prefix.dev/hadim/oxoffice -c conda-forge oxoffice=0.1.0` once; then as in 3 | same packages |
| 5. The CLI alone | `curl -fsSL https://github.com/oxoffice-dev/oxoffice/releases/download/v0.1.0/install.sh \| OXOFFICE_VERSION=0.1.0 OXOFFICE_NO_PATH_UPDATE=1 sh`; `OX=~/.oxoffice/bin/oxoffice` | no Python: the steps that need `PY` fall back to the built-in skill |
| 6. Nothing works | | tell the user in one line, do the task with the built-in skill |

**Windows**, in Git Bash: `python` (or `py -3`) for `python3`, and `$V/Scripts/` for `$V/bin/` in 3 and 4;
`$OX` reaches `oxoffice.exe`. Way 5 runs in PowerShell:

```powershell
$env:OXOFFICE_VERSION = "0.1.0"; $env:OXOFFICE_NO_PATH_UPDATE = "1"
irm https://github.com/oxoffice-dev/oxoffice/releases/download/v0.1.0/install.ps1 | iex
```

then `OX=~/.oxoffice/bin/oxoffice.exe` in Git Bash.

The shell installers check the archive against the release's `SHA256SUMS`, and its build provenance when the
GitHub CLI is signed in. The wheels on PyPI carry attestations, the conda packages are published through
prefix.dev's trusted publishing. A file downloaded by hand from the
[release](https://github.com/oxoffice-dev/oxoffice/releases/tag/v0.1.0) is checked with:

```bash
sha256sum -c SHA256SUMS --ignore-missing
gh attestation verify <file> --repo oxoffice-dev/oxoffice
```

Another version than 0.1.0 may differ from what the references cite: check a command with
`$OX <docx|pptx> <command> --help`, a Python name in the type stubs.
