# Installing oxoffice

The skill installs oxoffice by itself (`scripts/install.py`, below). The other channels are for using
oxoffice outside the skill, or when the user prefers one. Every channel carries the same build, and the
`oxoffice` command; the Python channels also carry the `oxoffice` package (`oxoffice.docx`, `oxoffice.pptx`).

| Channel | Command | Gives |
|---|---|---|
| The skill's installer | `python3 "$SKILL/../../scripts/install.py"` | its own environment under `~/.local/share/oxoffice-skills/current`: `bin/oxoffice` and a `bin/python` that imports oxoffice. The wheel of the version in `oxoffice.json`, checked against the release's `SHA256SUMS`. |
| PyPI | `pip install oxoffice`, `uv add oxoffice`, `uv tool install oxoffice` | the package and the command in that environment (`uv tool`: the command only) |
| One-off Python | `uv run --with oxoffice python script.py` | a throwaway environment for one script |
| conda (prefix.dev) | `pixi global install -c https://prefix.dev/hadim/oxoffice -c conda-forge oxoffice`, or `conda install -c https://repo.prefix.dev/hadim/oxoffice -c conda-forge oxoffice` | the command (`pixi global`), or the package in a conda environment |
| Shell installer, Linux and macOS | `curl -fsSL https://github.com/oxoffice-dev/oxoffice/releases/latest/download/install.sh \| sh` | the native `oxoffice` binary in `~/.oxoffice/bin`, no Python |
| Shell installer, Windows | `powershell -ExecutionPolicy Bypass -c "irm https://github.com/oxoffice-dev/oxoffice/releases/latest/download/install.ps1 \| iex"` | the same, in `%USERPROFILE%\.oxoffice\bin` |
| GitHub release | https://github.com/oxoffice-dev/oxoffice/releases | every archive, wheel and checksum, by hand |

Pin a version with `oxoffice==X.Y.Z` (PyPI, conda) or `OXOFFICE_VERSION=X.Y.Z` (shell installers).

Check a file downloaded by hand before using it:

```bash
sha256sum -c SHA256SUMS --ignore-missing
gh attestation verify <file> --repo oxoffice-dev/oxoffice
```

`oxoffice --version` shows the installed version. The skill's references describe the version in
`../../oxoffice.json`; with another version, a cited command may differ (`oxoffice <family> <command> --help`).
