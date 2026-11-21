# Publishing a version

The plugin is published from `main` through the marketplace of this repository
(`/plugin marketplace add oxoffice-dev/oxoffice-skills`). Versions are dates, `YYYY.M.D` (`2026.11.21`), and a
second one on the same day adds a number (`2026.11.21.2`). The oxoffice version the skills run is named in their
Setup and in the README, not in the plugin version.

## Before

- [ ] The oxoffice release the skills pin is out: PyPI `oxoffice==X`, prefix.dev `hadim/oxoffice`, the GitHub
      release with `install.sh` and `install.ps1`. `uvx --from oxoffice==X oxoffice --version` prints it.
- [ ] The skills cite only what that release has, and pin it everywhere (`grep -rn 'oxoffice==' plugin README.md`).
- [ ] oxoffice's `skill-examples` job passed on `main` with this repository's head (its log names the commit).
- [ ] The examples pass against the released package, not only `main`'s build: from a checkout of oxoffice at
      the release tag, in a venv with its `.github/requirements/ci.txt` and `oxoffice==X` from PyPI, the four
      commands of its `tests/skills/README.md`; then `run_examples.py` once more with `OX` and `PY` set to the
      Setup's `uvx` and `uv run` commands, which checks that way of the Setup too.
- [ ] `claude plugin validate .` and `claude plugin validate plugin` pass.
- [ ] Each `SKILL.md` is under 16 KiB, each description under 800 characters.

## Steps

1. Set `version` in `plugin/.claude-plugin/plugin.json` to the date.
2. Merge into `main` (squash, a conventional title). The first time only: create `main` at the first commit of
   `wip/port-skills`, open the pull request from `wip/port-skills` into it, make `main` the default branch after
   the merge, then delete `wip/port-skills`.
3. Tag the merge commit `vYYYY.M.D` and publish a GitHub release: what changed for an agent, the oxoffice version.
4. In a fresh Claude Code session: `/plugin marketplace add oxoffice-dev/oxoffice-skills`, `/plugin install
   oxoffice@oxoffice-skills`, then one Word task and one PowerPoint task, each on a file the skill has not seen.
