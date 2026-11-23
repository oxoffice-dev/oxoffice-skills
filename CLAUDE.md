# CLAUDE.md

A Claude plugin (`oxoffice`) with two auto-loading skills, `docx` and `pptx`, that make agents use the
[oxoffice](https://github.com/oxoffice-dev/oxoffice) CLI and Python package for every Word and PowerPoint task.
Public repository: nothing specific to a company, a project or a person goes in here.

## Map

- `plugin/`: what users install. Manifest (`.claude-plugin/plugin.json`), skills, icon and README pictures.
- `plugin/skills/<name>/SKILL.md`: what an agent reads first, with the order in which it finds oxoffice. Keep it
  under 16 KiB and its description under 800 characters; details go to `references/` (`cli.md`, `python-api.md`,
  `examples.md`, `install.md`, and `interop.md` for docx).
- `plugin/skills/<name>/references/examples.md`: the examples, run by this repository's CI (below).
- `tests/`: the checks of the examples and of the install ways (`tests/README.md`), run by `.github/workflows/ci.yml`.
- `.claude-plugin/marketplace.json`: the marketplace entry. `RELEASE.md`: how a version is published.

## Rules

1. The skills are written and checked with one oxoffice release, the floor their Setup names (`oxoffice>=0.3.0`),
   and take any later one: a new oxoffice release needs no new plugin while its API stays stable. CI runs the
   examples against the floor and the latest release. Raise the floor only when the skills start citing something
   an older release lacks: that changes every place that names it (`grep -rn '0\.2\.0' plugin README.md
   PRIVACY.md`) and means checking every command, flag and Python name the skills cite, then publishing per
   `RELEASE.md`. They hold no script: what an agent needs belongs in oxoffice (its `specs/2026-11-20-skill-helpers-in-the-library.md`).
2. No build here. The examples are the contract with oxoffice: this repository's CI runs every block of both
   `examples.md` against the floor and the latest release, checks that every name the skills cite exists, and
   installs oxoffice by every way `install.md` documents on Linux, macOS and Windows (conda where the runner has it), on
   each push, pull request and night
   (`tests/README.md` says how to run it locally). Run it before pushing a change to the skills.
3. An example reads `report.docx` or `deck.pptx` and writes new files; `ox` is the CLI, `oxpy` the Python (shell
   functions, which work in bash and zsh), and blocks run in order in one folder. Known limits come from
   `oxoffice limitations docx|pptx`, not from a page here.
4. Conventional commits, lowercase subject. Never commit or push unless asked. No attribution lines.
5. Never leave a document open in Word, PowerPoint or another GUI app.
