# CLAUDE.md

A Claude plugin (`oxoffice`) with two auto-loading skills, `docx` and `pptx`, that make agents use the
[oxoffice](https://github.com/oxoffice-dev/oxoffice) CLI and Python package for every Word and PowerPoint task.
Public repository: nothing specific to a company, a project or a person goes in here.

## Map

- `plugin/`: what users install. Manifest (`.claude-plugin/plugin.json`), skills, icon and README pictures.
- `plugin/skills/<name>/SKILL.md`: what an agent reads first, with the order in which it finds oxoffice. Keep it
  under 16 KiB and its description under 800 characters; details go to `references/` (`cli.md`, `python-api.md`,
  `examples.md`, `install.md`, and `interop.md` for docx).
- `plugin/skills/<name>/references/examples.md`: the examples, run by oxoffice's CI (below).
- `.claude-plugin/marketplace.json`: the marketplace entry. `RELEASE.md`: how a version is published.

## Rules

1. The skills describe one oxoffice release, the one their Setup pins (`oxoffice==0.2.0`). Moving to another
   release changes every place that names it (`grep -rn '0\.2\.0' plugin README.md PRIVACY.md`) and means checking
   every command, flag and Python name the skills cite against it. The skills cite oxoffice `main` and pin its next
   release, 0.2.0, not out yet: they are published once it is (`RELEASE.md`). They hold no script: what an agent
   needs belongs in oxoffice (its `specs/2026-11-20-skill-helpers-in-the-library.md`).
2. No tests and no build here. The examples are the contract with oxoffice: its `skill-examples` job runs every
   block of both `examples.md` on each push to oxoffice's `main`, against the wheel just built, from this
   repository's default branch, and checks that every name the skills cite exists (oxoffice
   `tests/skills/README.md` says how to run it locally). Run it before pushing a change to the skills.
3. An example reads `report.docx` or `deck.pptx` and writes new files; `ox` is the CLI, `oxpy` the Python (shell
   functions, which work in bash and zsh), and blocks run in order in one folder. Known limits come from
   `oxoffice limitations docx|pptx`, not from a page here.
4. Conventional commits, lowercase subject. Never commit or push unless asked. No attribution lines.
5. Never leave a document open in Word, PowerPoint or another GUI app.
