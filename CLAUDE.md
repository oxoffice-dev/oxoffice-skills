# CLAUDE.md

A Claude plugin (`oxoffice`) with two auto-loading skills, `docx` and `pptx`, that make agents use the
[oxoffice](https://github.com/oxoffice-dev/oxoffice) CLI and Python package for every Word and PowerPoint task.
Public repository: nothing specific to a company, a project or a person goes in here.

## Map

- `plugin/`: what users install. Manifest, skills, `scripts/install.py`, `oxoffice.json` (the oxoffice
  release the skills are written for), icon and README pictures.
- `plugin/skills/<name>/SKILL.md`: what an agent reads first. Keep it under 16 KB and its description under
  800 characters, details go to `references/` (`cli.md`, `python-api.md`, `recipes.md`, `gaps.md`, and
  `interop.md` for docx).
- `plugin/skills/<name>/scripts/`: helpers importable from Python and runnable as commands.
- `.claude-plugin/marketplace.json`: the marketplace entry.

## Rules

1. The skills describe one oxoffice release, the one in `plugin/oxoffice.json`. Moving to a new release means
   checking every command, flag and Python name the skills cite against it, and updating `gaps.md`.
2. No tests, no build and no release process here: oxoffice is built, tested and released in its own
   repository. Skill recipes are checked by hand against the pinned release before a change.
3. Conventional commits, lowercase subject. Never commit or push unless asked. No attribution lines.
4. Never leave a document open in Word, PowerPoint or another GUI app.
