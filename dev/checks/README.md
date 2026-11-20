# Checks used to port the skills (temporary)

The scripts that validated the port of the skills to oxoffice 0.1.0. They move to
the oxoffice repository, where the recipes run on every push to `main` against the
wheel just built (see oxoffice `specs/2026-11-20-test-suite-from-scratch.md`).
This folder goes away when they do, since this repository has no tests.

- `run_recipes.py`: runs every python and bash block of `references/recipes.md`, one
  folder per file, starting from `fixtures/` (`report.docx`, `deck.pptx`, a free font).
- `check_outputs.py`: the assertions on the recipes' outputs.
- `check_names.py`, `check_attrs.py`: every CLI command, flag, Python name, keyword and
  attribute the skills cite exists in the installed build.
- `fixtures/make_fixture_docx.py`, `make_fixture_pptx.py`: build the fixtures (Pillow,
  python-pptx). The font is Pillow's bundled Aileron.

They expect `OXOFFICE_HOME` to point at an install made by `plugin/scripts/install.py`
(`--wheel` for a local wheel).
