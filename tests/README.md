# Tests

The skills are documentation and examples. Two promises are checked by `.github/workflows/ci.yml`, against the
oxoffice release the skills were written with, their floor (`oxoffice>=X` in their Setup), and against the latest
release on PyPI (an oxoffice change that breaks an example shows here once released), on every push to `main`,
every pull request and every night:

| File | What |
| --- | --- |
| `check_install.py` | runs every way of installing oxoffice that `references/install.md` documents, as written (Git Bash substitutions on Windows), and checks `ox --version` and `oxpy` give the floor or a later version; Linux, macOS, Windows (conda only where the runner has it, with a warning elsewhere) |
| `make_fixtures.py` | writes the inputs of the examples: `report.docx`, `deck.pptx`, `Aileron-Regular.ttf`, `photo.png` |
| `fixtures/make_fixture_{docx,pptx}.py` | the synthetic report and deck, with the traits of files saved by Word and Google Docs |
| `run_examples.py` | runs every python and bash block of each skill's `references/examples.md`, in order, or `--jobs N` at a time |
| `check_outputs.py` | checks the files the examples wrote |
| `check_names.py` | checks that every function, keyword, attribute, command and flag the skills cite exists |

Locally, in a new virtual environment (about 20 seconds for the examples, a few minutes for the install ways):

```bash
pip install --require-hashes -r tests/requirements.txt
pip install oxoffice==0.5.0            # the floor (or oxoffice, the latest), in a second call: the first checks hashes
python tests/make_fixtures.py /tmp/skills/fixtures
python tests/run_examples.py --jobs 4 --skills plugin/skills --fixtures /tmp/skills/fixtures --work /tmp/skills
python tests/check_outputs.py /tmp/skills
python tests/check_names.py plugin/skills
python tests/check_install.py plugin/skills   # installs through uv, pip, pixi and conda
```

`run_examples.py` runs each skill in one fresh folder that starts with the fixtures. With `--jobs`, a block starts
once every earlier block that names one of its files is done; any failure, or a block that changed a fixture, runs
the skill again in order in a fresh folder, and only that run counts. Each block's code and output are in
`<work>/<skill>/_blockNN.log`, uploaded by the job when it fails. `OX` and `PY` give the commands behind `ox` and
`oxpy` (default: the `oxoffice` on PATH and this Python): set them to a way of the Setup to run the examples
through it.

A failure means one of three things: the release broke what the skills promise (say so in oxoffice, and adapt the
skills), the skills cite something that does not exist (fix the skills), or a documented install way stopped
working (fix `install.md`, or the release's packages). `check_outputs.py` knows what each example writes: an example
whose output changes changes it too. `check_names.py` matches a method name against every class of both packages,
and an attribute against every word of their sources: it catches a name that no longer exists anywhere, not one moved
to another class.

`requirements.txt` is compiled from `requirements.in` with hashes: `uv pip compile --universal --python-version 3.12
--generate-hashes tests/requirements.in -o tests/requirements.txt`.
