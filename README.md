<div align="center">

<img src="plugin/assets/icon.svg" alt="oxoffice-skills icon" width="128">

# oxoffice-skills

**Word and PowerPoint skills for AI agents: read, write, edit, review and render `.docx` and `.pptx` files with
[oxoffice](https://github.com/oxoffice-dev/oxoffice).**

[![License: MIT](https://img.shields.io/badge/license-MIT-22c55e)](LICENSE)
[![Platforms](https://img.shields.io/badge/platforms-Linux%20%7C%20macOS%20%7C%20Windows-0ea5a4)](#-install)
[![Built on oxoffice](https://img.shields.io/badge/built%20on-oxoffice-4f46e5)](https://github.com/oxoffice-dev/oxoffice)
[![Skills](https://img.shields.io/badge/skills-docx%20%C2%B7%20pptx-f97316)](plugin/skills/)

</div>

| ![A redline of two versions of a report, rendered by oxoffice](plugin/assets/screenshot-docx.png) | ![A slide built and rendered by oxoffice](plugin/assets/screenshot-pptx.png) |
|:---:|:---:|
| 📝 A redline of two versions of a report, in reviewer colours | 📊 A slide built and rendered from code |

<p align="center"><sub>Both pictures come straight out of oxoffice: no Word, no PowerPoint, no LibreOffice.</sub></p>

Give your agent real Office skills. This plugin teaches it to work on Word and PowerPoint files through the
`oxoffice` command line and the `oxoffice` Python package, instead of LibreOffice round trips, python-docx,
python-pptx or hand-written XML.

## ✨ Why oxoffice-skills

- ⚡ **Native and fast**: one `oxoffice` call for whole-file work, the Python package for anything finer. No office
  suite to install.
- 🎯 **Loads by itself**: the skills step in as soon as a `.docx` or `.pptx` shows up, even when you don't name a
  tool.
- ✍️ **Safe edits**: counted find and replace across runs, text edited without losing its formatting, and the input
  never written over.
- 🔍 **Real tracked changes**: list, accept and reject revisions, and build a redline of two versions.
- 📌 **Stable, not frozen**: the skills are written and checked with oxoffice 0.3.0 and run any later release,
  so a new oxoffice does not wait for a new plugin; CI runs every example against both, every night.
- 🧪 **Examples that run**: every example of the skills runs on each push to oxoffice's `main`, against the build
  of that commit.
- 🩹 **Honest about gaps**: `oxoffice limitations` lists what the build cannot do yet, each with a workaround, and
  the skills hand that one step to the built-in skills.

## 🚀 Install

In Claude Code:

```
/plugin marketplace add oxoffice-dev/oxoffice-skills
/plugin install oxoffice@oxoffice-skills
```

oxoffice is the one prerequisite, 0.3.0 or later (the skills are written and checked with 0.3.0). The skills use
the `oxoffice` already installed, else get it through `uv`, `pip`, `pixi` or `conda`. Without any of them, they
tell you how to install it from its [official instructions](https://github.com/oxoffice-dev/oxoffice#-install) and
fall back to the built-in Word and PowerPoint skills meanwhile.

oxoffice itself is also available on its own, for use outside the skills:

| Channel | Command |
| --- | --- |
| 🐍 PyPI | `pip install oxoffice`, `uv add oxoffice` or `uv tool install oxoffice` |
| 🟢 conda | `pixi global install -c https://prefix.dev/hadim/oxoffice -c conda-forge oxoffice` |
| 📦 GitHub release | the [releases page](https://github.com/oxoffice-dev/oxoffice/releases), with checksums and a signed build provenance |
| 🐚 Installers | the shell and PowerShell installers, in oxoffice's own [install instructions](https://github.com/oxoffice-dev/oxoffice#-install) |

Details in [`plugin/skills/docx/references/install.md`](plugin/skills/docx/references/install.md).

## 🧩 What is inside

| Path | What |
| --- | --- |
| [`plugin/skills/docx/`](plugin/skills/docx/) | the Word skill: reading, writing, editing, comments, tracked changes, redlines, conversions |
| [`plugin/skills/pptx/`](plugin/skills/pptx/) | the PowerPoint skill: reading, building, editing, rendering, merging decks |
| `plugin/skills/*/references/` | the CLI and Python references, the examples, how to run oxoffice |

## 🔒 Privacy

No server, no account, no telemetry. Your documents never leave your machine: oxoffice runs locally and makes
no network request. The one download is oxoffice itself, once, when no copy is found: from PyPI through uv or
pip, or from prefix.dev through pixi or conda, and it carries none of your data.
See [PRIVACY.md](PRIVACY.md).

## 📄 License

[MIT](LICENSE).
