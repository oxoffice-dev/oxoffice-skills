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
- 🔒 **Verified downloads**: the installer checks every download against the release's `SHA256SUMS`, and every
  release carries a signed build provenance.
- 🩹 **Honest about gaps**: known limits are listed with a workaround, or a hand-off to the built-in skills.

## 🚀 Install

In Claude Code:

```
/plugin marketplace add oxoffice-dev/oxoffice-skills
/plugin install oxoffice@oxoffice-skills
```

The first time a skill runs, `plugin/scripts/install.py` downloads the oxoffice release named in
[`plugin/oxoffice.json`](plugin/oxoffice.json) for your platform (Linux, macOS, Windows), checks it and installs it
in its own Python environment under `~/.local/share/oxoffice-skills`. Nothing else to set up.

## 🧩 What is inside

| Path | What |
| --- | --- |
| [`plugin/skills/docx/`](plugin/skills/docx/) | the Word skill: reading, writing, editing, comments, tracked changes, redlines, conversions |
| [`plugin/skills/pptx/`](plugin/skills/pptx/) | the PowerPoint skill: reading, building, editing, rendering, merging decks |
| [`plugin/scripts/install.py`](plugin/scripts/install.py) | the installer, standard library only |
| [`plugin/oxoffice.json`](plugin/oxoffice.json) | the oxoffice release the skills are written for |

## 🔒 Privacy

No server, no account, no telemetry. See [PRIVACY.md](PRIVACY.md).

## 📄 License

[MIT](LICENSE).
