# Privacy

oxoffice-skills is a Claude plugin: two skills and the scripts they call. It has no server, no account and no
telemetry.

## What the plugin collects

Nothing. It reads the Word and PowerPoint files you point it at, on your machine, and writes the files you ask
for there too. It keeps no copy of your documents and records nothing about you or your use of it.

## What leaves your machine

- **Nothing from your documents.** oxoffice runs locally and makes no network request.
- **One download per oxoffice version.** The first time a skill runs, `scripts/install.py` downloads the oxoffice
  wheel for your platform from the [oxoffice releases](https://github.com/oxoffice-dev/oxoffice/releases) and
  checks it against the release's `SHA256SUMS`. That request reaches GitHub like any download from github.com and
  carries none of your data. GitHub's privacy statement applies to it.
- **Your conversation with Claude.** When Claude reads a document through these skills, the text it reads becomes
  part of your conversation, which the Claude product you use handles under its own terms and privacy policy.
  The plugin adds no other destination.
