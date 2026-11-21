# Privacy

oxoffice-skills is a Claude plugin: two skills, their references and a few helper scripts. It has no server, no
account and no telemetry.

## What the plugin collects

Nothing. It reads the Word and PowerPoint files you point it at, on your machine, and writes the files you ask
for there too. It keeps no copy of your documents and records nothing about you or your use of it.

## What leaves your machine

- **Nothing from your documents.** oxoffice runs locally and makes no network request.
- **One download per oxoffice version, when it is missing.** The skills use the oxoffice 0.2.0 already on your
  machine; otherwise they fetch it once from the channel they find (PyPI through `uv` or `pip`, the prefix.dev conda
  channel through `pixi` or `conda`, or the [oxoffice releases](https://github.com/oxoffice-dev/oxoffice/releases)
  on GitHub). That request carries none of your data, and the service's own privacy statement applies to it.
- **Your conversation with Claude.** When Claude reads a document through these skills, the text it reads becomes
  part of your conversation, which the Claude product you use handles under its own terms and privacy policy.
  The plugin adds no other destination.
