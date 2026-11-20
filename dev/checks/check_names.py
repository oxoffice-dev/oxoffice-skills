"""Port of rdocx-skills tests/test_skill_docs.py to oxoffice: every Python name and keyword argument and every CLI
command and flag the skills cite exists in the installed oxoffice; then, for information, the public classes and CLI
commands no skill cites. Run with the oxoffice Python ($O/python)."""
import ast
import builtins
import functools
import importlib
import json
import pkgutil
import re
import subprocess
import sys
from pathlib import Path

import oxoffice
import oxoffice.docx
import oxoffice.pptx

SKILLS = Path(__file__).resolve().parents[2] / "plugin" / "skills"
BIN = Path(sys.executable).parent
DOCS = sorted(SKILLS.glob("*/SKILL.md")) + sorted(SKILLS.glob("*/references/*.md"))
FOREIGN = {"Presentation", "Document", "fromstring", "tostring", "xpath", "related_part", "ZipFile"}
stats = {"spans": 0, "calls": 0, "keywords": 0, "cli_spans": 0, "flags": 0}


def rel(path):
    return path.relative_to(SKILLS)


def spans(path):
    """(line, code span) for inline spans outside fences, plus each `oxoffice ...` line of a fenced bash block."""
    fenced = False
    for n, line in enumerate(path.read_text().splitlines(), 1):
        if line.startswith("```"):
            fenced = not fenced
            continue
        if fenced:
            s = line.split(" #")[0].strip()
            for prefix in ("$O/oxoffice ", "oxoffice "):
                if s.startswith(prefix):
                    yield n, "oxoffice " + s[len(prefix):], True
        else:
            yield from ((n, s, False) for s in re.findall(r"`([^`]+)`", line))


@functools.lru_cache(maxsize=None)
def python_api():
    api = {}
    pkg = Path(oxoffice.__file__).parent
    sources = [f for f in pkg.rglob("*.py*") if f.suffix in (".py", ".pyi")]
    for f in sources + sorted(SKILLS.glob("*/scripts/*.py")):
        for node in ast.walk(ast.parse(f.read_text())):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                a = node.args
                names = {x.arg for x in a.posonlyargs + a.args + a.kwonlyargs} - {"self", "cls"}
                api.setdefault(node.name, []).append(names | ({"**"} if a.kwarg else set()))
            elif isinstance(node, ast.ClassDef):
                api.setdefault(node.name, []).append(None)
            elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
                api.setdefault(node.target.id, []).append(None)
            elif isinstance(node, ast.Assign):
                for t in node.targets:
                    if isinstance(t, ast.Name):
                        api.setdefault(t.id, []).append(None)
    return api


def calls(span):
    for m in re.finditer(r"(?<![\w])([A-Za-z_]\w*)\(", span):
        depth, args, cur = 0, [], ""
        for ch in span[m.end():]:
            if ch in "([{":
                depth += 1
            elif ch in ")]}":
                if depth == 0:
                    break
                depth -= 1
            if ch == "," and depth == 0:
                args.append(cur.strip())
                cur = ""
            else:
                cur += ch
        args.append(cur.strip())
        yield m.group(1), [a for a in args if a]


def python_problems(path):
    api = python_api()
    for n, span, fenced in spans(path):
        if fenced or re.match(r"(oxoffice (docx|pptx)|python3?|pip|bash|soffice) |\$", span):
            continue
        stats["spans"] += 1
        for name, args in calls(span):
            if hasattr(builtins, name) or name in FOREIGN:
                continue
            stats["calls"] += 1
            if name not in api:
                yield f"{rel(path)}:{n}: no {name}: `{span}`"
                continue
            sigs = [s for s in api[name] if s is not None]
            for a in args:
                kw = re.match(r"([a-z_]\w*)=(?!=)", a)
                if kw:
                    stats["keywords"] += 1
                if sigs and kw and not any(kw.group(1) in s or "**" in s for s in sigs):
                    yield f"{rel(path)}:{n}: {name}() takes no keyword {kw.group(1)}: `{span}`"


@functools.lru_cache(maxsize=None)
def cli_help(*command):
    res = subprocess.run([str(BIN / "oxoffice"), *command, "--help"], capture_output=True, text=True)
    return res.stdout if res.returncode == 0 else None


def subcommands(help_text):
    section = help_text.split("Commands:", 1)[1].split("\n\n", 1)[0] if "Commands:" in help_text else ""
    return set(re.findall(r"^  (\S+)", section, re.M)) - {"help"}


def cli_problems(path):
    for n, span, _ in spans(path):
        tokens = span.replace("[", " ").replace("]", " ").split()
        if len(tokens) < 2 or tokens[0] != "oxoffice" or tokens[1] not in ("docx", "pptx"):
            continue
        stats["cli_spans"] += 1
        commands, i = [(tokens[1],)], 2
        while i < len(tokens):
            known = set.union(*(subcommands(cli_help(*c) or "") for c in commands))
            words = tokens[i].split("/")
            if not known or not all(w in known for w in words):
                break
            commands = [c + (w,) for c in commands for w in words if w in subcommands(cli_help(*c) or "")]
            i += 1
        if len(commands) == 1 and len(commands[0]) == 1 and len(tokens) > 2 and not tokens[2].startswith("-"):
            if re.fullmatch(r"[a-z]+", tokens[2]) and tokens[2] not in ("COMMAND",):
                yield f"{rel(path)}:{n}: no command oxoffice {tokens[1]} {tokens[2]}: `{span}`"
                continue
        helps = [cli_help(*c) or "" for c in commands]
        for token in tokens[i:]:
            for flag in token.split("/"):
                flag = flag.split("=")[0]
                if re.fullmatch(r"--?[A-Za-z][\w-]*", flag):
                    stats["flags"] += 1
                    if not any(re.search(rf"(?<![\w-]){re.escape(flag)}(?![\w-])", h) for h in helps):
                        yield f"{rel(path)}:{n}: oxoffice {' '.join(commands[0])} has no {flag}: `{span}`"


problems = []
for path in DOCS:
    problems += list(python_problems(path))
    problems += list(cli_problems(path))
print("\n".join(problems) or "no problem")
print(stats)

# ---- reverse, for information: public classes and CLI commands that no skill cites
NOT_DOCUMENTED = {
    "oxoffice.docx.TabStops", "oxoffice.docx.CellCollection", "oxoffice.docx.CellParagraphCollection",
    "oxoffice.docx.ParagraphCollection", "oxoffice.docx.RowCollection", "oxoffice.docx.RunCollection",
    "oxoffice.docx.TableCollection", "oxoffice.pptx.AdjustmentCollection", "oxoffice.pptx.ColumnCollection",
    "oxoffice.pptx.ParagraphCollection", "oxoffice.pptx.PlaceholderCollection", "oxoffice.pptx.RowCollection",
    "oxoffice.pptx.RunCollection", "oxoffice.pptx.SlideCollection", "oxoffice.pptx.SlideLayoutCollection",
    "oxoffice.pptx.Background", "oxoffice.pptx.Column", "oxoffice.pptx.Font", "oxoffice.pptx.Hyperlink",
    "oxoffice.pptx.Image", "oxoffice.pptx.Row", "oxoffice.pptx.ShapeClickAction", "oxoffice.pptx.ShapeHyperlink",
    "oxoffice.pptx.SlideLayout", "oxoffice.pptx.SlideMasterCollection",
}


def skill_code(skill):
    out = []
    for path in [SKILLS / skill / "SKILL.md", *sorted((SKILLS / skill / "references").glob("*.md"))]:
        for k, part in enumerate(re.split(r"^```.*$", path.read_text(), flags=re.M)):
            out += [part] if k % 2 else re.findall(r"`([^`]+)`", part)
    return "\n".join(out)


def python_surface(module):
    classes = {}
    for info in [None, *pkgutil.walk_packages(module.__path__, module.__name__ + ".")]:
        if info and info.name.endswith("._native"):
            continue
        mod = importlib.import_module(info.name) if info else module
        for name, obj in vars(mod).items():
            if (not name.startswith("_") and isinstance(obj, type) and obj not in vars(builtins).values()
                    and obj.__module__.split(".")[0] in ("oxoffice", "builtins")):
                classes.setdefault(obj, set()).add(name)
    return {f"{module.__name__}.{obj.__name__}": names for obj, names in classes.items()}


for module, skill in ((oxoffice.docx, "docx"), (oxoffice.pptx, "pptx")):
    words = set(re.findall(r"[A-Za-z_]\w*", skill_code(skill)))
    missing = sorted(k for k, names in python_surface(module).items() if not names & words and k not in NOT_DOCUMENTED)
    print(f"{skill}: public classes not cited: {missing}")
    code = skill_code(skill)
    surface = set()

    def walk(command):
        surface.add(command)
        for sub in sorted(subcommands(cli_help(*command) or "")):
            walk(command + (sub,))
    walk((skill,))
    uncited = [" ".join(c) for c in sorted(surface) if len(c) > 1
               and not re.search(r"oxoffice " + r"\s+".join(map(re.escape, c)) + r"(?![\w-])", code)]
    print(f"{skill}: CLI commands not cited: {uncited}")
