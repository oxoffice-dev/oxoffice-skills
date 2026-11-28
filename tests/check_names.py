"""Every name the skills cite exists in the installed oxoffice: Python functions and their keyword arguments,
attributes, CLI commands and flags, in SKILL.md and the references. Exit 1 on the first list of problems. Then, for
information, the public classes and CLI commands that no skill cites.

    python tests/check_names.py SKILLS_DIR
"""
import ast
import builtins
import functools
import importlib
import pkgutil
import re
import shutil
import subprocess
import sys
from pathlib import Path

import oxoffice
import oxoffice.docx
import oxoffice.pptx
import oxoffice.xlsx

SKILLS = Path(sys.argv[1]).resolve()
OX = shutil.which("oxoffice")
DOCS = sorted(SKILLS.glob("*/SKILL.md")) + sorted(SKILLS.glob("*/references/*.md"))
# python-pptx, python-docx, lxml and standard library names the skills cite on purpose
FOREIGN = {"Presentation", "Document", "fromstring", "tostring", "xpath", "related_part", "ZipFile", "unpack"}
# worksheet functions the xlsx skill writes in formulas
FOREIGN |= {"IF", "SUM", "SORT", "XLOOKUP", "VLOOKUP", "DATE", "LET", "LAMBDA"}
stats = {"spans": 0, "calls": 0, "keywords": 0, "cli_spans": 0, "flags": 0}


def rel(path):
    return path.relative_to(SKILLS)


def spans(path):
    """(line, code span) for inline spans outside fences, plus each `oxoffice ...` line of a fenced bash block."""
    fenced = False
    lines = path.read_text().splitlines()
    for n, line in enumerate(lines, 1):
        if line.lstrip().startswith("```"):
            fenced = not fenced
            continue
        if fenced:
            k = n
            while line.rstrip().endswith("\\") and k < len(lines):   # a command continued on the next line
                line = line.rstrip()[:-1] + " " + lines[k]
                k += 1
            s = line.split(" #")[0].strip()
            for prefix in ("ox ", "oxoffice "):
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
                # a class is called with the arguments of its __new__ or __init__
                inits = [f.args for f in node.body if isinstance(f, (ast.FunctionDef, ast.AsyncFunctionDef))
                         and f.name in ("__new__", "__init__")]
                for a in inits:
                    names = {x.arg for x in a.posonlyargs + a.args + a.kwonlyargs} - {"self", "cls"}
                    api.setdefault(node.name, []).append(names | ({"**"} if a.kwarg else set()))
                if not inits:
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
        if fenced or re.match(r"(oxoffice|ox|oxpy|python3?|pip|bash|soffice|uvx?|pixi|conda|curl|irm|sha256sum|gh) |(ox|oxpy)\(\)|\$|[A-Z]+=", span):
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
    res = subprocess.run([OX, *command, "--help"], capture_output=True, text=True)
    return res.stdout if res.returncode == 0 else None


def subcommands(help_text):
    section = help_text.split("Commands:", 1)[1].split("\n\n", 1)[0] if "Commands:" in help_text else ""
    return set(re.findall(r"^  (\S+)", section, re.M)) - {"help"}


def cli_problems(path):
    for n, span, _ in spans(path):
        tokens = span.replace("[", " ").replace("]", " ").split()
        if len(tokens) < 2 or tokens[0] not in ("oxoffice", "ox") or tokens[1] not in ("docx", "pptx", "xlsx"):
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
        # a command that has subcommands takes one: a lowercase word there that is not one is a wrong command
        known = set.union(*(subcommands(cli_help(*c) or "") for c in commands))
        if known and i < len(tokens) and re.fullmatch(r"[a-z][a-z-]*(/[a-z][a-z-]*)*", tokens[i]):
            wrong = [w for w in tokens[i].split("/") if w not in known]
            yield f"{rel(path)}:{n}: no command oxoffice {' '.join(commands[0])} {'/'.join(wrong)}: `{span}`"
            continue
        helps = [cli_help(*c) or "" for c in commands]
        for token in tokens[i:]:
            for flag in token.split("/"):
                flag = flag.split("=")[0]
                if flag in ("||", "-eq"):
                    break
                if re.fullmatch(r"--?[A-Za-z][\w-]*", flag):
                    stats["flags"] += 1
                    if not any(re.search(rf"(?<![\w-]){re.escape(flag)}(?![\w-])", h) for h in helps):
                        yield f"{rel(path)}:{n}: oxoffice {' '.join(commands[0])} has no {flag}: `{span}`"


@functools.lru_cache(maxsize=None)
def source_words():
    pkg = Path(oxoffice.__file__).parent
    words = set()
    for f in [*pkg.rglob("*.py"), *pkg.rglob("*.pyi"), *SKILLS.glob("*/scripts/*.py")]:
        words |= set(re.findall(r"\w+", f.read_text()))
    return words


def attribute_problems(path):
    """Every `.attribute` of a Python code span is a word of the installed sources or of the skills' scripts."""
    for n, span, fenced in spans(path):
        if fenced or "://" in span or re.match(r"(oxoffice|ox|oxpy|python3?|pip|bash|uvx?|pixi|conda|curl|irm) |(ox|oxpy)\(\)|\$|[A-Z]+=", span):
            continue
        code = re.sub(r"\"[^\"]*\"|'[^']*'", "''", span)
        for attr in re.findall(r"(?<=[\w)\]])\.([A-Za-z_]\w*)(?!\w*\*)", code):
            # file extensions, python-pptx's private attributes, placeholders such as MSO_SHAPE.X
            if (attr in ("md", "py", "json", "docx", "pptx", "xlsx", "xlsm", "xltx", "xltm", "csv", "tsv", "dotx", "potx", "xml", "png", "pdf", "ttf", "pyi")
                    or attr.startswith("_") or re.fullmatch(r"[A-Z]", attr) or attr in FOREIGN):
                continue
            stats["attributes"] += 1
            if attr not in source_words():
                yield f"{rel(path)}:{n}: no attribute {attr}: `{span}`"


stats["attributes"] = 0
problems = []
for path in DOCS:
    problems += list(python_problems(path))
    problems += list(cli_problems(path))
    problems += list(attribute_problems(path))
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


for module, skill in ((oxoffice.docx, "docx"), (oxoffice.pptx, "pptx"), (oxoffice.xlsx, "xlsx")):
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

sys.exit(1 if problems else 0)
