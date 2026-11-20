"""Every `.attribute` cited in a Python code span of the skills (outside fences) appears as a word in the installed
oxoffice sources (stubs, Python modules) or the helper scripts; the rest is listed for review."""
import re, sys
from pathlib import Path
import oxoffice
SKILLS = Path(__file__).resolve().parents[2] / "plugin" / "skills"
pkg = Path(oxoffice.__file__).parent
words = set()
for f in [*pkg.rglob("*.py"), *pkg.rglob("*.pyi"), *SKILLS.glob("*/scripts/*.py")]:
    words |= set(re.findall(r"\w+", f.read_text()))
# python-docx / python-pptx / stdlib names the skills cite on purpose as foreign or absent
missing = {}
count = 0
for path in sorted(SKILLS.glob("*/SKILL.md")) + sorted(SKILLS.glob("*/references/*.md")):
    fenced = False
    for n, line in enumerate(path.read_text().splitlines(), 1):
        if line.startswith("```"):
            fenced = not fenced
            continue
        if fenced:
            continue
        for span in re.findall(r"`([^`]+)`", line):
            if re.match(r"(oxoffice (docx|pptx)|python3?|pip|bash|soffice|\$)", span) or re.search(r"\.(md|py|json|docx|pptx|xml|png|pdf|ttf|pyi|dotx|potx)\b", span) and " " not in span:
                continue
            for attr in re.findall(r"(?<=[\w)\]])\.([A-Za-z_]\w*)", span):
                count += 1
                if attr not in words:
                    missing.setdefault(attr, []).append(f"{path.relative_to(SKILLS)}:{n}: `{span[:90]}`")
print("attributes checked:", count)
for attr, where in sorted(missing.items()):
    print(attr, "|", where[0], f"(+{len(where) - 1})" if len(where) > 1 else "")
