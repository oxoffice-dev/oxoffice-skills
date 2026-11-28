"""Write the input files of the skill examples into a folder: report.docx, deck.pptx and budget.xlsx (synthetic
documents, see fixtures/), Aileron-Regular.ttf (the CC0 font Pillow bundles, for the font examples) and photo.png (a drawn
picture). Every file is the same from run to run. Needs Pillow, python-pptx and XlsxWriter.

    python tests/make_fixtures.py OUT_DIR
"""

import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent


def main() -> None:
    out = Path(sys.argv[1])
    out.mkdir(parents=True, exist_ok=True)
    for script, name in (("make_fixture_docx.py", "report.docx"), ("make_fixture_pptx.py", "deck.pptx"),
                         ("make_fixture_xlsx.py", "budget.xlsx")):
        subprocess.run([sys.executable, str(HERE / "fixtures" / script), str(out / name)], check=True)
    (out / "Aileron-Regular.ttf").write_bytes(ImageFont.load_default(12).font_bytes)
    photo = Image.new("RGB", (400, 300), (64, 112, 150))
    draw = ImageDraw.Draw(photo)
    draw.rectangle((40, 180, 360, 300), fill=(96, 120, 80))
    draw.ellipse((260, 30, 330, 100), fill=(236, 200, 90))
    photo.save(out / "photo.png", optimize=False)
    print(f"{out}: report.docx, deck.pptx, budget.xlsx, Aileron-Regular.ttf, photo.png")


if __name__ == "__main__":
    main()
