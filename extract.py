with open('app.py', 'r') as f:
    lines = f.readlines()

out = [
    "from docx import Document\n",
    "from docx.shared import Pt, Inches, RGBColor, Cm, Emu\n",
    "from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK\n",
    "from docx.enum.table import WD_TABLE_ALIGNMENT\n",
    "from docx.oxml.ns import qn\n",
    "from docx.oxml import OxmlElement\n",
    "import io\n",
    "import json\n",
    "\n"
]

start_idx = 0
for i, line in enumerate(lines):
    if "# --- HELPER FUNCTIONS ---" in line:
        start_idx = i
        break

for line in lines[start_idx:]:
    if "def ocr_extract_tables" in line:
        break # We might need to handle this manually, or just use regex to remove it. Let's just output it and then we can strip it out if it fails or just comment the body.
    out.append(line)

with open('templater_core.py', 'w') as f:
    f.writelines(out)

