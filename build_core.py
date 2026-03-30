import re

with open('app.py', 'r') as f:
    content = f.read()

# Split the content at "# --- HELPER FUNCTIONS ---"
if "# --- HELPER FUNCTIONS ---" in content:
    _, helpers = content.split("# --- HELPER FUNCTIONS ---", 1)
else:
    helpers = content # Fallback

# Remove the ocr_extract_tables definition and any flask code
# Luckily ocr_extract_tables is before HELPER FUNCTIONS so it got removed anyway
# Wait, NO. ocr_extract_tables was AT line 1681 and HELPER FUNCTIONS AT line 520.
# So HELPER FUNCTIONS is first, then the other functions.
# Let's check where apply_styles ends. It ends at the end of the file except `if __name__ == '__main__':`

# We need everything from HELPER FUNCTIONS down to `if __name__ == '__main__':`
if "if __name__ == '__main__':" in helpers:
    helpers, _ = helpers.split("if __name__ == '__main__':", 1)

# Now, remove the `ocr_extract_tables` body to avoid dependency issues.
# We will use regex to find: `def ocr_extract_tables(doc):` and replace its body with `pass`
import ast

def remove_function(source, func_name):
    try:
        module = ast.parse(source)
    except SyntaxError:
        return source # Fallback
        
    class FunctionRemover(ast.NodeTransformer):
        def visit_FunctionDef(self, node):
            if node.name == func_name:
                return None # Remove the node
            return node
    
    new_module = FunctionRemover().visit(module)
    return ast.unparse(new_module)

helpers_clean = remove_function(helpers, 'ocr_extract_tables')

# Remove the call to ocr_extract_tables inside apply_styles
helpers_clean = re.sub(r'ocr_extract_tables\(doc\)', 'pass', helpers_clean)

# Also remove docx2pdf
helpers_clean = re.sub(r'from docx2pdf import convert.*', '', helpers_clean)

# Imports needed:
imports = """
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import json
import re\n
"""

with open('templater_core.py', 'w') as f:
    f.write(imports + helpers_clean)

print("Generated templater_core.py")
