import os
import re

with open('templates/tool_docs.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Inline sidebar and config modal
with open('templates/components/sidebar.html', 'r', encoding='utf-8') as f:
    sidebar = f.read()

with open('templates/components/config_modal.html', 'r', encoding='utf-8') as f:
    config_modal = f.read()

html = html.replace("{% include 'components/sidebar.html' %}", sidebar)
html = html.replace("{% include 'components/config_modal.html' %}", config_modal)

# Inline CSS
try:
    with open('static/css/fonts.css', 'r', encoding='utf-8') as f:
        # keep standard relative font paths if possible, but actually let's just leave the link for fonts.
        # It's better to keep links for things in static/ if we upload to GH pages.
        pass
except:
    pass

# We will just fix links to point to GitHub Pages relative paths
html = html.replace('href="/static/', 'href="static/')
html = html.replace('src="/static/', 'src="static/')

# Remove PDF format button since it's unsupported in Browser Pyodide
html = html.replace('''<button class="format-btn" data-format="pdf" title="Solo PDF">
                            <i class="fa-solid fa-file-pdf"></i>
                            <span>PDF</span>
                        </button>''', '')
html = html.replace('''<button class="format-btn" data-format="both" title="Word + PDF">
                            <i class="fa-solid fa-copy"></i>
                            <span>Ambos</span>
                        </button>''', '')

# Remove Plantillas block completely to simplify
html = re.sub(r'<!-- Templates -->.*?</div>.*?</div>', '', html, flags=re.DOTALL)

# Insert Pyodide in head
pyodide_script = '<script src="https://cdn.jsdelivr.net/pyodide/v0.25.0/full/pyodide.js"></script>'
html = html.replace('</head>', f'    {pyodide_script}\n</head>')

# Inline JS
with open('static/js/main_web.js', 'r', encoding='utf-8') as f:
    main_js = f.read()

# Replace script tag with inline script
html = re.sub(r'<script src="static/js/main.js"></script>', f'<script>\n{main_js}\n</script>', html)

# Replace the EDD template code in inline JS (it will error if buttons are missing, but let's be safe)
html = re.sub(r'if \(exportTemplateBtn\).*?}\);.*?}', '', html, flags=re.DOTALL)
html = re.sub(r'if \(importTemplateBtn\).*?}', '', html, flags=re.DOTALL)
html = re.sub(r'if \(templateInput\).*?}\);.*?}', '', html, flags=re.DOTALL)

with open('DOC-WEBAPP.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("Generated DOC-WEBAPP.html with inlined JS code.")
