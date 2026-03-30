import re

with open("static/js/main.js", "r") as f:
    js = f.read()

# Replace export_format pdf stuff
js = js.replace("const ext = (exportFormat === 'pdf') ? '.pdf' : '.docx';", "const ext = '.docx';")
js = js.replace("const fmt = exportFormat === 'both' ? 'DOCX + PDF' : exportFormat.toUpperCase();", "const fmt = 'DOCX';")

# Find the start of processBtn listener and replace it entirely up to the closing brace
pattern = re.compile(r"if \(processBtn\) \{.*?\}\);\n    \}", re.DOTALL)

pyodide_logic = """if (processBtn) {
        processBtn.addEventListener('click', async () => {
            if (selectedFiles.size === 0) return;
            processBtn.disabled = true;
            processBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> PROCESANDO...';
            log('Iniciando entorno WebAssembly (Primera vez puede demorar)...', 'info');

            if (!window.pyodide) {
                try {
                    window.pyodide = await loadPyodide();
                    await window.pyodide.loadPackage("micropip");
                    const micropip = window.pyodide.pyimport("micropip");
                    await micropip.install("python-docx");
                    await micropip.install("lxml");
                    
                    // Fetch and evaluate templater_core.py
                    const core_resp = await fetch('templater_core.py');
                    const core_code = await core_resp.text();
                    window.pyodide.runPython(core_code);
                } catch(e) {
                    log('Error iniciando Pyodide: ' + e, 'error');
                    processBtn.disabled = false;
                    processBtn.innerHTML = '<span>INICIAR PROCESO</span><i class="fa-solid fa-play"></i>';
                    return;
                }
            }
            
            log('Procesando cola...', 'info');

            const pyodide = window.pyodide;
            const prefix = filePrefixInput ? filePrefixInput.value : '';
            const suffix = fileSuffixInput ? fileSuffixInput.value : '';
            const forceStyles = document.getElementById('forceStyles');
            
            let cfg = getAdvancedConfig();
            cfg._paper_size = paperSize;
            cfg._force_styles = forceStyles ? forceStyles.checked : true;
            
            // Helper to write input files to Pyodide FS
            async function writeInputFile(inputId, filename) {
                const input = document.getElementById(inputId);
                if (input && input.files && input.files.length > 0) {
                    const arrayBuffer = await input.files[0].arrayBuffer();
                    pyodide.FS.writeFile('/tmp/' + filename, new Uint8Array(arrayBuffer));
                }
            }
            // Create /tmp folder if it doesn't exist
            try { pyodide.FS.mkdir('/tmp'); } catch(e) {}
            
            // Write images to /tmp
            await writeInputFile('coverInput', 'custom_cover.png');
            await writeInputFile('headerInput', 'custom_header.png');
            await writeInputFile('footerInput', 'custom_footer.png');
            await writeInputFile('backpageInput', 'custom_backpage.png');
            
            let success_count = 0;
            let error_count = 0;
            
            for (let [name, v] of selectedFiles.entries()) {
                v.status = 'processing';
                selectedFiles.set(name, v);
                renderQueue();
                log('Transformando: ' + name + '...', 'info');
                
                try {
                    // Write file to Pyodide FS
                    let arrayBuffer = await v.file.arrayBuffer();
                    pyodide.FS.writeFile('/tmp/input.docx', new Uint8Array(arrayBuffer));
                    // Write config
                    pyodide.FS.writeFile('/tmp/config.json', JSON.stringify(cfg));
                    
                    // Execute python wrapper
                    pyodide.runPython(`
import json
from docx import Document

with open('/tmp/config.json', 'r') as f:
    config = json.load(f)

# Override upload path for Pyodide
import types
app = types.SimpleNamespace()
app.config = {'UPLOAD_FOLDER': '/tmp'}
globals()['app'] = app

doc = Document('/tmp/input.docx')
apply_styles(doc, config, config.get('_paper_size', 'letter'))
doc.save('/tmp/output.docx')
                    `);
                    
                    // Read output
                    let outBuf = pyodide.FS.readFile('/tmp/output.docx');
                    let blob = new Blob([outBuf], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                    
                    let outName = prefix + name.replace('.docx', '') + suffix + '.docx';
                    
                    // Trigger download
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = outName;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                    
                    v.status = 'done';
                    v.resultDocx = outName;
                    success_count++;
                    log('Completado: ' + name, 'success');
                } catch(e) {
                    v.status = 'error';
                    error_count++;
                    log('[ERROR] ' + name + ' -> ' + e.message, 'error');
                }
                
                selectedFiles.set(name, v);
                renderQueue();
            }
            
            log('✅ Cola finalizada — ' + success_count + ' archivo(s) procesados correctamente.', 'success');
            if (error_count > 0) log('⚠️ ' + error_count + ' archivo(s) fallaron.', 'error');
            
            processBtn.disabled = false;
            processBtn.innerHTML = '<span>INICIAR PROCESO</span><i class="fa-solid fa-play"></i>';
        });
    }"""

js_new = pattern.sub(pyodide_logic, js)

with open("static/js/main_web.js", "w") as f:
    f.write(js_new)

print("Generated static/js/main_web.js")
