document.addEventListener('DOMContentLoaded', () => {
    // === ELEMENTS ===
    const coverInput = document.getElementById('coverInput');
    const coverUploadBtn = document.getElementById('coverUploadBtn');
    const coverThumb = document.getElementById('coverThumb');
    const coverRow = document.getElementById('coverRow');

    const headerInput = document.getElementById('headerInput');
    const headerUploadBtn = document.getElementById('headerUploadBtn');
    const headerThumb = document.getElementById('headerThumb');
    const headerRow = document.getElementById('headerRow');

    const footerInput = document.getElementById('footerInput');
    const footerUploadBtn = document.getElementById('footerUploadBtn');
    const footerThumb = document.getElementById('footerThumb');
    const footerRow = document.getElementById('footerRow');

    const backpageInput = document.getElementById('backpageInput');
    const backpageUploadBtn = document.getElementById('backpageUploadBtn');
    const backpageThumb = document.getElementById('backpageThumb');
    const backpageRow = document.getElementById('backpageRow');

    const mainDropZone = document.getElementById('mainDropZone');
    const docInput = document.getElementById('docInput');
    const processBtn = document.getElementById('processBtn');

    // Production Queue
    const productionQueue = document.getElementById('productionQueue');
    const queueList = document.getElementById('queueList');
    const queueCount = document.getElementById('queueCount');
    const clearQueueBtn = document.getElementById('clearQueueBtn');
    const addMoreBtn = document.getElementById('addMoreBtn');

    // Console (Static Panel)
    const consoleLogs = document.getElementById('consoleLogs');

    // === STATE ===
    let selectedFiles = new Map();
    let paperSize = 'letter';
    let exportFormat = 'docx';

    // === TESSERACT DETECTION ===
    function checkTesseract() {
        const statusEl = document.getElementById('tesseractStatus');
        const dotEl = document.getElementById('tessDot');
        const labelEl = document.getElementById('tessLabel');
        if (!statusEl) return;
        statusEl.style.display = 'flex';
        labelEl.textContent = 'Verificando...';
        dotEl.className = 'tess-dot';
        fetch('/check_tesseract')
            .then(r => r.json())
            .then(data => {
                if (data.available && data.has_ocr_deps) {
                    dotEl.className = 'tess-dot ok';
                    const ver = data.version || '';
                    labelEl.textContent = ver.charAt(0).toUpperCase() + ver.slice(1);
                } else if (data.available && !data.has_ocr_deps) {
                    dotEl.className = 'tess-dot fail';
                    labelEl.textContent = 'Falta: pip install img2table';
                } else {
                    dotEl.className = 'tess-dot fail';
                    labelEl.textContent = 'No detectado — brew install tesseract';
                }
            })
            .catch(() => {
                dotEl.className = 'tess-dot fail';
                labelEl.textContent = 'Error de conexión';
            });
    }

    // Show/hide indicator based on OCR toggle
    const ocrToggle = document.getElementById('cfgOcrTables');
    if (ocrToggle) {
        ocrToggle.addEventListener('change', () => {
            if (ocrToggle.checked) checkTesseract();
            else {
                const st = document.getElementById('tesseractStatus');
                if (st) st.style.display = 'none';
            }
        });
    }

    // === DEBUG CONSOLE ===
    const debugToggle = document.getElementById('debugConsoleToggle');
    const debugBody = document.getElementById('debugConsoleBody');
    const debugLog = document.getElementById('debugLog');
    const debugChevron = document.getElementById('debugChevron');
    const refreshDebugBtn = document.getElementById('refreshDebugBtn');

    function loadDebugInfo() {
        if (!debugLog) return;
        debugLog.textContent = '> Cargando diagnóstico...\n';
        fetch('/debug_info')
            .then(r => r.json())
            .then(info => {
                let out = '=== EDTECH DOC TEMPLATER DEBUG ===\n';
                out += `> Frozen: ${info.frozen}\n`;
                out += `> Platform: ${info.platform}\n`;
                out += `> Arch: ${info.arch}\n`;
                out += `> Python: ${info.python_version}\n`;
                out += `> Bundle Dir: ${info.bundle_dir}\n`;
                out += `> User Data: ${info.user_data}\n`;
                out += `> Upload: ${info.upload_folder}\n`;
                out += `> Output: ${info.output_folder}\n`;
                out += `> Static: ${info.static_folder}\n`;
                out += `> Templates: ${info.template_folder}\n`;
                out += `> HAS_OCR: ${info.has_ocr}\n`;
                out += '--- Tesseract ---\n';
                const t = info.tesseract || {};
                out += `> Available: ${t.available}\n`;
                out += `> Version: ${t.version || 'N/A'}\n`;
                if (t.bundled_path) out += `> Bundled Path: ${t.bundled_path}\n`;
                if (t.bundled_exists !== undefined) out += `> Bundled Exists: ${t.bundled_exists}\n`;
                if (t.bundled_files) out += `> Bundled Dylibs: ${t.bundled_files.join(', ')}\n`;
                if (t.cmd) out += `> Used Cmd: ${t.cmd}\n`;
                if (t.returncode !== undefined) out += `> Return Code: ${t.returncode}\n`;
                if (t.stderr) out += `> Stderr: ${t.stderr}\n`;
                if (t.error) out += `> Error: ${t.error}\n`;
                if (info.dyld_library_path) out += `> DYLD_LIBRARY_PATH: ${info.dyld_library_path}\n`;
                if (info.dyld_fallback) out += `> DYLD_FALLBACK: ${info.dyld_fallback}\n`;
                out += '=================================\n';
                debugLog.textContent = out;
                debugLog.scrollTop = debugLog.scrollHeight;
            })
            .catch(err => {
                debugLog.textContent = `> Error: ${err.message}\n`;
            });
    }

    if (debugToggle && debugBody) {
        debugToggle.addEventListener('click', () => {
            const isHidden = debugBody.style.display === 'none';
            debugBody.style.display = isHidden ? 'block' : 'none';
            if (debugChevron) debugChevron.style.transform = isHidden ? 'rotate(180deg)' : '';
            if (isHidden && debugLog && !debugLog.textContent.trim()) loadDebugInfo();
        });
    }
    if (refreshDebugBtn) refreshDebugBtn.addEventListener('click', loadDebugInfo);
    const copyDebugBtn = document.getElementById('copyDebugBtn');
    if (copyDebugBtn) copyDebugBtn.addEventListener('click', () => {
        if (debugLog && debugLog.textContent) {
            navigator.clipboard.writeText(debugLog.textContent).then(() => {
                copyDebugBtn.innerHTML = '<i class="fa-solid fa-check"></i> Copiado';
                setTimeout(() => { copyDebugBtn.innerHTML = '<i class="fa-solid fa-copy"></i> Copiar'; }, 2000);
            }).catch(() => {
                // Fallback for pywebview where clipboard API might not work
                const ta = document.createElement('textarea');
                ta.value = debugLog.textContent;
                document.body.appendChild(ta);
                ta.select();
                document.execCommand('copy');
                document.body.removeChild(ta);
                copyDebugBtn.innerHTML = '<i class="fa-solid fa-check"></i> Copiado';
                setTimeout(() => { copyDebugBtn.innerHTML = '<i class="fa-solid fa-copy"></i> Copiar'; }, 2000);
            });
        }
    });

    // Show/hide Índice tab based on TOC toggle
    const tocToggle = document.getElementById('cfgToc');
    const tabTocBtn = document.getElementById('tabTocBtn');
    if (tocToggle && tabTocBtn) {
        tocToggle.addEventListener('change', () => {
            tabTocBtn.disabled = !tocToggle.checked;
            if (!tocToggle.checked) {
                // If the TOC tab is currently active, switch to General
                if (tabTocBtn.classList.contains('active')) {
                    tabTocBtn.classList.remove('active');
                    document.getElementById('tab-toc').classList.add('hidden-tab');
                    document.querySelector('[data-tab="tab-general"]').classList.add('active');
                    document.getElementById('tab-general').classList.remove('hidden-tab');
                }
            }
        });
    }

    // === PAGE IDENTITY ===
    const brandSub = document.querySelector('.brand-sub');
    if (brandSub) {
        brandSub.textContent = 'DOC templater';
    }

    // Config Button Logic
    const sidebarConfigBtn = document.getElementById('openConfigBtnSidebar');
    if (sidebarConfigBtn) {
        sidebarConfigBtn.addEventListener('click', () => {
            if (configModal) configModal.classList.add('active');
        });
    }

    // === FOLDER PICKER ===
    const selectFolderBtn = document.getElementById('selectFolderBtn');
    const outputFolderDisplay = document.getElementById('outputFolderDisplay');
    const outputFolderValue = document.getElementById('outputFolderValue');
    if (selectFolderBtn) {
        selectFolderBtn.addEventListener('click', async () => {
            // Try pywebview native dialog first
            if (window.pywebview && window.pywebview.api && window.pywebview.api.select_folder) {
                try {
                    const folder = await window.pywebview.api.select_folder();
                    if (folder) {
                        outputFolderDisplay.value = folder;
                        outputFolderValue.value = folder;
                        log(`Carpeta seleccionada: ${folder}`, 'success');
                    }
                } catch (e) {
                    log('Error al seleccionar carpeta: ' + e.message, 'error');
                }
            } else {
                // Dev mode fallback: prompt
                const folder = prompt('Ingresa la ruta completa de la carpeta de salida:', outputFolderValue.value || '');
                if (folder && folder.trim()) {
                    outputFolderDisplay.value = folder.trim();
                    outputFolderValue.value = folder.trim();
                    log(`Carpeta de salida: ${folder.trim()}`, 'info');
                }
            }
        });
    }

    // === PREFIX / SUFFIX & FILENAME PREVIEW ===
    const filePrefixInput = document.getElementById('filePrefix');
    const fileSuffixInput = document.getElementById('fileSuffix');
    const namePreviewText = document.getElementById('namePreviewText');

    function updateNamePreview() {
        const prefix = filePrefixInput ? filePrefixInput.value : '';
        const suffix = fileSuffixInput ? fileSuffixInput.value : '';
        const ext = '.docx';
        const bothSuffix = (exportFormat === 'both') ? ' (.docx + .pdf)' : '';
        const preview = `${prefix}documento${suffix}${ext}${bothSuffix}`;
        if (namePreviewText) namePreviewText.textContent = preview;
    }

    if (filePrefixInput) filePrefixInput.addEventListener('input', updateNamePreview);
    if (fileSuffixInput) fileSuffixInput.addEventListener('input', updateNamePreview);
    updateNamePreview();

    // === FORMAT SELECTOR ===
    document.querySelectorAll('.format-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.format-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            exportFormat = btn.getAttribute('data-format');
            updateNamePreview();
        });
    });

    // === IMAGE UPLOAD BUTTON HANDLERS ===
    // Helper: Use pywebview native dialog if available, else fallback to HTML input
    function triggerImageUpload(slot, inputEl, onSuccess) {
        if (window.pywebview && window.pywebview.api && window.pywebview.api.select_image) {
            // Native mode: use pywebview file dialog
            window.pywebview.api.select_image(slot).then(result => {
                if (result && result.success) {
                    onSuccess(result);
                }
            }).catch(() => {
                console.log('Native file dialog cancelled or errored for ' + slot);
            });
        } else {
            // Dev mode: use standard HTML file input
            inputEl.click();
        }
    }

    if (coverUploadBtn) {
        coverUploadBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            triggerImageUpload('cover', coverInput, (result) => {
                coverThumb.style.backgroundImage = `url(${result.path}?t=${Date.now()})`;
                coverThumb.classList.add('has-image');
                coverRow.classList.add('uploaded');
                updateCoverPreview();
                log(`Portada cargada: ${result.path}`, 'success');
            });
        });
        coverInput.addEventListener('change', (e) => handleCoverUpload(e.target.files[0]));
    }
    if (headerUploadBtn) {
        headerUploadBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            triggerImageUpload('header', headerInput, (result) => {
                const ext = (result.filename || '').split('.').pop().toLowerCase();
                if (['emf', 'wmf'].includes(ext)) {
                    headerThumb.classList.add('has-image');
                    headerThumb.style.background = 'var(--primary-light)';
                    headerThumb.innerHTML = '<i class="fa-solid fa-vector-square" style="font-size:10px;color:var(--primary);line-height:28px;text-align:center;width:100%;display:block"></i>';
                } else {
                    headerThumb.style.backgroundImage = `url(${result.path}?t=${Date.now()})`;
                    headerThumb.classList.add('has-image');
                }
                headerRow.classList.add('uploaded');
                log(`Encabezado cargado: ${result.path}`, 'success');
            });
        });
        headerInput.addEventListener('change', (e) => handleHeaderUpload(e.target.files[0]));
    }
    if (footerUploadBtn) {
        footerUploadBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            triggerImageUpload('footer', footerInput, (result) => {
                const ext = (result.filename || '').split('.').pop().toLowerCase();
                if (['emf', 'wmf'].includes(ext)) {
                    footerThumb.classList.add('has-image');
                    footerThumb.style.background = 'var(--primary-light)';
                    footerThumb.innerHTML = '<i class="fa-solid fa-vector-square" style="font-size:10px;color:var(--primary);line-height:28px;text-align:center;width:100%;display:block"></i>';
                } else {
                    footerThumb.style.backgroundImage = `url(${result.path}?t=${Date.now()})`;
                    footerThumb.classList.add('has-image');
                }
                footerRow.classList.add('uploaded');
                log(`Pie de página cargado: ${result.path}`, 'success');
            });
        });
        footerInput.addEventListener('change', (e) => handleFooterUpload(e.target.files[0]));
    }
    if (backpageUploadBtn) {
        backpageUploadBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            triggerImageUpload('backpage', backpageInput, (result) => {
                backpageThumb.style.backgroundImage = `url(${result.path}?t=${Date.now()})`;
                backpageThumb.classList.add('has-image');
                backpageRow.classList.add('uploaded');
                log(`Contraportada cargada: ${result.path}`, 'success');
            });
        });
        backpageInput.addEventListener('change', (e) => handleBackpageUpload(e.target.files[0]));
    }

    function handleCoverUpload(file) {
        if (!file || !file.type.startsWith('image/')) {
            log('Solo imágenes (PNG, JPG) para la portada.', 'error');
            return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
            coverThumb.style.backgroundImage = `url(${e.target.result})`;
            coverThumb.classList.add('has-image');
            coverRow.classList.add('uploaded');
        };
        reader.readAsDataURL(file);

        // Fetch removed for Pyodide (files read locally at process)
        Promise.resolve({path: 'local'})
            .then(d => {
                log(`Portada cargada localmente`, 'success');
                updateCoverPreview();
            })
            .catch(() => log('Error', 'error'));
    }

    function handleBackpageUpload(file) {
        if (!file || !file.type.startsWith('image/')) {
            log('Formato no soportado. Use PNG o JPG.', 'error');
            return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
            backpageThumb.style.backgroundImage = `url(${e.target.result})`;
            backpageThumb.classList.add('has-image');
            backpageRow.classList.add('uploaded');
        };
        reader.readAsDataURL(file);

        // Fetch removed
        Promise.resolve()
            .then(() => log(`Contraportada cargada localmente`, 'success'));
    }

    function handleHeaderUpload(file) {
        if (!file) return;
        const isImage = file.type.startsWith('image/');
        const isVector = /\.(emf|wmf)$/i.test(file.name);

        if (!isImage && !isVector) {
            log('Formato no soportado. Use PNG, JPG, EMF o WMF.', 'error');
            return;
        }

        if (isImage) {
            const reader = new FileReader();
            reader.onload = (e) => {
                headerThumb.style.backgroundImage = `url(${e.target.result})`;
                headerThumb.classList.add('has-image');
            };
            reader.readAsDataURL(file);
        } else {
            headerThumb.classList.add('has-image');
            headerThumb.style.background = 'var(--primary-light)';
            headerThumb.innerHTML = '<i class="fa-solid fa-vector-square" style="font-size:10px;color:var(--primary);line-height:28px;text-align:center;width:100%;display:block"></i>';
        }
        headerRow.classList.add('uploaded');

        // Fetch removed
        Promise.resolve()
            .then(() => log(`Encabezado cargado localmente`, 'success'));
    }

    function handleFooterUpload(file) {
        if (!file) return;
        const isImage = file.type.startsWith('image/');
        const isVector = /\.(emf|wmf)$/i.test(file.name);

        if (!isImage && !isVector) {
            log('Formato no soportado. Use PNG, JPG, EMF o WMF.', 'error');
            return;
        }

        if (isImage) {
            const reader = new FileReader();
            reader.onload = (e) => {
                footerThumb.style.backgroundImage = `url(${e.target.result})`;
                footerThumb.classList.add('has-image');
            };
            reader.readAsDataURL(file);
        } else {
            footerThumb.classList.add('has-image');
            footerThumb.style.background = 'var(--primary-light)';
            footerThumb.innerHTML = '<i class="fa-solid fa-vector-square" style="font-size:10px;color:var(--primary);line-height:28px;text-align:center;width:100%;display:block"></i>';
        }
        footerRow.classList.add('uploaded');

        // Fetch removed
        Promise.resolve()
            .then(() => log(`Pie de página cargado localmente`, 'success'));
    }

    // === DOCUMENTS LOGIC ===


    if (mainDropZone) {
        mainDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
        });
        mainDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            handleDocs(e.dataTransfer.files);
        });
    }

    if (docInput) {
        docInput.addEventListener('change', (e) => handleDocs(e.target.files));
    }

    function handleDocs(files) {
        if (!files || files.length === 0) return;

        const allowedExt = '.docx';

        let addedCount = 0;
        Array.from(files).forEach(f => {
            if (f.name.toLowerCase().endsWith(allowedExt) && !selectedFiles.has(f.name)) {
                selectedFiles.set(f.name, {
                    file: f,
                    status: 'pend',
                    resultDocx: null
                });
                addedCount++;
            } else if (!f.name.toLowerCase().endsWith(allowedExt)) {
                log(`Ignorado: ${f.name} (Requiere ${allowedExt})`, 'error');
            }
        });

        if (addedCount > 0) {
            renderQueue();
            log(`${addedCount} archivo(s) añadidos.`, 'success');
        } else if (selectedFiles.size === 0) {
            log('No se añadieron archivos válidos.', 'error');
        } else {
            renderQueue();
        }
    }

    function renderQueue() {
        // Toggle Visibility
        if (selectedFiles.size > 0) {
            if (mainDropZone) mainDropZone.style.display = 'none';
            productionQueue.classList.add('visible');
            productionQueue.style.display = 'flex';
            processBtn.disabled = false;
        } else {
            productionQueue.classList.remove('visible');
            productionQueue.style.display = 'none';
            if (mainDropZone) mainDropZone.style.display = 'flex';
            processBtn.disabled = true;
        }

        queueCount.innerText = selectedFiles.size;
        queueList.innerHTML = '';

        selectedFiles.forEach((data, name) => {
            // Determine Info display
            let infoHtml = '';
            let iconClass = 'fa-file-word';
            let iconColor = '';

            // File
            const sizeMB = (data.file.size / 1024 / 1024).toFixed(2);
            infoHtml = `
                <div class="q-info" style="display:flex; align-items:center; gap:15px; flex:1;">
                    <i class="fa-solid ${iconClass} q-icon" style="color:${iconColor}"></i>
                    <span class="q-name">${name}</span>
                    <span class="q-size">${sizeMB} MB</span>
                </div>
            `;

            let statusBadge = '<div class="q-status">Pendiente</div>';
            if (data.status === 'processing') statusBadge = '<div class="q-status processing">Procesando...</div>';
            if (data.status === 'done') statusBadge = '<div class="q-status done">Completado</div>';
            if (data.status === 'error') statusBadge = '<div class="q-status error">Error</div>';

            let actions = '';
            if (data.status === 'done') {
                actions = `<div class="q-status done-check"><i class="fa-solid fa-check"></i> Guardado</div>`;
            } else {
                actions = `<div class="btn-mini disabled"><i class="fa-solid fa-clock"></i> Esperando</div>`;
            }

            const div = document.createElement('div');
            div.className = 'queue-item';
            div.innerHTML = `
                ${infoHtml}
                ${statusBadge}
                <div class="q-actions">${actions}</div>
            `;
            queueList.appendChild(div);
        });
    }

    // === QUEUE ACTIONS ===
    if (clearQueueBtn) {
        clearQueueBtn.addEventListener('click', () => {
            selectedFiles.clear();
            renderQueue();
            const ga = document.getElementById('globalActions');
            if (ga) ga.style.display = 'none';
            log('Cola vaciada.', 'info');
        });
    }


    if (addMoreBtn) {
        addMoreBtn.addEventListener('click', () => {
            docInput.click(); // Re-open file dialog
        });
    }

    // === ADVANCED CONFIG LOGIC ===
    const openConfigBtn = document.getElementById('openConfigBtn');
    const configModal = document.getElementById('configModal');
    const closeConfigBtn = document.getElementById('closeConfigBtn');
    const saveConfigBtn = document.getElementById('saveConfigBtn');

    const importTemplateBtn = document.getElementById('importTemplateBtn');
    const exportTemplateBtn = document.getElementById('exportTemplateBtn');
    const templateInput = document.getElementById('templateInput');

    // ─── .EDD TEMPLATE: SAVE ───────────────────────────────────────────────────
    if (exportTemplateBtn) {
        exportTemplateBtn.addEventListener('click', async () => {
            const origHTML = exportTemplateBtn.innerHTML;
            exportTemplateBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';
            exportTemplateBtn.disabled = true;

            try {
                const cfg = getAdvancedConfig();
                // Also capture paper size and force-styles state
                cfg._paper_size = paperSize;
                const forceEl = document.getElementById('forceStyles');
                cfg._force_styles = forceEl ? forceEl.checked : true;

                if (typeof JSZip === 'undefined') throw new Error('JSZip no disponible');
                const zip = new JSZip();

                const manifest = { format: 'edd', version: '1.0', app: 'EDTech Doc Templater', created: new Date().toISOString() };
                zip.file('manifest.json', JSON.stringify(manifest, null, 2));
                zip.file('config.json', JSON.stringify(cfg, null, 2));

                const imgFolder = zip.folder('images');
                if (coverInput && coverInput.files.length > 0) imgFolder.file('cover.' + coverInput.files[0].name.split('.').pop(), coverInput.files[0]);
                if (headerInput && headerInput.files.length > 0) imgFolder.file('header.' + headerInput.files[0].name.split('.').pop(), headerInput.files[0]);
                if (footerInput && footerInput.files.length > 0) imgFolder.file('footer.' + footerInput.files[0].name.split('.').pop(), footerInput.files[0]);
                if (backpageInput && backpageInput.files.length > 0) imgFolder.file('backpage.' + backpageInput.files[0].name.split('.').pop(), backpageInput.files[0]);

                const blob = await zip.generateAsync({type: "blob"});
                const url = URL.createObjectURL(blob);
                const filename = prompt('Nombre del archivo de plantilla:', 'plantilla') || 'plantilla';
                if (!filename) {
                     URL.revokeObjectURL(url);
                     return;
                }
                const a = document.createElement('a');
                a.href = url;
                a.download = filename.replace(/\.edd$/i, '') + '.edd';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                log('✓ Plantilla .edd guardada.', 'success');
            } catch (e) {
                log(`Error guardando plantilla: ${e.message}`, 'error');
            } finally {
                exportTemplateBtn.innerHTML = origHTML;
                exportTemplateBtn.disabled = false;
            }
        });
    }

    // ─── .EDD TEMPLATE: LOAD ───────────────────────────────────────────────────
    if (importTemplateBtn) {
        importTemplateBtn.addEventListener('click', () => templateInput.click());
    }

    if (templateInput) {
        templateInput.addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (!file) return;
            templateInput.value = ''; // reset so same file can be re-loaded

            const origHTML = importTemplateBtn.innerHTML;
            importTemplateBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';
            importTemplateBtn.disabled = true;

            log('Cargando plantilla .edd…', 'info');

            try {
                if (typeof JSZip === 'undefined') throw new Error('JSZip no disponible');
                const zip = await JSZip.loadAsync(file);

                if (!zip.file('config.json')) throw new Error('Archivo config.json no encontrado en el EDD');
                const configText = await zip.file('config.json').async('string');
                const d = { config: JSON.parse(configText), images: {} };

                // Restore config fields
                applyConfig(d.config || {});

                // Restore image thumbnails
                // Trigger input change to load them into file objects
                async function restoreImage(slot, inputEl) {
                    const imgFolder = zip.folder('images');
                    if (!imgFolder) return;
                    let foundName = null;
                    imgFolder.forEach((relPath, f) => {
                        if (relPath.startsWith(slot)) foundName = relPath;
                    });
                    if (foundName) {
                        const fContent = imgFolder.file(foundName);
                        if (fContent) {
                            const blob = await fContent.async('blob');
                            const dt = new DataTransfer();
                            let ext = foundName.split('.').pop().toLowerCase();
                            let mime = blob.type || (ext === 'png' ? 'image/png' : 'image/jpeg');
                            dt.items.add(new File([blob], `${slot}.${ext}`, { type: mime }));
                            if (inputEl) {
                                inputEl.files = dt.files;
                                inputEl.dispatchEvent(new Event('change'));
                            }
                        }
                    }
                }

                await restoreImage('cover', coverInput);
                await restoreImage('header', headerInput);
                await restoreImage('footer', footerInput);
                await restoreImage('backpage', backpageInput);

                log('✓ Plantilla .edd cargada correctamente.', 'success');
            } catch (e) {
                log(`Error al cargar plantilla: ${e.message}`, 'error');
            } finally {
                importTemplateBtn.innerHTML = origHTML;
                importTemplateBtn.disabled = false;
            }
        });
    }

    // ─── APPLY CONFIG TO FORM ──────────────────────────────────────────────────
    function applyConfig(cfg) {
        if (!cfg) return;

        const setVal = (id, val) => { const el = document.getElementById(id); if (el && val !== undefined && val !== null) el.value = val; };
        const setBool = (id, val) => { const el = document.getElementById(id); if (el && val !== undefined) el.checked = Boolean(val); };
        const setColor = (pickerId, hexId, val) => {
            if (!val) return;
            // Normalize: always ensure # prefix
            let normalized = String(val).trim();
            if (normalized && !normalized.startsWith('#')) normalized = '#' + normalized;
            // Validate it's a real hex color before applying
            const isValid = /^#[0-9A-Fa-f]{6}$/.test(normalized);
            const safeVal = isValid ? normalized : null;
            if (!safeVal) return;
            const p = document.getElementById(pickerId);
            const h = document.getElementById(hexId);
            if (p) p.value = safeVal;
            if (h) h.value = safeVal.toUpperCase();
        };

        // General
        setVal('cfgFont', cfg.font_name);
        setVal('cfgFontSize', cfg.font_size);
        setVal('cfgLineSpacing', cfg.line_spacing);
        setVal('cfgTextAlignment', cfg.text_align);
        setColor('cfgLinkColor', 'cfgLinkColorHex', cfg.link_color);

        // Paper size
        if (cfg._paper_size) {
            paperSize = cfg._paper_size;
            document.querySelectorAll('.paper-card').forEach(c => {
                c.classList.toggle('active', c.getAttribute('data-value') === paperSize);
            });
        }
        // Force styles
        if (cfg._force_styles !== undefined) {
            const el = document.getElementById('forceStyles');
            if (el) el.checked = Boolean(cfg._force_styles);
        }

        // Cover
        const cv = cfg.cover || {};
        setVal('cfgCoverFont', cv.font);
        setVal('cfgCoverFontSize', cv.size);
        setVal('cfgCoverAlign', cv.align);
        setColor('cfgCoverColor', 'cfgCoverColorHex', cv.color);
        setBool('cfgCoverBold', cv.bold);
        setBool('cfgCoverItalic', cv.italic);
        if (cv.pos_y !== undefined) {
            setVal('cfgCoverPosY', cv.pos_y);
            const lbl = document.getElementById('coverPosYLabel');
            if (lbl) lbl.textContent = cv.pos_y;
        }
        if (cv.pos_x !== undefined) {
            setVal('cfgCoverPosX', cv.pos_x);
            const lbl = document.getElementById('coverPosXLabel');
            if (lbl) lbl.textContent = cv.pos_x;
        }
        if (cv.width !== undefined) {
            setVal('cfgCoverWidth', cv.width);
            const lbl = document.getElementById('coverWidthLabel');
            if (lbl) lbl.textContent = cv.width;
        }

        // Headings
        const hd = cfg.headings || {};
        const dt = hd.doc_title || {};
        setVal('cfgDocTitleFont', dt.font);
        setVal('cfgDocTitleSize', dt.size);
        setVal('cfgDocTitleAlign', dt.align);
        setColor('cfgDocTitleColor', 'cfgDocTitleColorHex', dt.color);
        setBool('cfgDocTitleBold', dt.bold);
        setBool('cfgDocTitleItalic', dt.italic);

        const h1 = hd.h1 || {};
        setVal('cfgH1Size', h1.size);
        setColor('cfgH1Color', 'cfgH1ColorHex', h1.color);
        setBool('cfgH1Bold', h1.bold);

        const h2 = hd.h2 || {};
        setVal('cfgH2Size', h2.size);
        setColor('cfgH2Color', 'cfgH2ColorHex', h2.color);
        setBool('cfgH2Bold', h2.bold);

        const h3 = hd.h3 || {};
        setVal('cfgH3Size', h3.size);
        setColor('cfgH3Color', 'cfgH3ColorHex', h3.color);
        setBool('cfgH3Bold', h3.bold);

        // Page numbers
        const pn = cfg.page_numbers || {};
        setVal('cfgPageNumStyle', pn.style);
        setVal('cfgPageNumPosition', pn.position);
        setVal('cfgPageNumFormat', pn.format);
        setBool('cfgPageNumEnabled', pn.enabled);

        // Margins
        const mg = cfg.margins || {};
        setVal('cfgMarginTop', mg.top);
        setVal('cfgMarginBottom', mg.bottom);
        setVal('cfgMarginLeft', mg.left);
        setVal('cfgMarginRight', mg.right);

        // Tables
        const tb = cfg.tables || {};
        setColor('cfgTableHeaderBg', 'cfgTableHeaderBgHex', tb.header_bg);
        setColor('cfgTableHeaderText', 'cfgTableHeaderTextHex', tb.header_text);
        setVal('cfgTableBorderV', tb.border_v || 'single');
        setVal('cfgTableBorderH', tb.border_h || 'single');
        setColor('cfgTableBorderVColor', 'cfgTableBorderVColorHex', tb.border_v_color || '#000000');
        setColor('cfgTableBorderHColor', 'cfgTableBorderHColorHex', tb.border_h_color || '#000000');
        setColor('cfgTableOutlineColor', 'cfgTableOutlineColorHex', tb.border_outline_color || '#000000');
        setVal('cfgTableOutlineSz', tb.border_outline_sz || '4');
        document.getElementById('cfgTableZebra').checked = (tb.zebra !== false);
        setColor('cfgTableZebraColor', 'cfgTableZebraColorHex', tb.zebra_color || '#f1f5f9');
        setBool('cfgTableNumberAlign', tb.align_numbers);
        setBool('cfgOcrTables', tb.ocr_tables);

        // TOC
        const toc = cfg.toc || {};
        setBool('cfgToc', toc.enabled);
        setVal('cfgTocTitle', toc.title || 'ÍNDICE');
        setVal('cfgTocTitleSize', toc.title_size || '18');
        setColor('cfgTocTitleColor', 'cfgTocTitleColorHex', toc.title_color || '#000000');
        setBool('cfgTocTitleBold', toc.title_bold !== false);
        setBool('cfgTocTitleItalic', toc.title_italic);
        setVal('cfgTocStyle', toc.style || 'dotted');
        setVal('cfgTocDepth', toc.depth || '2');
        setColor('cfgTocEntryColor', 'cfgTocEntryColorHex', toc.entry_color || '#000000');
        setColor('cfgTocPageColor', 'cfgTocPageColorHex', toc.page_color || '#000000');
        setColor('cfgTocLeaderColor', 'cfgTocLeaderColorHex', toc.leader_color || '#999999');
        // Enable/disable tab button
        if (tabTocBtn) tabTocBtn.disabled = !toc.enabled;

        // Refresh live previews
        updateTablePreview();
        updateCoverPreview();
        updateMarginsPreview();
    }

    // Auto-load saved config from localStorage on page init
    try {
        const saved = localStorage.getItem('doc_templater_config');
        if (saved) {
            applyConfig(JSON.parse(saved));
            console.log('Config restored from localStorage');
        }
    } catch (e) { console.warn('Config load error', e); }

    // Force default toggle states (always on fresh load)
    const _fs = document.getElementById('forceStyles');
    if (_fs) _fs.checked = true;
    const _ocr = document.getElementById('cfgOcrTables');
    if (_ocr) _ocr.checked = false;
    const _toc = document.getElementById('cfgToc');
    if (_toc) { _toc.checked = false; if (tabTocBtn) tabTocBtn.disabled = true; }


    // === TABLE PREVIEW LOGIC ===
    function updateTablePreview() {
        const bg = document.getElementById('cfgTableHeaderBg').value;
        const text = document.getElementById('cfgTableHeaderText').value;

        const colorV = document.getElementById('cfgTableBorderVColor').value;
        const colorH = document.getElementById('cfgTableBorderHColor').value;
        const colorO = document.getElementById('cfgTableOutlineColor').value;
        const zebraColor = document.getElementById('cfgTableZebraColor')?.value || '#f1f5f9';

        const borderUsageV = document.getElementById('cfgTableBorderV').value;
        const borderUsageH = document.getElementById('cfgTableBorderH').value;
        const zebra = document.getElementById('cfgTableZebra').checked;

        const table = document.querySelector('.preview-table');
        if (!table) return;

        // Apply Styles via CSS Variables
        table.style.setProperty('--preview-bg', bg);
        table.style.setProperty('--preview-text', text);

        function getBorderCss(type, color) {
            if (type === 'none') return 'none';
            if (type === 'thick') return '2px solid ' + color;
            if (type === 'dashed') return '1px dashed ' + color;
            return '1px solid ' + color;
        }

        const borderV = getBorderCss(borderUsageV, colorV);
        const borderH = getBorderCss(borderUsageH, colorH);
        // Outline - assuming simple solid for preview if any inside usage is active, or user just wants outline color
        // We'll apply it to the table element which often has border-collapse
        // But standard HTML table border applies to outline.
        table.style.border = '1px solid ' + colorO;

        table.style.setProperty('--preview-border-v', borderV);
        table.style.setProperty('--preview-border-h', borderH);

        // Zebra Striping
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach((row, index) => {
            if (zebra && index % 2 !== 0) {
                row.style.background = zebraColor;
            } else {
                row.style.background = 'transparent';
            }
        });
    }

    // Listeners for Preview
    const previewInputs = ['cfgTableHeaderBg', 'cfgTableHeaderText', 'cfgTableBorderV', 'cfgTableBorderH',
        'cfgTableBorderVColor', 'cfgTableBorderHColor', 'cfgTableOutlineColor', 'cfgTableZebraColor', 'cfgTableZebra'];
    previewInputs.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', updateTablePreview);
            el.addEventListener('change', updateTablePreview);
        }
    });

    // === COLOR PICKER SYNC ===
    const colorPairs = [
        { picker: 'cfgLinkColor', hex: 'cfgLinkColorHex' },
        { picker: 'cfgCoverColor', hex: 'cfgCoverColorHex' },
        { picker: 'cfgH1Color', hex: 'cfgH1ColorHex' },
        { picker: 'cfgH2Color', hex: 'cfgH2ColorHex' },
        { picker: 'cfgH3Color', hex: 'cfgH3ColorHex' },
        { picker: 'cfgTableHeaderBg', hex: 'cfgTableHeaderBgHex' },
        { picker: 'cfgTableHeaderText', hex: 'cfgTableHeaderTextHex' },
        { picker: 'cfgTableBorderVColor', hex: 'cfgTableBorderVColorHex' },
        { picker: 'cfgTableBorderHColor', hex: 'cfgTableBorderHColorHex' },
        { picker: 'cfgTableOutlineColor', hex: 'cfgTableOutlineColorHex' },
        { picker: 'cfgTableZebraColor', hex: 'cfgTableZebraColorHex' },
        { picker: 'cfgTocTitleColor', hex: 'cfgTocTitleColorHex' },
        { picker: 'cfgTocEntryColor', hex: 'cfgTocEntryColorHex' },
        { picker: 'cfgTocPageColor', hex: 'cfgTocPageColorHex' },
        { picker: 'cfgTocLeaderColor', hex: 'cfgTocLeaderColorHex' },
        { picker: 'cfgDocTitleColor', hex: 'cfgDocTitleColorHex' }
    ];

    colorPairs.forEach(pair => {
        const pEl = document.getElementById(pair.picker);
        const hEl = document.getElementById(pair.hex);

        if (pEl && hEl) {
            // Picker -> Hex
            pEl.addEventListener('input', () => {
                hEl.value = pEl.value.toUpperCase();
                updateTablePreview(); // Update preview if applicable
            });
            pEl.addEventListener('change', () => {
                hEl.value = pEl.value.toUpperCase();
                updateTablePreview();
            });

            // Hex -> Picker
            hEl.addEventListener('input', () => {
                let val = hEl.value.trim();
                // Auto-prepend # if user types without it
                if (val && !val.startsWith('#')) {
                    val = '#' + val;
                    hEl.value = val; // update display
                }
                const isValidHex = /^#[0-9A-F]{6}$/i.test(val);
                if (isValidHex) {
                    pEl.value = val;
                    updateTablePreview();
                }
            });
        }
    });

    // Tab Logic
    const tabs = document.querySelectorAll('.tab-btn');
    const contents = document.querySelectorAll('.tab-content');

    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            // Remove active class from all tabs
            tabs.forEach(t => t.classList.remove('active'));
            contents.forEach(c => c.classList.add('hidden-tab'));

            // Activate clicked tab
            tab.classList.add('active');
            const targetId = tab.dataset.tab;
            document.getElementById(targetId).classList.remove('hidden-tab');
        });
    });

    if (openConfigBtn && configModal) {
        openConfigBtn.addEventListener('click', () => {
            configModal.classList.add('active');
        });
    }

    if (openConfigBtn) {
        openConfigBtn.addEventListener('click', () => {
            configModal.classList.add('active');
        });
    }

    if (closeConfigBtn) {
        closeConfigBtn.addEventListener('click', () => {
            configModal.classList.remove('active');
        });
    }

    if (configModal) {
        configModal.addEventListener('click', (e) => {
            if (e.target === configModal) configModal.classList.remove('active');
        });
    }

    if (saveConfigBtn) {
        saveConfigBtn.addEventListener('click', () => {
            configModal.classList.remove('active');
            // Persist to localStorage
            try {
                const cfg = getAdvancedConfig();
                cfg._paper_size = paperSize;
                const forceEl = document.getElementById('forceStyles');
                cfg._force_styles = forceEl ? forceEl.checked : true;
                localStorage.setItem('doc_templater_config', JSON.stringify(cfg));
            } catch (e) { console.warn('Config save error', e); }
            log('Configuración guardada ✓', 'success');
        });
    }

    // Capture Advanced Config
    function getAdvancedConfig() {
        const cfgFont = document.getElementById('cfgFont');
        if (!cfgFont) return {};

        const getBool = (id) => { const el = document.getElementById(id); return el ? el.checked : false; };
        const getVal = (id, def) => { const el = document.getElementById(id); return el ? el.value : def; };

        return {
            font_name: getVal('cfgFont', 'Calibri'),
            font_size: getVal('cfgFontSize', '11'),
            line_spacing: getVal('cfgLineSpacing', '1.15'),
            link_color: getVal('cfgLinkColor', '#0563C1'),
            text_align: getVal('cfgTextAlignment', 'left'),

            cover: {
                font: getVal('cfgCoverFont', 'Calibri'),
                size: getVal('cfgCoverFontSize', '36'),
                color: getVal('cfgCoverColor', '#000000'),
                align: getVal('cfgCoverAlign', 'center'),
                bold: getBool('cfgCoverBold'),
                italic: getBool('cfgCoverItalic'),
                pos_y: parseInt(getVal('cfgCoverPosY', '55'), 10),
                pos_x: parseInt(getVal('cfgCoverPosX', '50'), 10),
                width: parseInt(getVal('cfgCoverWidth', '80'), 10),
            },

            headings: {
                doc_title: {
                    font: getVal('cfgDocTitleFont', ''),
                    size: getVal('cfgDocTitleSize', '18'),
                    color: getVal('cfgDocTitleColor', '#000000'),
                    align: getVal('cfgDocTitleAlign', 'left'),
                    bold: getBool('cfgDocTitleBold'),
                    italic: getBool('cfgDocTitleItalic'),
                },
                h1: {
                    size: getVal('cfgH1Size', '16'),
                    color: getVal('cfgH1Color', '#000000'),
                    bold: getBool('cfgH1Bold')
                },
                h2: {
                    size: getVal('cfgH2Size', '13'),
                    color: getVal('cfgH2Color', '#2c6eff'),
                    bold: getBool('cfgH2Bold')
                },
                h3: {
                    size: getVal('cfgH3Size', '12'),
                    color: getVal('cfgH3Color', '#555555'),
                    bold: getBool('cfgH3Bold')
                }
            },

            page_numbers: {
                enabled: getBool('cfgPageNumEnabled'),
                style: getVal('cfgPageNumStyle', 'arabic'),
                position: getVal('cfgPageNumPosition', 'center'),
                format: getVal('cfgPageNumFormat', 'page_only'),
            },

            margins: {
                top: parseFloat(getVal('cfgMarginTop', '3.7')),
                bottom: parseFloat(getVal('cfgMarginBottom', '3.0')),
                left: parseFloat(getVal('cfgMarginLeft', '3.0')),
                right: parseFloat(getVal('cfgMarginRight', '3.0')),
            },

            tables: {
                header_bg: getVal('cfgTableHeaderBg', '#2c6eff'),
                header_text: getVal('cfgTableHeaderText', '#FFFFFF'),
                border_v: getVal('cfgTableBorderV', 'single'),
                border_h: getVal('cfgTableBorderH', 'single'),
                border_v_color: getVal('cfgTableBorderVColor', '#000000'),
                border_h_color: getVal('cfgTableBorderHColor', '#000000'),
                border_outline_color: getVal('cfgTableOutlineColor', '#000000'),
                border_outline_sz: parseInt(getVal('cfgTableOutlineSz', '4'), 10),
                zebra: getBool('cfgTableZebra'),
                zebra_color: getVal('cfgTableZebraColor', '#f1f5f9'),
                align_numbers: getBool('cfgTableNumberAlign'),
                ocr_tables: getBool('cfgOcrTables'),
            },

            toc: {
                enabled: getBool('cfgToc'),
                title: getVal('cfgTocTitle', 'ÍNDICE'),
                title_size: getVal('cfgTocTitleSize', '18'),
                title_color: getVal('cfgTocTitleColor', '#000000'),
                title_bold: getBool('cfgTocTitleBold'),
                title_italic: getBool('cfgTocTitleItalic'),
                style: getVal('cfgTocStyle', 'dotted'),
                depth: parseInt(getVal('cfgTocDepth', '2'), 10),
                entry_color: getVal('cfgTocEntryColor', '#000000'),
                page_color: getVal('cfgTocPageColor', '#000000'),
                leader_color: getVal('cfgTocLeaderColor', '#999999'),
            },
        };
    }

    // === PROCESS LOGIC ===
    if (processBtn) {
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
            
            // ZIP instance for multiple files
            let outputZip = null;
            if (typeof JSZip !== 'undefined' && selectedFiles.size > 1) {
                outputZip = new JSZip();
            }
            let outFilesCount = 0;
            
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
                    
                    let outName = prefix + name.replace(/(\.docx)$/i, '') + suffix + '.docx';
                    
                    if (outputZip) {
                        // Add to zip
                        outputZip.file(outName, blob);
                        outFilesCount++;
                    } else {
                        // Trigger download immediately
                        const url = URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = outName;
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.revokeObjectURL(url);
                    }
                    
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
            
            if (outputZip && outFilesCount > 0) {
                log('Empaquetando documentos en ZIP...', 'info');
                try {
                    const zipBlob = await outputZip.generateAsync({type:"blob"});
                    const url = URL.createObjectURL(zipBlob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'Documentos_Procesados.zip';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                    log('Descarga: Documentos_Procesados.zip iniciada.', 'success');
                } catch(e) {
                    log('Error generando ZIP: ' + e.message, 'error');
                }
            }
            log('✅ Cola finalizada — ' + success_count + ' archivo(s) procesados correctamente.', 'success');
            if (error_count > 0) log('⚠️ ' + error_count + ' archivo(s) fallaron.', 'error');
            
            processBtn.disabled = false;
            processBtn.innerHTML = '<span>INICIAR PROCESO</span><i class="fa-solid fa-play"></i>';
        });
    }

    // === PAPER SELECTOR ===
    document.querySelectorAll('.paper-card').forEach(card => {
        card.addEventListener('click', () => {
            document.querySelectorAll('.paper-card').forEach(c => c.classList.remove('active'));
            card.classList.add('active');
            paperSize = card.getAttribute('data-value');
        });
    });

    // === COVER PREVIEW LOGIC ===
    function updateCoverPreview() {
        const titleEl = document.getElementById('coverPreviewTitle');
        const guideEl = document.getElementById('coverPreviewGuide');
        const bgEl = document.getElementById('coverPreviewBg');
        const posLabel = document.getElementById('coverPosYLabel');
        if (!titleEl) return;

        const font = document.getElementById('cfgCoverFont')?.value || 'Calibri';
        const size = parseInt(document.getElementById('cfgCoverFontSize')?.value || '36', 10);
        const color = document.getElementById('cfgCoverColor')?.value || '#000000';
        const align = document.getElementById('cfgCoverAlign')?.value || 'center';
        const bold = document.getElementById('cfgCoverBold')?.checked || false;
        const italic = document.getElementById('cfgCoverItalic')?.checked || false;
        const posY = parseInt(document.getElementById('cfgCoverPosY')?.value || '55', 10);
        const posX = parseInt(document.getElementById('cfgCoverPosX')?.value || '50', 10);
        const widthPct = parseInt(document.getElementById('cfgCoverWidth')?.value || '80', 10);

        // Scale font size for preview (real 36pt -> ~11px in mini)
        const scaledSize = Math.max(8, Math.min(16, Math.round(size / 3.2)));

        titleEl.style.fontFamily = font + ', sans-serif';
        titleEl.style.fontSize = scaledSize + 'px';
        titleEl.style.color = color;
        titleEl.style.textAlign = align;
        titleEl.style.fontWeight = bold ? '700' : '400';
        titleEl.style.fontStyle = italic ? 'italic' : 'normal';
        titleEl.style.top = posY + '%';
        // In Word, posY% dictates the top edge (space_before), so we remove vertical centering
        titleEl.style.transform = 'translate(-50%, 0)';

        // Width and X position
        const w = widthPct;
        titleEl.style.width = w + '%';
        titleEl.style.left = posX + '%';
        titleEl.style.right = 'auto';

        if (guideEl) {
            guideEl.style.top = posY + '%';
        }
        if (posLabel) {
            posLabel.textContent = posY;
        }
        const posXLabel = document.getElementById('coverPosXLabel');
        if (posXLabel) posXLabel.textContent = posX;
        const widthLabel = document.getElementById('coverWidthLabel');
        if (widthLabel) widthLabel.textContent = widthPct;

        // Sync cover image from sidebar thumbnail
        if (bgEl) {
            const coverThumbEl = document.getElementById('coverThumb');
            if (coverThumbEl && coverThumbEl.classList.contains('has-image')) {
                bgEl.style.backgroundImage = coverThumbEl.style.backgroundImage;
            } else {
                bgEl.style.backgroundImage = 'none';
            }
        }
    }

    // Cover preview live listeners
    const coverPreviewInputs = ['cfgCoverFont', 'cfgCoverFontSize', 'cfgCoverColor', 'cfgCoverColorHex',
        'cfgCoverAlign', 'cfgCoverBold', 'cfgCoverItalic', 'cfgCoverPosY', 'cfgCoverPosX', 'cfgCoverWidth'];
    coverPreviewInputs.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', updateCoverPreview);
            el.addEventListener('change', updateCoverPreview);
        }
    });

    // Position slider label sync
    ['cfgCoverPosY:coverPosYLabel', 'cfgCoverPosX:coverPosXLabel', 'cfgCoverWidth:coverWidthLabel'].forEach(pair => {
        const [sliderId, labelId] = pair.split(':');
        const slider = document.getElementById(sliderId);
        if (slider) {
            slider.addEventListener('input', () => {
                const lbl = document.getElementById(labelId);
                if (lbl) lbl.textContent = slider.value;
            });
        }
    });

    // === DRAG TITLE IN PREVIEW (X + Y) ===
    const coverPreview = document.getElementById('coverMiniPreview');
    const coverTitle = document.getElementById('coverPreviewTitle');
    if (coverPreview && coverTitle) {
        let isDragging = false;

        coverTitle.addEventListener('mousedown', (e) => {
            e.preventDefault();
            isDragging = true;
            coverTitle.style.transition = 'none';
            const guideEl = document.getElementById('coverPreviewGuide');
            if (guideEl) guideEl.style.transition = 'none';
        });

        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            const rect = coverPreview.getBoundingClientRect();
            const relY = e.clientY - rect.top;
            const relX = e.clientX - rect.left;
            let pctY = Math.round((relY / rect.height) * 100);
            let pctX = Math.round((relX / rect.width) * 100);
            pctY = Math.max(10, Math.min(90, pctY));
            pctX = Math.max(10, Math.min(90, pctX));

            const sliderY = document.getElementById('cfgCoverPosY');
            if (sliderY) sliderY.value = pctY;
            const lblY = document.getElementById('coverPosYLabel');
            if (lblY) lblY.textContent = pctY;

            const sliderX = document.getElementById('cfgCoverPosX');
            if (sliderX) sliderX.value = pctX;
            const lblX = document.getElementById('coverPosXLabel');
            if (lblX) lblX.textContent = pctX;

            coverTitle.style.top = pctY + '%';
            coverTitle.style.left = pctX + '%';
            const guideEl = document.getElementById('coverPreviewGuide');
            if (guideEl) guideEl.style.top = pctY + '%';
        });

        document.addEventListener('mouseup', () => {
            if (isDragging) {
                isDragging = false;
                coverTitle.style.transition = '';
                const guideEl = document.getElementById('coverPreviewGuide');
                if (guideEl) guideEl.style.transition = '';
            }
        });
    }

    // === MARGINS PREVIEW LOGIC ===
    function updateMarginsPreview() {
        const top = parseFloat(document.getElementById('cfgMarginTop')?.value || '3.7');
        const bottom = parseFloat(document.getElementById('cfgMarginBottom')?.value || '3.0');
        const left = parseFloat(document.getElementById('cfgMarginLeft')?.value || '3.0');
        const right = parseFloat(document.getElementById('cfgMarginRight')?.value || '3.0');

        // Page is ~27.9cm tall (letter), scale margins to percentages
        const pageH = 27.9, pageW = 21.6;
        const topPct = (top / pageH * 100).toFixed(1);
        const bottomPct = (bottom / pageH * 100).toFixed(1);
        const leftPct = (left / pageW * 100).toFixed(1);
        const rightPct = (right / pageW * 100).toFixed(1);

        const zt = document.getElementById('marginZoneTop');
        const zb = document.getElementById('marginZoneBottom');
        const zl = document.getElementById('marginZoneLeft');
        const zr = document.getElementById('marginZoneRight');
        const cz = document.getElementById('marginContentZone');

        if (zt) zt.style.height = topPct + '%';
        if (zb) zb.style.height = bottomPct + '%';
        if (zl) zl.style.width = leftPct + '%';
        if (zr) zr.style.width = rightPct + '%';

        if (cz) {
            cz.style.top = topPct + '%';
            cz.style.bottom = bottomPct + '%';
            cz.style.left = leftPct + '%';
            cz.style.right = rightPct + '%';
        }

        // Labels
        const lt = document.getElementById('mlTop');
        const lb = document.getElementById('mlBottom');
        const ll = document.getElementById('mlLeft');
        const lr = document.getElementById('mlRight');
        if (lt) lt.textContent = top.toFixed(1);
        if (lb) lb.textContent = bottom.toFixed(1);
        if (ll) ll.textContent = left.toFixed(1);
        if (lr) lr.textContent = right.toFixed(1);

        // Sync header/footer images from sidebar
        const headerPreview = document.getElementById('marginPreviewHeader');
        const footerPreview = document.getElementById('marginPreviewFooter');
        const headerThumb = document.getElementById('headerThumb');
        const footerThumb = document.getElementById('footerThumb');

        if (headerPreview && headerThumb && headerThumb.classList.contains('has-image')) {
            headerPreview.style.backgroundImage = headerThumb.style.backgroundImage;
        } else if (headerPreview) {
            headerPreview.style.backgroundImage = 'none';
        }
        if (footerPreview && footerThumb && footerThumb.classList.contains('has-image')) {
            footerPreview.style.backgroundImage = footerThumb.style.backgroundImage;
        } else if (footerPreview) {
            footerPreview.style.backgroundImage = 'none';
        }
    }

    // Margins preview listeners
    ['cfgMarginTop', 'cfgMarginBottom', 'cfgMarginLeft', 'cfgMarginRight'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', updateMarginsPreview);
            el.addEventListener('change', updateMarginsPreview);
        }
    });

    // === TAB SWITCH — refresh previews ===
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const target = btn.getAttribute('data-tab');
            if (target === 'tab-cover') setTimeout(updateCoverPreview, 50);
            if (target === 'tab-margins') setTimeout(updateMarginsPreview, 50);
            if (target === 'tab-tables') setTimeout(updateTablePreview, 50);
        });
    });

    // Initial render
    updateCoverPreview();
    updateMarginsPreview();

    // === UTILS ===
    function log(msg, type = 'info') {
        if (!consoleLogs) { console.log(`[${type}] ${msg}`); return; }
        const div = document.createElement('div');
        div.className = `log-item ${type}`;
        const time = new Date().toLocaleTimeString();
        div.innerText = `[${time}] ${msg}`;
        consoleLogs.appendChild(div);
        consoleLogs.scrollTop = consoleLogs.scrollHeight;
        // Also append to debug console
        const dl = document.getElementById('debugLog');
        if (dl) {
            dl.textContent += `[${time}] [${type.toUpperCase()}] ${msg}\n`;
            dl.scrollTop = dl.scrollHeight;
        }
    }
});
