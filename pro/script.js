
        lucide.createIcons();

        // UI references
        const dropzone = document.getElementById('dropzone');
        const fileInput = document.getElementById('fileInput');
        const folderInput = document.getElementById('folderInput');
        const stateIcon = document.getElementById('stateIcon');
        const stateTitle = document.getElementById('stateTitle');
        const stateDesc = document.getElementById('stateDesc');
        const btnPreview = document.getElementById('btnPreview');
        const btnCode = document.getElementById('btnCode');
        const btnCopy = document.getElementById('btnCopy');
        const btnDownload = document.getElementById('btnDownload');
        const btnDownloadText = document.getElementById('btnDownloadText');
        const previewArea = document.getElementById('previewArea');
        const codeArea = document.getElementById('codeArea');
        const emptyState = document.getElementById('emptyState');
        const fileSidebar = document.getElementById('fileSidebar');
        const fileList = document.getElementById('fileList');
        const batchCountBadge = document.getElementById('batchCountBadge');
        const chkLimitRows = document.getElementById('chkLimitRows');
        const chkExcludeHidden = document.getElementById('chkExcludeHidden');

        // State
        let currentResults = [];
        let currentActiveIndex = 0;

        // Helpers
        function getColumnLetter(colIndex) {
            let temp, letter = '';
            while (colIndex > 0) {
                temp = (colIndex - 1) % 26;
                letter = String.fromCharCode(temp + 65) + letter;
                colIndex = (colIndex - temp - 1) / 26;
            }
            return letter;
        }

        function getTextColor(bgHex) {
            if (!bgHex) return '#FFFFFF';
            const hex = bgHex.replace('#', '');
            const r = parseInt(hex.substr(0, 2), 16) || 0;
            const g = parseInt(hex.substr(2, 2), 16) || 0;
            const b = parseInt(hex.substr(4, 2), 16) || 0;
            const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
            return luminance > 0.5 ? '#000000' : '#FFFFFF';
        }

        // Format cell value safely (no [object Object], handling formulas, richText, hyperlinks, dates, and merged cells)
        function formatCellValue(cell) {
            if (!cell || cell.value === null || cell.value === undefined) return '';

            // Merged Cell Slave: keep empty to avoid repeating text and cluttering tables
            if (cell.isMerged && cell.master && cell.address !== cell.master.address) {
                return '';
            }

            const val = cell.value;

            // Formulas: cell.formula or cell.value formula objects
            const formula = cell.formula || (val && typeof val === 'object' ? (val.formula || val.sharedFormula) : null);
            if (formula) {
                let resVal = cell.result !== undefined ? cell.result : (val && typeof val === 'object' && val.result !== undefined ? val.result : '');
                let cleanRes = '';
                if (resVal instanceof Date) {
                    cleanRes = resVal.toLocaleDateString('vi-VN');
                } else if (resVal !== null && resVal !== undefined && typeof resVal !== 'object') {
                    cleanRes = String(resVal);
                } else if (resVal && typeof resVal === 'object' && resVal.error) {
                    cleanRes = String(resVal.error);
                }
                return cleanRes !== '' ? `${cleanRes} (\`=${formula}\`)` : `\`=${formula}\``;
            }

            // RichText Object
            if (typeof val === 'object' && Array.isArray(val.richText)) {
                return val.richText.map(t => t.text || '').join('');
            }

            // Hyperlink Object
            if (typeof val === 'object' && (val.text || val.hyperlink)) {
                const text = val.text || val.hyperlink;
                return val.hyperlink ? `[${text}](${val.hyperlink})` : String(text);
            }

            // Error Object
            if (typeof val === 'object' && val.error) {
                return String(val.error);
            }

            // Date Object
            if (val instanceof Date) {
                return val.toLocaleDateString('vi-VN');
            }

            // Primitive Values
            if (typeof val !== 'object') {
                return String(val);
            }

            return '';
        }

        // 1. EXCEL CONVERTER (ExcelJS + JSZip Hybrid Engine)
        async function convertExcel(buffer, filename, excludeHidden) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);

            const parts = [];
            const partsPreview = [];
            const extractedImages = {};
            const assignedImages = new Set();
            let imgCounter = 0;

            const styleBlock = `<style>
  .excel-table-wrap { overflow-x: auto; margin-bottom: 24px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.08); background: rgba(0,0,0,0.2); }
  .excel-table { border-collapse: collapse; width: 100%; min-width: 650px; font-size: 13px; }
  .excel-table th { border: 1px solid rgba(255,255,255,0.1); padding: 6px 10px; text-align: center; font-weight: 600; background-color: rgba(255,255,255,0.06); color: #cbd5e1; }
  .excel-table td { border: 1px solid rgba(255,255,255,0.08); padding: 6px 10px; vertical-align: top; white-space: pre-wrap; }
  .excel-table .row-idx { font-weight: 600; background-color: rgba(255,255,255,0.04); text-align: center; width: 45px; color: #94a3b8; }
</style>\n`;

            parts.push(styleBlock);
            partsPreview.push(styleBlock);

            // 1. Extract all media images from ExcelJS
            const imageMap = {};
            if (workbook.media && workbook.media.length) {
                workbook.media.forEach((med, idx) => {
                    if (med.type === 'image' || med.buffer) {
                        imgCounter++;
                        const ext = med.extension || 'png';
                        const saveName = `excel_img_${imgCounter}.${ext}`;
                        const base64Str = med.buffer.toString('base64');
                        const mime = `image/${ext === 'jpg' ? 'jpeg' : ext}`;
                        const dataUri = `data:${mime};base64,${base64Str}`;
                        
                        extractedImages[saveName] = { base64: base64Str, mime: mime, data_uri: dataUri };
                        imageMap[med.index !== undefined ? med.index : idx] = saveName;
                    }
                });
            }

            // 2. Extra media scan via JSZip in xl/media/
            try {
                const zip = await JSZip.loadAsync(buffer);
                for (const [relativePath, zipEntry] of Object.entries(zip.files)) {
                    if (relativePath.startsWith('xl/media/') && !zipEntry.dir) {
                        const ext = relativePath.split('.').pop() || 'png';
                        const base64Str = await zipEntry.async('base64');
                        const exists = Object.values(extractedImages).some(im => im.base64 === base64Str);
                        if (!exists) {
                            imgCounter++;
                            const saveName = `excel_img_${imgCounter}.${ext}`;
                            const mime = `image/${ext === 'jpg' ? 'jpeg' : ext}`;
                            const dataUri = `data:${mime};base64,${base64Str}`;
                            extractedImages[saveName] = { base64: base64Str, mime: mime, data_uri: dataUri };
                        }
                    }
                }
            } catch (e) {
                console.warn('JSZip media scan warning:', e);
            }

            // 3. Process each worksheet
            workbook.eachSheet((worksheet, sheetId) => {
                if (excludeHidden && worksheet.state === 'hidden') return;

                const sheetTitle = worksheet.state === 'hidden' ? `${worksheet.name} (sheet ẩn)` : worksheet.name;
                parts.push(`\n## ${sheetTitle}\n`);
                partsPreview.push(`\n## ${sheetTitle}\n`);

                // Check and attach sheet images
                const sheetImageObjects = worksheet.getImages ? worksheet.getImages() : [];
                const sheetImageNames = [];
                if (sheetImageObjects && sheetImageObjects.length) {
                    sheetImageObjects.forEach(imgObj => {
                        const imgName = imageMap[imgObj.imageId];
                        if (imgName) {
                            sheetImageNames.push(imgName);
                            assignedImages.add(imgName);
                        }
                    });
                }

                if (sheetImageNames.length > 0) {
                    const imgMd = sheetImageNames.map(name => `![${name}](${name})`).join('\n\n') + '\n';
                    parts.push(imgMd);
                    partsPreview.push(imgMd);
                }

                const maxCol = worksheet.columnCount || 0;

                // Collect row content
                const rowsData = [];
                worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                    const rowCells = [];
                    let hasDataInRow = false;
                    const colLimit = Math.max(maxCol, row.cellCount || 1);

                    for (let c = 1; c <= colLimit; c++) {
                        const cell = row.getCell(c);
                        const text = formatCellValue(cell);
                        if (text) hasDataInRow = true;

                        let tdStyle = '';
                        if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) {
                            let argb = cell.fill.fgColor.argb;
                            if (argb && argb.length === 8) {
                                const hex = `#${argb.substring(2)}`;
                                const txtColor = getTextColor(hex);
                                tdStyle = `background-color: ${hex}; color: ${txtColor};`;
                            }
                        }

                        const styleAttr = tdStyle ? ` style='${tdStyle}'` : '';
                        rowCells.push(`        <td${styleAttr}>${text}</td>`);
                    }

                    if (hasDataInRow || rowCells.some(rc => rc.includes('style='))) {
                        rowsData.push({ rowNumber, rowCells, colLimit });
                    }
                });

                if (rowsData.length > 0) {
                    const tableId = `excel-table-${sheetId}`;
                    const effectiveColCount = Math.max(...rowsData.map(r => r.colLimit), 1);
                    const headerHtml = [
                        `<div class='excel-table-wrap'>`,
                        `  <table class='excel-table' id='${tableId}'>`,
                        `    <thead><tr><th class='row-idx'></th>`,
                        Array.from({ length: effectiveColCount }, (_, i) => `<th>${getColumnLetter(i + 1)}</th>`).join(''),
                        `</tr></thead><tbody>`
                    ].join('\n');

                    parts.push(headerHtml);
                    partsPreview.push(headerHtml);

                    let renderedCount = 0;
                    let previewTruncated = false;

                    rowsData.forEach(({ rowNumber, rowCells }) => {
                        const rowHtml = `      <tr>\n        <td class='row-idx'>${rowNumber}</td>\n${rowCells.join('\n')}\n      </tr>`;
                        parts.push(rowHtml);

                        if (renderedCount < 100) {
                            partsPreview.push(rowHtml);
                            renderedCount++;
                        } else if (!previewTruncated) {
                            partsPreview.push(`      <tr><td colspan='${effectiveColCount + 1}' style='text-align: center; font-style: italic; color: #94a3b8; background-color: rgba(255,255,255,0.02); padding: 12px;'>... Đã ẩn bớt các dòng từ đây để tránh lag trình duyệt. Tải file về để xem đầy đủ ...</td></tr>`);
                            previewTruncated = true;
                        }
                    });

                    const footerHtml = `    </tbody>\n  </table>\n</div>\n`;
                    parts.push(footerHtml);
                    partsPreview.push(footerHtml);
                } else if (sheetImageNames.length === 0) {
                    // No table rows and no sheet images
                    parts.push(`*(Sheet này không chứa dữ liệu bảng dạng văn bản)*\n`);
                    partsPreview.push(`*(Sheet này không chứa dữ liệu bảng dạng văn bản)*\n`);
                }
            });

            // 4. Attach unassigned images to markdown
            const unassigned = Object.keys(extractedImages).filter(name => !assignedImages.has(name));
            if (unassigned.length > 0) {
                const extraImgMd = `\n### Hình ảnh trong tài liệu Excel\n\n` + unassigned.map(name => `![${name}](${name})`).join('\n\n') + '\n';
                parts.push(extraImgMd);
                partsPreview.push(extraImgMd);
            }

            return {
                markdown: parts.join('\n'),
                markdown_preview: partsPreview.join('\n'),
                images: extractedImages,
                has_images: Object.keys(extractedImages).length > 0
            };
        }

        // 2. POWERPOINT CONVERTER (JSZip + XML DOM)
        async function convertPPTX(buffer, filename) {
            const zip = await JSZip.loadAsync(buffer);
            const parts = [];
            const extractedImages = {};
            let imgCounter = 0;

            // Extract all media images
            for (const [relativePath, zipEntry] of Object.entries(zip.files)) {
                if (relativePath.startsWith('ppt/media/') && !zipEntry.dir) {
                    imgCounter++;
                    const ext = relativePath.split('.').pop() || 'png';
                    const saveName = `slide_img_${imgCounter}.${ext}`;
                    const base64Str = await zipEntry.async('base64');
                    const mime = `image/${ext === 'jpg' ? 'jpeg' : ext}`;
                    extractedImages[saveName] = { base64: base64Str, mime: mime, data_uri: `data:${mime};base64,${base64Str}` };
                }
            }

            // Extract slide texts in order
            const slideFiles = Object.keys(zip.files).filter(k => k.match(/^ppt\/slides\/slide\d+\.xml$/)).sort((a, b) => {
                const numA = parseInt(a.match(/\d+/)[0]);
                const numB = parseInt(b.match(/\d+/)[0]);
                return numA - numB;
            });

            const parser = new DOMParser();
            for (let i = 0; i < slideFiles.length; i++) {
                parts.push(`\n## Slide ${i + 1}\n`);
                const slideXmlStr = await zip.file(slideFiles[i]).async('text');
                const doc = parser.parseFromString(slideXmlStr, 'application/xml');

                // Extract paragraphs & tables
                const textNodes = doc.querySelectorAll('a\\:p, p');
                textNodes.forEach(p => {
                    const textContent = Array.from(p.querySelectorAll('a\\:t, t')).map(t => t.textContent).join('');
                    if (textContent.trim()) {
                        parts.push(textContent.trim());
                    }
                });
            }

            // Append extracted images
            if (Object.keys(extractedImages).length > 0) {
                parts.push(`\n### Hình ảnh trong bài thuyết trình\n`);
                for (const imgName of Object.keys(extractedImages)) {
                    parts.push(`![${imgName}](${imgName})\n`);
                }
            }

            const fullText = parts.join('\n\n');
            return {
                markdown: fullText,
                markdown_preview: fullText,
                images: extractedImages,
                has_images: Object.keys(extractedImages).length > 0
            };
        }

        // 3. WORD CONVERTER (Mammoth.js)
        async function convertDOCX(buffer) {
            const extractedImages = {};
            let imgCounter = 0;

            const options = {
                convertImage: mammoth.images.inline(function(element) {
                    return element.read("base64").then(function(imageBuffer) {
                        imgCounter++;
                        const ext = element.contentType.split('/').pop() || 'png';
                        const saveName = `docx_img_${imgCounter}.${ext}`;
                        const mime = element.contentType;
                        const dataUri = `data:${mime};base64,${imageBuffer}`;
                        
                        extractedImages[saveName] = { base64: imageBuffer, mime: mime, data_uri: dataUri };
                        return { src: saveName };
                    });
                })
            };

            const result = await mammoth.convertToHtml({ arrayBuffer: buffer }, options);
            let html = result.value;

            // Convert basic HTML to Markdown
            html = html.replace(/<h1>(.*?)<\/h1>/gi, '# $1\n\n')
                       .replace(/<h2>(.*?)<\/h2>/gi, '## $1\n\n')
                       .replace(/<h3>(.*?)<\/h3>/gi, '### $1\n\n')
                       .replace(/<h4>(.*?)<\/h4>/gi, '#### $1\n\n')
                       .replace(/<p>(.*?)<\/p>/gi, '$1\n\n')
                       .replace(/<strong>(.*?)<\/strong>/gi, '**$1**')
                       .replace(/<em>(.*?)<\/em>/gi, '*$1*')
                       .replace(/<code>(.*?)<\/code>/gi, '`$1`')
                       .replace(/<a href="([^"]+)">(.*?)<\/a>/gi, '[$2]($1)')
                       .replace(/<img src="([^"]+)"\s*\/?>/gi, '![$1]($1)')
                       .replace(/<li>(.*?)<\/li>/gi, '- $1\n')
                       .replace(/<\/?(ol|ul)>/gi, '\n');

            return {
                markdown: html.trim(),
                markdown_preview: html.trim(),
                images: extractedImages,
                has_images: Object.keys(extractedImages).length > 0
            };
        }

        // 4. PDF CONVERTER (PDF.js)
        async function convertPDF(buffer) {
            const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
            const parts = [];

            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                const pageText = textContent.items.map(item => item.str).join(' ');
                if (pageText.trim()) {
                    parts.push(`## Trang ${i}\n\n${pageText.trim()}`);
                }
            }

            const fullText = parts.length ? parts.join('\n\n') : 'Không thể trích xuất văn bản từ PDF này (có thể là file scan dạng ảnh).';
            return {
                markdown: fullText,
                markdown_preview: fullText,
                images: {},
                has_images: false
            };
        }

        // 5. VBA MACROS & SOURCE CODE EXTRACTOR (MS-OVBA & OLE Stream Decompressor)
        function decompressMSOVBA(data) {
            if (!data || data.length < 3 || data[0] !== 0x01) return null;
            let pos = 1;
            let out = [];
            while (pos < data.length) {
                if (pos + 2 > data.length) break;
                const header = data[pos] | (data[pos + 1] << 8);
                pos += 2;
                const size = (header & 0x0FFF) + 3;
                const isCompressed = (header & 0x8000) !== 0;
                const chunkEnd = Math.min(pos + size - 2, data.length);

                if (!isCompressed) {
                    for (let i = pos; i < chunkEnd; i++) out.push(data[i]);
                    pos = chunkEnd;
                } else {
                    let chunkPos = pos;
                    while (chunkPos < chunkEnd) {
                        const flagByte = data[chunkPos++];
                        for (let bit = 0; bit < 8 && chunkPos < chunkEnd; bit++) {
                            if ((flagByte & (1 << bit)) === 0) {
                                out.push(data[chunkPos++]);
                            } else {
                                if (chunkPos + 2 > data.length) break;
                                const copyToken = data[chunkPos] | (data[chunkPos + 1] << 8);
                                chunkPos += 2;
                                let bitCount = Math.max(Math.ceil(Math.log2(out.length)), 4);
                                if (bitCount > 12) bitCount = 12;
                                const lengthMask = 0xFFFF >> bitCount;
                                const offsetMask = ~lengthMask;
                                const length = (copyToken & lengthMask) + 3;
                                const offset = ((copyToken & offsetMask) >> (16 - bitCount)) + 1;
                                const copyStart = out.length - offset;
                                for (let k = 0; k < length; k++) {
                                    if (copyStart + k >= 0 && copyStart + k < out.length) {
                                        out.push(out[copyStart + k]);
                                    }
                                }
                            }
                        }
                    }
                    pos = chunkEnd;
                }
            }
            return new TextDecoder('utf-8', { fatal: false }).decode(new Uint8Array(out));
        }

        function extractCleanVBAModules(vbaBuffer) {
            const extractedBlocks = [];
            const seenSignatures = new Set();

            for (let i = 0; i < vbaBuffer.length - 10; i++) {
                if (vbaBuffer[i] === 0x01 && (vbaBuffer[i + 2] & 0xF0) === 0xB0) {
                    const decomp = decompressMSOVBA(vbaBuffer.subarray(i));
                    if (!decomp) continue;

                    const text = decomp.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
                    const lines = text.split('\n');

                    let currentRoutine = [];
                    let inRoutine = false;

                    for (let line of lines) {
                        const trimmed = line.trim();
                        if (trimmed.startsWith('Attribute VB_')) continue;

                        if (/^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([a-zA-Z0-9_]+)/i.test(trimmed)) {
                            inRoutine = true;
                            currentRoutine = [line];
                        } else if (inRoutine) {
                            currentRoutine.push(line);
                            if (/^End\s+(?:Sub|Function|Property)/i.test(trimmed)) {
                                const block = currentRoutine.join('\n').trim();
                                let isClean = true;
                                for (let c = 0; c < block.length; c++) {
                                    const code = block.charCodeAt(c);
                                    if (code !== 9 && code !== 10 && code !== 13 && (code < 32 || (code > 126 && code < 160))) {
                                        isClean = false;
                                        break;
                                    }
                                }
                                if (isClean && !seenSignatures.has(block)) {
                                    seenSignatures.add(block);
                                    extractedBlocks.push(block);
                                }
                                inRoutine = false;
                                currentRoutine = [];
                            }
                        }
                    }
                }
            }

            return extractedBlocks;
        }

        async function extractVBA(buffer) {
            try {
                const zip = await JSZip.loadAsync(buffer);
                const vbaEntry = zip.file(/vbaProject\.bin$/i)[0];
                if (!vbaEntry) return '';

                const vbaBuffer = await vbaEntry.async('uint8array');
                const routines = extractCleanVBAModules(vbaBuffer);

                if (routines.length > 0) {
                    return `\n\n## Mã nguồn VBA Macros (Code đính kèm)\n\n\`\`\`vba\nOption Explicit\n\n${routines.join('\n\n')}\n\`\`\`\n`;
                } else {
                    return `\n\n## VBA Macros\n*Phát hiện module VBA Macros trong tệp Office này (vbaProject.bin).*\n`;
                }
            } catch (e) {
                console.error("VBA extraction error:", e);
            }
            return '';
        }

        // MAIN CONVERSION ROUTER
        async function convertDocumentNative(buffer, filename, excludeHidden) {
            const lower = filename.toLowerCase();
            let res;

            if (lower.endsWith('.xlsx') || lower.endsWith('.xlsm') || lower.endsWith('.xltx') || lower.endsWith('.xltm')) {
                res = await convertExcel(buffer, filename, excludeHidden);
            } else if (lower.endsWith('.pptx') || lower.endsWith('.pptm') || lower.endsWith('.potx')) {
                res = await convertPPTX(buffer, filename);
            } else if (lower.endsWith('.docx') || lower.endsWith('.docm')) {
                res = await convertDOCX(buffer);
            } else if (lower.endsWith('.pdf')) {
                res = await convertPDF(buffer);
            } else if (lower.endsWith('.csv') || lower.endsWith('.json') || lower.endsWith('.xml') || lower.endsWith('.html') || lower.endsWith('.txt') || lower.endsWith('.md')) {
                const text = new TextDecoder('utf-8').decode(buffer);
                res = { markdown: text, markdown_preview: text, images: {}, has_images: false };
            } else {
                throw new Error(`Định dạng file không được hỗ trợ: ${filename}`);
            }

            // VBA Check
            if (lower.match(/\.(xlsm|docm|pptm)$/)) {
                const vbaInfo = await extractVBA(buffer);
                if (vbaInfo) {
                    res.markdown += vbaInfo;
                    res.markdown_preview += vbaInfo;
                }
            }

            res.filename = filename;
            return res;
        }

        // Render Active Document
        function renderActiveDocument() {
            if (!currentResults.length) return;
            const doc = currentResults[currentActiveIndex];
            if (!doc) return;

            const isLimit = chkLimitRows.checked;
            const mdContent = isLimit ? (doc.markdown_preview || doc.markdown) : doc.markdown;

            codeArea.value = mdContent;

            // Substitute image filenames with Data URIs in HTML preview
            let previewMd = mdContent;
            if (doc.images) {
                for (const [imgName, imgMeta] of Object.entries(doc.images)) {
                    if (imgMeta && imgMeta.data_uri) {
                        const escapedName = imgName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                        previewMd = previewMd.replace(new RegExp(`src=['"]${escapedName}['"]`, 'g'), `src="${imgMeta.data_uri}"`);
                        previewMd = previewMd.replace(new RegExp(`\\]\\(${escapedName}\\)`, 'g'), `](${imgMeta.data_uri})`);
                    }
                }
            }

            previewArea.innerHTML = marked.parse(previewMd);
        }

        chkLimitRows.addEventListener('change', renderActiveDocument);

        function selectFile(index) {
            currentActiveIndex = index;
            document.querySelectorAll('.file-item').forEach((el, i) => {
                if (i === index) {
                    el.className = 'file-item cursor-pointer p-2.5 rounded-xl text-xs truncate border transition-all bg-violet-600/30 border-violet-500/60 text-white font-semibold shadow-md shadow-violet-600/20';
                } else {
                    el.className = 'file-item cursor-pointer p-2.5 rounded-xl text-xs truncate border transition-all border-transparent text-slate-400 hover:bg-slate-800/50 hover:text-slate-200';
                }
            });
            renderActiveDocument();
        }

        function handleConversionComplete(results, isBatch) {
            currentResults = results;
            currentActiveIndex = 0;

            stateIcon.innerHTML = '<i data-lucide="check-circle" class="w-5 h-5 text-emerald-400"></i>';
            stateTitle.textContent = 'Hoàn tất chuyển đổi!';
            stateTitle.className = 'text-xs font-semibold text-emerald-400';
            stateDesc.textContent = isBatch 
                ? `Đã xử lý xong ${results.length} files.`
                : `Đã chuyển đổi thành công: ${results[0].filename}`;

            dropzone.classList.remove('pointer-events-none', 'opacity-70');
            emptyState.classList.add('hidden');
            btnCopy.disabled = false;
            btnDownload.disabled = false;

            if (isBatch) {
                fileSidebar.classList.remove('hidden');
                batchCountBadge.textContent = results.length;
                fileList.innerHTML = '';
                results.forEach((res, idx) => {
                    const item = document.createElement('div');
                    item.className = 'file-item cursor-pointer p-2.5 rounded-xl text-xs truncate border transition-all border-transparent text-slate-400 hover:bg-slate-800/50 hover:text-slate-200';
                    item.title = res.filename;
                    item.textContent = res.filename;
                    item.onclick = () => selectFile(idx);
                    fileList.appendChild(item);
                });
                selectFile(0);
                btnDownloadText.textContent = 'Tải tất cả (ZIP)';
            } else {
                fileSidebar.classList.add('hidden');
                const single = results[0];
                btnDownloadText.textContent = single.has_images ? 'Tải ZIP (kèm ảnh)' : 'Tải .md';
                renderActiveDocument();
            }

            if (btnPreview.classList.contains('bg-violet-600')) {
                previewArea.classList.remove('hidden');
                codeArea.classList.add('hidden');
            } else {
                codeArea.classList.remove('hidden');
                previewArea.classList.add('hidden');
            }

            lucide.createIcons();
        }

        function handleConversionError(errorMsg) {
            stateIcon.innerHTML = '<i data-lucide="alert-triangle" class="w-5 h-5 text-rose-400"></i>';
            stateTitle.textContent = 'Lỗi chuyển đổi!';
            stateTitle.className = 'text-xs font-semibold text-rose-400';
            stateDesc.textContent = errorMsg;
            dropzone.classList.remove('pointer-events-none', 'opacity-70');
            lucide.createIcons();
        }

        // Tab Switching
        btnPreview.addEventListener('click', () => {
            btnPreview.className = 'px-3 py-1 text-xs font-medium rounded-md bg-violet-600 text-white transition-all shadow-sm';
            btnCode.className = 'px-3 py-1 text-xs font-medium rounded-md text-slate-400 hover:text-slate-200 transition-all';
            previewArea.classList.remove('hidden');
            codeArea.classList.add('hidden');
        });

        btnCode.addEventListener('click', () => {
            btnCode.className = 'px-3 py-1 text-xs font-medium rounded-md bg-violet-600 text-white transition-all shadow-sm';
            btnPreview.className = 'px-3 py-1 text-xs font-medium rounded-md text-slate-400 hover:text-slate-200 transition-all';
            codeArea.classList.remove('hidden');
            previewArea.classList.add('hidden');
        });

        // Copy button
        btnCopy.addEventListener('click', () => {
            if (!codeArea.value) return;
            navigator.clipboard.writeText(codeArea.value);
            const originalContent = btnCopy.innerHTML;
            btnCopy.innerHTML = '<i data-lucide="check" class="w-3.5 h-3.5 text-emerald-400"></i> <span>Đã sao chép!</span>';
            lucide.createIcons();
            setTimeout(() => {
                btnCopy.innerHTML = originalContent;
                lucide.createIcons();
            }, 2000);
        });

        // Native Download Helper
        function downloadBlob(blob, filename) {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        // Client-side ZIP & MD Generation
        btnDownload.addEventListener('click', async () => {
            if (!currentResults.length) return;

            const isBatch = currentResults.length > 1;
            const isLimit = chkLimitRows.checked;

            if (isBatch) {
                const zip = new JSZip();
                for (const doc of currentResults) {
                    const baseName = doc.filename.replace(/\.[^/.]+$/, "");
                    const mdContent = isLimit ? (doc.markdown_preview || doc.markdown) : doc.markdown;
                    zip.file(`${baseName}.md`, mdContent);

                    if (doc.images) {
                        for (const [imgName, imgMeta] of Object.entries(doc.images)) {
                            if (imgMeta && imgMeta.base64) {
                                zip.file(imgName, imgMeta.base64, { base64: true });
                            }
                        }
                    }
                }
                const content = await zip.generateAsync({ type: 'blob' });
                downloadBlob(content, 'documents_markdown.zip');
            } else {
                const doc = currentResults[0];
                const baseName = doc.filename.replace(/\.[^/.]+$/, "");
                const mdContent = isLimit ? (doc.markdown_preview || doc.markdown) : doc.markdown;

                if (doc.has_images) {
                    const zip = new JSZip();
                    zip.file(`${baseName}.md`, mdContent);
                    for (const [imgName, imgMeta] of Object.entries(doc.images)) {
                        if (imgMeta && imgMeta.base64) {
                            zip.file(imgName, imgMeta.base64, { base64: true });
                        }
                    }
                    const content = await zip.generateAsync({ type: 'blob' });
                    downloadBlob(content, `${baseName}.zip`);
                } else {
                    const blob = new Blob([mdContent], { type: 'text/markdown;charset=utf-8' });
                    downloadBlob(blob, `${baseName}.md`);
                }
            }
        });

        // File Handler - Processes files instantly
        async function processFiles(files) {
            if (!files.length) return;

            stateIcon.innerHTML = '<div class="w-5 h-5 border-2 border-violet-500 border-t-transparent rounded-full animate-spin"></div>';
            stateTitle.textContent = 'Đang xử lý...';
            stateTitle.className = 'text-xs font-semibold text-violet-300';
            dropzone.classList.add('pointer-events-none', 'opacity-70');

            const results = [];
            const excludeHidden = chkExcludeHidden.checked;

            try {
                for (let i = 0; i < files.length; i++) {
                    const file = files[i];
                    stateDesc.textContent = `Đang chuyển đổi ${file.name} (${i + 1}/${files.length})...`;
                    
                    const buffer = await file.arrayBuffer();
                    const result = await convertDocumentNative(buffer, file.name, excludeHidden);
                    results.push(result);
                }

                handleConversionComplete(results, files.length > 1);

            } catch (err) {
                console.error("Conversion error:", err);
                handleConversionError(err.message || 'Lỗi xử lý file.');
            }
        }

        // Drag and drop setup
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => {
            document.body.addEventListener(evt, e => {
                e.preventDefault();
                e.stopPropagation();
            });
        });

        ['dragenter', 'dragover'].forEach(evt => {
            document.body.addEventListener(evt, () => dropzone.classList.add('drag-active', 'scale-[1.01]'));
        });

        ['dragleave', 'drop'].forEach(evt => {
            document.body.addEventListener(evt, () => dropzone.classList.remove('drag-active', 'scale-[1.01]'));
        });

        document.body.addEventListener('drop', async (e) => {
            const items = e.dataTransfer.items;
            if (items && items.length) {
                const files = await getFilesFromDataTransferItems(items);
                if (files.length) processFiles(files);
            } else if (e.dataTransfer.files.length) {
                processFiles(Array.from(e.dataTransfer.files));
            }
        });

        async function getFilesFromDataTransferItems(items) {
            const files = [];
            async function traverse(item) {
                if (!item) return;
                if (item.isFile) {
                    const f = await new Promise(res => item.file(res));
                    files.push(f);
                } else if (item.isDirectory) {
                    const reader = item.createReader();
                    const entries = await new Promise(res => reader.readEntries(res));
                    for (const entry of entries) {
                        await traverse(entry);
                    }
                }
            }
            for (let i = 0; i < items.length; i++) {
                const entry = items[i].webkitGetAsEntry ? items[i].webkitGetAsEntry() : null;
                if (entry) await traverse(entry);
            }
            return files;
        }

        fileInput.addEventListener('change', function() {
            if (this.files.length) {
                processFiles(Array.from(this.files));
                this.value = '';
            }
        });

        folderInput.addEventListener('change', function() {
            if (this.files.length) {
                processFiles(Array.from(this.files));
                this.value = '';
            }
        });
