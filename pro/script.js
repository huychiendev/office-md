
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

        // --- DRAWINGML FLOWCHART & MERMAID GRAPH ENGINE ---
        function getXmlElements(parent, tagName) {
            if (!parent) return [];
            const plainTag = tagName.includes(':') ? tagName.split(':')[1] : tagName;
            const prefixTag = tagName.includes(':') ? tagName : `xdr:${tagName}`;
            const res1 = parent.getElementsByTagName ? Array.from(parent.getElementsByTagName(prefixTag)) : [];
            const res2 = parent.getElementsByTagName ? Array.from(parent.getElementsByTagName(plainTag)) : [];
            const res3 = parent.getElementsByTagName ? Array.from(parent.getElementsByTagName(`a:${plainTag}`)) : [];
            const set = new Set([...res1, ...res2, ...res3]);
            return Array.from(set);
        }

        function getXmlFirst(parent, tagName) {
            const list = getXmlElements(parent, tagName);
            return list.length > 0 ? list[0] : null;
        }

        function getXmlText(parent, tagName) {
            const el = getXmlFirst(parent, tagName);
            return el ? (el.textContent || '').trim() : '';
        }

        async function extractDrawingFlowcharts(zip, workbook) {
            const flowcharts = {};
            if (!zip || !workbook || typeof DOMParser === 'undefined') return flowcharts;

            const parser = new DOMParser();

            try {
                const wbXmlStr = await zip.file('xl/workbook.xml')?.async('string');
                const wbRelsStr = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
                if (!wbXmlStr) return flowcharts;

                const wbDoc = parser.parseFromString(wbXmlStr, 'text/xml');
                const wbRelsDoc = wbRelsStr ? parser.parseFromString(wbRelsStr, 'text/xml') : null;

                const relIdToTarget = {};
                if (wbRelsDoc) {
                    const rels = wbRelsDoc.getElementsByTagName('Relationship');
                    for (let i = 0; i < rels.length; i++) {
                        const r = rels[i];
                        const id = r.getAttribute('Id');
                        const target = r.getAttribute('Target') || '';
                        relIdToTarget[id] = 'xl/' + target.replace(/^\.\.\//, '').replace(/^\//, '');
                    }
                }

                const sheetIdToFile = {};
                const sheetElements = getXmlElements(wbDoc, 'sheet');
                for (let i = 0; i < sheetElements.length; i++) {
                    const s = sheetElements[i];
                    const name = s.getAttribute('name');
                    const sheetId = s.getAttribute('sheetId') || String(i + 1);
                    const rId = s.getAttribute('r:id') || s.getAttribute('id') || s.getAttribute('relationships:id');
                    if (rId && relIdToTarget[rId]) {
                        sheetIdToFile[sheetId] = relIdToTarget[rId];
                        sheetIdToFile[name] = relIdToTarget[rId];
                    } else {
                        sheetIdToFile[sheetId] = `xl/worksheets/sheet${sheetId}.xml`;
                        sheetIdToFile[name] = `xl/worksheets/sheet${sheetId}.xml`;
                    }
                }

                for (const worksheet of workbook.worksheets) {
                    const sheetFile = sheetIdToFile[String(worksheet.id)] || sheetIdToFile[worksheet.name] || `xl/worksheets/sheet${worksheet.id}.xml`;
                    const sheetRelsFile = sheetFile.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
                    const sheetRelsStr = await zip.file(sheetRelsFile)?.async('string');
                    if (!sheetRelsStr) continue;

                    const sheetRelsDoc = parser.parseFromString(sheetRelsStr, 'text/xml');
                    const sheetRels = sheetRelsDoc.getElementsByTagName('Relationship');
                    let drawingFile = null;
                    for (let i = 0; i < sheetRels.length; i++) {
                        const r = sheetRels[i];
                        const type = r.getAttribute('Type') || '';
                        if (type.includes('drawing')) {
                            const target = r.getAttribute('Target') || '';
                            drawingFile = 'xl/' + target.replace(/^\.\.\//, '').replace(/^\//, '');
                            break;
                        }
                    }

                    if (!drawingFile || !zip.file(drawingFile)) continue;

                    const drawingXmlStr = await zip.file(drawingFile).async('string');
                    const dDoc = parser.parseFromString(drawingXmlStr, 'text/xml');

                    function parseAnchor(node) {
                        if (!node) return null;
                        const col = parseInt(getXmlText(node, 'col') || '0', 10);
                        const row = parseInt(getXmlText(node, 'row') || '0', 10);
                        const colOff = parseInt(getXmlText(node, 'colOff') || '0', 10);
                        const rowOff = parseInt(getXmlText(node, 'rowOff') || '0', 10);
                        return { col, row, colOff, rowOff };
                    }

                    const shapes = [];
                    const connectors = [];

                    const twoCellAnchors = getXmlElements(dDoc, 'twoCellAnchor');
                    const oneCellAnchors = getXmlElements(dDoc, 'oneCellAnchor');
                    const allAnchors = [...twoCellAnchors, ...oneCellAnchors];

                    for (const anc of allAnchors) {
                        const sp = getXmlFirst(anc, 'sp');
                        const cxnSp = getXmlFirst(anc, 'cxnSp');
                        const fromAnc = parseAnchor(getXmlFirst(anc, 'from'));
                        const toAnc = parseAnchor(getXmlFirst(anc, 'to'));

                        if (sp) {
                            const cNvPr = getXmlFirst(sp, 'cNvPr');
                            const textElements = getXmlElements(sp, 't');
                            const textLines = textElements.map(t => (t.textContent || '').trim()).filter(Boolean);
                            const textStr = textLines.join('\n').trim();
                            const geomEl = getXmlFirst(sp, 'prstGeom');
                            const geom = geomEl ? (geomEl.getAttribute('prst') || 'rect') : 'rect';

                            shapes.push({
                                id: cNvPr ? cNvPr.getAttribute('id') : '',
                                name: cNvPr ? cNvPr.getAttribute('name') : '',
                                text: textStr,
                                geom: geom,
                                from: fromAnc,
                                to: toAnc
                            });
                        } else if (cxnSp) {
                            const cNvPr = getXmlFirst(cxnSp, 'cNvPr');
                            const stCxn = getXmlFirst(cxnSp, 'stCxn');
                            const endCxn = getXmlFirst(cxnSp, 'endCxn');
                            const ln = getXmlFirst(cxnSp, 'ln');
                            const tail = ln ? getXmlFirst(ln, 'tailEnd') : null;
                            const head = ln ? getXmlFirst(ln, 'headEnd') : null;
                            const xfrm = getXmlFirst(cxnSp, 'xfrm');
                            const geomEl = getXmlFirst(cxnSp, 'prstGeom');

                            connectors.push({
                                id: cNvPr ? cNvPr.getAttribute('id') : '',
                                name: cNvPr ? cNvPr.getAttribute('name') : '',
                                st_id: stCxn ? stCxn.getAttribute('id') : null,
                                end_id: endCxn ? endCxn.getAttribute('id') : null,
                                from: fromAnc,
                                to: toAnc,
                                geom: geomEl ? (geomEl.getAttribute('prst') || 'line') : 'line',
                                tail: tail ? (tail.getAttribute('type') || null) : null,
                                head: head ? (head.getAttribute('type') || null) : null,
                                rot: xfrm ? xfrm.getAttribute('rot') : null,
                                flipH: xfrm ? xfrm.getAttribute('flipH') : null,
                                flipV: xfrm ? xfrm.getAttribute('flipV') : null,
                            });
                        }
                    }

                    // Deduplicate connectors
                    const uniqueConnectors = [];
                    const seenCxn = new Set();
                    for (const c of connectors) {
                        if (!c.from || !c.to) continue;
                        if (c.from.col === c.to.col && c.from.row === c.to.row && c.from.colOff === c.to.colOff && c.from.rowOff === c.to.rowOff) {
                            continue;
                        }
                        const key = `${c.from.col}_${c.from.row}_${c.to.col}_${c.to.row}_${c.geom}_${c.tail}_${c.head}_${c.flipV}`;
                        if (!seenCxn.has(key)) {
                            seenCxn.add(key);
                            uniqueConnectors.push(c);
                        }
                    }

                    if (uniqueConnectors.length === 0 && shapes.length === 0) continue;

                    const shapesWithLongText = shapes.filter(s => s.text && s.text.length > 15);
                    let nodes = [];
                    let floatingLabels = [];

                    if (shapesWithLongText.length >= 3) {
                        // Pattern A: Shapes are nodes
                        shapesWithLongText.forEach((s, idx) => {
                            nodes.push({
                                id: `N${idx + 1}`,
                                shape_id: s.id,
                                label: s.text,
                                col: [s.from.col, s.to ? s.to.col : s.from.col],
                                row: [s.from.row, s.to ? s.to.row : s.from.row],
                                lane: 'Chung'
                            });
                        });
                        floatingLabels = shapes.filter(s => s.text && s.text.length <= 15);
                    } else {
                        // Pattern B: Grid cells are nodes, shapes are condition labels
                        let minDiagRow = 9999;
                        let maxDiagRow = 0;
                        for (const c of uniqueConnectors) {
                            minDiagRow = Math.min(minDiagRow, c.from.row, c.to.row);
                            maxDiagRow = Math.max(maxDiagRow, c.from.row, c.to.row);
                        }
                        for (const s of shapes) {
                            if (s.from && s.to) {
                                minDiagRow = Math.min(minDiagRow, s.from.row, s.to.row);
                                maxDiagRow = Math.max(maxDiagRow, s.from.row, s.to.row);
                            }
                        }

                        const diagRowStart = Math.max(1, minDiagRow - 4);
                        let diagRowEnd = Math.min(worksheet.rowCount || 50, maxDiagRow + 3);

                        for (let r = diagRowStart; r <= (worksheet.rowCount || 50); r++) {
                            const row = worksheet.getRow(r);
                            const c1Val = String(row.getCell(1).value || row.getCell(2).value || '');
                            if (/^(II\.|2\.|Bảng|Báo biểu)/i.test(c1Val.trim()) && r > minDiagRow) {
                                diagRowEnd = Math.min(diagRowEnd, r - 1);
                                break;
                            }
                        }

                        const rawBlocks = [];
                        for (let r = diagRowStart; r <= diagRowEnd; r++) {
                            const row = worksheet.getRow(r);
                            for (let c = 1; c <= Math.min(worksheet.columnCount || 35, 35); c++) {
                                const cell = row.getCell(c);
                                if (cell.value !== null && cell.value !== undefined) {
                                    if (cell.isMerged && cell.master && cell.address !== cell.master.address) {
                                        continue;
                                    }
                                    let text = formatCellValue(cell).trim();
                                    if (!text) continue;

                                    let minR = r - 1, maxR = r - 1, minC = c - 1, maxC = c - 1;
                                    if (cell.isMerged && cell.master) {
                                        minR = cell.master.row - 1;
                                        minC = cell.master.col - 1;
                                        maxR = minR;
                                        maxC = minC;
                                    }

                                    rawBlocks.push({
                                        col_start: minC, col_end: maxC,
                                        row_start: minR, row_end: maxR,
                                        text: text
                                    });
                                }
                            }
                        }

                        const laneKeywords = ['qad', 'lab', 'xưởng', 'ppc', 'qc', 'kho', 'wms', 'kế hoạch', '车间', '采购', '部门'];
                        const lanes = [];
                        for (const b of rawBlocks) {
                            const tLower = b.text.toLowerCase();
                            const isLane = laneKeywords.some(kw => tLower.includes(kw)) && b.text.length < 25 && (b.row_start <= 8 || b.col_start <= 2);
                            if (isLane) {
                                lanes.push({
                                    id: `L_${lanes.length + 1}`,
                                    name: b.text,
                                    col_start: b.col_start, col_end: b.col_end,
                                    row_start: b.row_start, row_end: b.row_end
                                });
                            }
                        }

                        const processBlocks = rawBlocks.filter(b => 
                            !lanes.some(l => l.col_start === b.col_start && l.row_start === b.row_start) &&
                            !/lưu trình|流程|irp/i.test(b.text) &&
                            b.text.length > 1
                        );

                        const mergedNodes = [];
                        const usedBlocks = new Set();
                        for (let i = 0; i < processBlocks.length; i++) {
                            if (usedBlocks.has(i)) continue;
                            const b1 = processBlocks[i];
                            let pairIdx = null;
                            for (let j = 0; j < processBlocks.length; j++) {
                                if (j !== i && !usedBlocks.has(j)) {
                                    const b2 = processBlocks[j];
                                    if (b2.col_start === b1.col_start && Math.abs(b2.row_start - b1.row_end) <= 4) {
                                        pairIdx = j;
                                        break;
                                    }
                                }
                            }

                            if (pairIdx !== null) {
                                const b2 = processBlocks[pairIdx];
                                usedBlocks.add(i);
                                usedBlocks.add(pairIdx);
                                mergedNodes.push({
                                    text: `${b1.text}\n${b2.text}`,
                                    col: [Math.min(b1.col_start, b2.col_start), Math.max(b1.col_end, b2.col_end)],
                                    row: [Math.min(b1.row_start, b2.row_start), Math.max(b1.row_end, b2.row_end)]
                                });
                            } else {
                                usedBlocks.add(i);
                                mergedNodes.push({
                                    text: b1.text,
                                    col: [b1.col_start, b1.col_end],
                                    row: [b1.row_start, b1.row_end]
                                });
                            }
                        }

                        mergedNodes.forEach((mn, idx) => {
                            let assignedLane = 'Chung';
                            if (mn.col[0] <= 5) {
                                assignedLane = mn.row[0] <= 15 ? 'QAD' : 'LAB';
                            } else if (mn.col[0] >= 24) {
                                assignedLane = 'PPC';
                            } else if (mn.row[0] >= 18 && mn.row[0] <= 34 && mn.col[0] <= 18) {
                                assignedLane = 'LAB';
                            } else if (mn.row[0] <= 15) {
                                assignedLane = 'Xưởng';
                            }

                            nodes.push({
                                id: `N${idx + 1}`,
                                label: mn.text,
                                col: mn.col,
                                row: mn.row,
                                lane: assignedLane
                            });
                        });

                        floatingLabels = shapes;
                    }

                    function findClosestNode(col, row) {
                        let best = null;
                        let bestDist = 9999;
                        for (const n of nodes) {
                            const [cMin, cMax] = n.col;
                            const [rMin, rMax] = n.row;
                            if (col >= cMin && col <= cMax + 1 && row >= rMin && row <= rMax + 1) {
                                return n;
                            }
                            const dist = Math.min(Math.abs(col - cMin), Math.abs(col - cMax)) + Math.min(Math.abs(row - rMin), Math.abs(row - rMax));
                            if (dist < bestDist && dist <= 4) {
                                bestDist = dist;
                                best = n;
                            }
                        }
                        return best;
                    }

                    const edges = [];
                    const seenEdges = new Set();

                    for (const c of uniqueConnectors) {
                        const sNode = findClosestNode(c.from.col, c.from.row);
                        const tNode = findClosestNode(c.to.col, c.to.row);
                        if (sNode && tNode && sNode.id !== tNode.id) {
                            let condLabel = '';
                            for (const fl of floatingLabels) {
                                const flC = (fl.from.col + (fl.to ? fl.to.col : fl.from.col)) / 2;
                                const flR = (fl.from.row + (fl.to ? fl.to.row : fl.from.row)) / 2;
                                const minC = Math.min(c.from.col, c.to.col) - 2;
                                const maxC = Math.max(c.from.col, c.to.col) + 2;
                                const minR = Math.min(c.from.row, c.to.row) - 2;
                                const maxR = Math.max(c.from.row, c.to.row) + 2;
                                if (flC >= minC && flC <= maxC && flR >= minR && flR <= maxR) {
                                    condLabel = fl.text.replace(/\n/g, ' / ');
                                    break;
                                }
                            }

                            const edgeKey = `${sNode.id}_${tNode.id}_${condLabel}`;
                            if (!seenEdges.has(edgeKey)) {
                                seenEdges.add(edgeKey);
                                edges.push({
                                    from: sNode.id,
                                    to: tNode.id,
                                    label: condLabel,
                                    type: c.geom
                                });
                            }
                        }
                    }

                    const connectedIds = new Set();
                    for (const e of edges) {
                        connectedIds.add(e.from);
                        connectedIds.add(e.to);
                    }

                    const finalNodes = nodes.filter(n => connectedIds.has(n.id));
                    if (finalNodes.length === 0) continue;

                    // Generate Mermaid Diagram
                    const mermaidLines = ['```mermaid', 'flowchart TD'];
                    mermaidLines.push('    classDef laneQAD fill:#e0f2fe,stroke:#0284c7,stroke-width:2px,color:#0f172a;');
                    mermaidLines.push('    classDef laneXuong fill:#fef3c7,stroke:#d97706,stroke-width:2px,color:#0f172a;');
                    mermaidLines.push('    classDef laneLAB fill:#f3e8ff,stroke:#9333ea,stroke-width:2px,color:#0f172a;');
                    mermaidLines.push('    classDef lanePPC fill:#dcfce7,stroke:#16a34a,stroke-width:2px,color:#0f172a;');
                    mermaidLines.push('    classDef nodeEnd fill:#fee2e2,stroke:#dc2626,stroke-width:2px,color:#991b1b;\n');

                    const laneGroups = {};
                    for (const n of finalNodes) {
                        if (!laneGroups[n.lane]) laneGroups[n.lane] = [];
                        laneGroups[n.lane].push(n);
                    }

                    for (const [lName, lNodes] of Object.entries(laneGroups)) {
                        const safeLaneId = lName.replace(/[^a-zA-Z0-9_]/g, '_') + '_Lane';
                        mermaidLines.push(`    subgraph ${safeLaneId}["Bộ phận ${lName}"]`);
                        for (const n of lNodes) {
                            const cleanTxt = n.label.replace(/"/g, "'").replace(/\n/g, '<br/>');
                            let cls = 'laneQAD';
                            const lowerLane = lName.toLowerCase();
                            if (lowerLane.includes('xưởng') || lowerLane.includes('车间')) cls = 'laneXuong';
                            else if (lowerLane.includes('lab')) cls = 'laneLAB';
                            else if (lowerLane.includes('ppc')) cls = 'lanePPC';
                            if (cleanTxt.toLowerCase().includes('kết thúc') || cleanTxt.toLowerCase().includes('结尾')) cls = 'nodeEnd';
                            mermaidLines.push(`        ${n.id}["${cleanTxt}"]:::${cls}`);
                        }
                        mermaidLines.push('    end\n');
                    }

                    for (const e of edges) {
                        const lbl = e.label ? `|${e.label}|` : '';
                        mermaidLines.push(`    ${e.from} -->${lbl} ${e.to}`);
                    }

                    mermaidLines.push('```\n');

                    // Generate Structured Specification Table
                    const specLines = [
                        '\n### Chi tiết Các bước trong Quy trình (Process Steps Specification)\n',
                        '| Mã bước | Tên bước nghiệp vụ & Thao tác | Bộ phận phụ trách | Tọa độ ô |',
                        '| :--- | :--- | :--- | :--- |'
                    ];

                    for (const n of finalNodes) {
                        const cleanTxt = n.label.replace(/\n/g, ' <br/> ');
                        const coordStr = `C${n.col[0] + 1}..C${n.col[1] + 1}, R${n.row[0] + 1}..R${n.row[1] + 1}`;
                        specLines.push(`| **${n.id}** | ${cleanTxt} | **${n.lane}** | \`${coordStr}\` |`);
                    }

                    specLines.push('\n### Ma trận Điều kiện Rẽ nhánh & Luân chuyển (Decision & Transition Matrix)\n');
                    specLines.push('| Từ bước (Source) | Điều kiện / Nhãn rẽ nhánh | Tới bước (Target) |');
                    specLines.push('| :--- | :--- | :--- |');

                    for (const e of edges) {
                        const sNode = finalNodes.find(n => n.id === e.from);
                        const tNode = finalNodes.find(n => n.id === e.to);
                        const sName = sNode ? sNode.label.split('\n')[0] : e.from;
                        const tName = tNode ? tNode.label.split('\n')[0] : e.to;
                        const lbl = e.label ? `\`${e.label}\`` : '*(Luồng mặc định)*';
                        specLines.push(`| **${e.from}** (${sName}) | ${lbl} | **${e.to}** (${tName}) |`);
                    }

                    const fullMarkdown = mermaidLines.join('\n') + '\n' + specLines.join('\n') + '\n';
                    flowcharts[worksheet.id] = fullMarkdown;
                    flowcharts[worksheet.name] = fullMarkdown;
                }
            } catch (e) {
                console.warn('Error extracting DrawingML flowcharts:', e);
            }

            return flowcharts;
        }

        // 1. EXCEL CONVERTER (ExcelJS + JSZip Hybrid Engine)
        async function convertExcel(buffer, filename, excludeHidden) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);

            const zip = await JSZip.loadAsync(buffer);
            const sheetFlowcharts = await extractDrawingFlowcharts(zip, workbook);

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

                // Insert Flowchart / Mermaid diagram if available
                const flowchartMd = sheetFlowcharts[worksheet.id] || sheetFlowcharts[worksheet.name];
                if (flowchartMd) {
                    parts.push(flowchartMd);
                    partsPreview.push(flowchartMd);
                }

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
            if (window.mermaid) {
                try {
                    window.mermaid.run({ nodes: previewArea.querySelectorAll('.language-mermaid') });
                } catch (e) {
                    console.warn('Mermaid render:', e);
                }
            }
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
