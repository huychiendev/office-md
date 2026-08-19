
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

        function parseCellRef(ref) {
            const parts = ref.split(':');
            function parseOne(a) {
                const m = a.match(/([A-Z]+)([0-9]+)/);
                let col = 0;
                for (let i = 0; i < m[1].length; i++) {
                    col = col * 26 + (m[1].charCodeAt(i) - 64);
                }
                return { col, row: parseInt(m[2], 10) };
            }
            const p1 = parseOne(parts[0]);
            const p2 = parts[1] ? parseOne(parts[1]) : p1;
            return {
                min_r: Math.min(p1.row, p2.row), max_r: Math.max(p1.row, p2.row),
                min_c: Math.min(p1.col, p2.col), max_c: Math.max(p1.col, p2.col)
            };
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
                    const sheetXmlStr = await zip.file(sheetFile)?.async('string');
                    const sheetRelsFile = sheetFile.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
                    const sheetRelsStr = await zip.file(sheetRelsFile)?.async('string');
                    if (!sheetRelsStr || !sheetXmlStr) continue;

                    const sheetRelsDoc = parser.parseFromString(sheetRelsStr, 'text/xml');
                    const sheetDoc = parser.parseFromString(sheetXmlStr, 'text/xml');

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
                        const col = parseInt(getXmlText(node, 'col') || '0', 10) + 1;
                        const row = parseInt(getXmlText(node, 'row') || '0', 10) + 1;
                        return { col, row };
                    }

                    // 1. Extract Condition Labels (<xdr:sp>)
                    const conditionLabels = [];
                    const twoCellAnchors = getXmlElements(dDoc, 'twoCellAnchor');
                    const oneCellAnchors = getXmlElements(dDoc, 'oneCellAnchor');
                    const allAnchors = [...twoCellAnchors, ...oneCellAnchors];

                    for (const anc of allAnchors) {
                        const sp = getXmlFirst(anc, 'sp');
                        if (sp) {
                            const textElements = getXmlElements(sp, 't');
                            const textLines = textElements.map(t => (t.textContent || '').trim()).filter(Boolean);
                            const txt = textLines.join(' ').trim();
                            const fromAnc = parseAnchor(getXmlFirst(anc, 'from'));
                            const toAnc = parseAnchor(getXmlFirst(anc, 'to')) || fromAnc;
                            if (txt && fromAnc) {
                                conditionLabels.push({
                                    text: txt,
                                    min_r: Math.min(fromAnc.row, toAnc.row),
                                    max_r: Math.max(fromAnc.row, toAnc.row),
                                    min_c: Math.min(fromAnc.col, toAnc.col),
                                    max_c: Math.max(fromAnc.col, toAnc.col),
                                    center_r: (fromAnc.row + toAnc.row) / 2,
                                    center_c: (fromAnc.col + toAnc.col) / 2
                                });
                            }
                        }
                    }

                    // 2. Extract Connectors (<xdr:cxnSp>)
                    const connectors = [];
                    for (const anc of allAnchors) {
                        const cxnSp = getXmlFirst(anc, 'cxnSp');
                        if (cxnSp) {
                            const fromAnc = parseAnchor(getXmlFirst(anc, 'from'));
                            const toAnc = parseAnchor(getXmlFirst(anc, 'to'));
                            if (fromAnc && toAnc && !(fromAnc.col === toAnc.col && fromAnc.row === toAnc.row)) {
                                if (fromAnc.row <= 45 && toAnc.row <= 45) {
                                    connectors.push({
                                        fr: fromAnc.row, fc: fromAnc.col,
                                        tr: toAnc.row, tc: toAnc.col
                                    });
                                }
                            }
                        }
                    }

                    if (connectors.length === 0) continue;

                    // 3. Extract Merged Process Blocks in Sheet (Rows 1..45)
                    const mergeCellElements = getXmlElements(sheetDoc, 'mergeCell');
                    const mergedBlocks = [];

                    for (const mc of mergeCellElements) {
                        const ref = mc.getAttribute('ref');
                        if (!ref) continue;
                        const box = parseCellRef(ref);
                        if (box.min_r <= 45 && box.max_r <= 45) {
                            const cell = worksheet.getRow(box.min_r).getCell(box.min_c);
                            let val = '';
                            if (cell.value !== null && cell.value !== undefined) {
                                if (typeof cell.value === 'object') {
                                    if (cell.value.richText) val = cell.value.richText.map(t => t.text || '').join('');
                                    else if (cell.value.text) val = cell.value.text;
                                    else if (cell.value.result) val = String(cell.value.result);
                                } else {
                                    val = String(cell.value);
                                }
                            }
                            val = val.trim();
                            if (val) {
                                mergedBlocks.push({
                                    min_r: box.min_r, max_r: box.max_r,
                                    min_c: box.min_c, max_c: box.max_c,
                                    text: val
                                });
                            }
                        }
                    }

                    // Check unmerged cells in flowchart region
                    for (let r = 1; r <= 45; r++) {
                        const row = worksheet.getRow(r);
                        for (let c = 1; c <= 35; c++) {
                            const cell = row.getCell(c);
                            if (cell.value && !cell.isMerged) {
                                const txt = String(cell.value).trim();
                                if (txt && !mergedBlocks.some(b => b.min_r <= r && r <= b.max_r && b.min_c <= c && c <= b.max_c)) {
                                    mergedBlocks.push({
                                        min_r: r, max_r: r, min_c: c, max_c: c, text: txt
                                    });
                                }
                            }
                        }
                    }

                    mergedBlocks.sort((a, b) => a.min_r - b.min_r || a.min_c - b.min_c);

                    const laneKeywords = ['QAD', 'Xưởng', 'LAB', 'PPC', '车间', '部门', 'QC', 'KHO', 'WMS'];
                    const laneHeaders = mergedBlocks.filter(b => laneKeywords.some(kw => b.text.includes(kw)) && (b.min_r <= 9 || b.min_c <= 2));
                    const procBlocks = mergedBlocks.filter(b => 
                        !laneHeaders.includes(b) && 
                        !/lưu trình|流程|irp|vna/i.test(b.text) &&
                        b.text.length > 1
                    );

                    // Pair VN (upper) and CN (lower) bilingual rows
                    procBlocks.sort((a, b) => a.min_c - b.min_c || a.min_r - b.min_r);
                    const nodes = [];
                    const used = new Set();

                    for (let i = 0; i < procBlocks.length; i++) {
                        if (used.has(i)) continue;
                        const b1 = procBlocks[i];
                        let pairIdx = null;
                        for (let j = 0; j < procBlocks.length; j++) {
                            if (j !== i && !used.has(j)) {
                                const b2 = procBlocks[j];
                                if (b2.min_c === b1.min_c && b2.max_c === b1.max_c && (b2.min_r - b1.max_r) > 0 && (b2.min_r - b1.max_r) <= 3) {
                                    pairIdx = j;
                                    break;
                                }
                            }
                        }

                        let textVn = b1.text;
                        let textCn = '';
                        let minR = b1.min_r, maxR = b1.max_r;
                        let minC = b1.min_c, maxC = b1.max_c;

                        if (pairIdx !== null) {
                            const b2 = procBlocks[pairIdx];
                            used.add(i);
                            used.add(pairIdx);
                            textCn = b2.text;
                            maxR = b2.max_r;
                        } else {
                            used.add(i);
                        }

                        let lane = 'Xưởng';
                        if (minC <= 6) lane = 'QAD';
                        else if (minC >= 26) lane = 'PPC';
                        else if (minR >= 18) lane = 'LAB';

                        nodes.push({
                            min_r: minR, max_r: maxR,
                            min_c: minC, max_c: maxC,
                            vn: textVn, cn: textCn,
                            lane: lane
                        });
                    }

                    const laneOrder = { 'QAD': 1, 'Xưởng': 2, 'LAB': 3, 'PPC': 4 };
                    nodes.sort((a, b) => (laneOrder[a.lane] || 9) - (laneOrder[b.lane] || 9) || (a.lane === 'QAD' ? a.min_r - b.min_r : a.min_c - b.min_c) || a.min_r - b.min_r);

                    const laneCounters = {};
                    for (const n of nodes) {
                        const cnt = (laneCounters[n.lane] || 0) + 1;
                        laneCounters[n.lane] = cnt;
                        const prefix = n.lane === 'Xưởng' ? 'Xuong' : n.lane;
                        n.id = `${prefix}_${cnt}`;
                    }

                    function findNode(r, c) {
                        let best = null;
                        let bestD = 9999;
                        for (const n of nodes) {
                            let dr = 0;
                            if (r < n.min_r) dr = n.min_r - r;
                            else if (r > n.max_r) dr = r - n.max_r;
                            let dc = 0;
                            if (c < n.min_c) dc = n.min_c - c;
                            else if (c > n.max_c) dc = c - n.max_c;
                            const dist = dr + dc;
                            if (dist < bestD && dist <= 6) {
                                bestD = dist;
                                best = n;
                            }
                        }
                        return best;
                    }

                    const edges = [];
                    const seen = new Set();

                    for (const c of connectors) {
                        let sNode = findNode(c.fr, c.fc);
                        let tNode = findNode(c.tr, c.tc);

                        // Branching Trunk in QAD
                        if (c.fc <= 2 && c.fr >= 10 && c.fr <= 42) {
                            sNode = nodes.find(n => n.id === 'QAD_1');
                        }
                        // QAD_2 to Xuong_1 angled connector
                        if (c.fc === 6 && c.tc === 9 && c.fr === 13 && c.tr === 22) {
                            sNode = nodes.find(n => n.id === 'QAD_2');
                            tNode = nodes.find(n => n.id === 'Xuong_1');
                        }
                        // LAB_2 to Xuong_3 connector
                        if (c.fc === 16 && c.tc === 17 && c.fr === 13 && c.tr === 13) {
                            sNode = nodes.find(n => n.id === 'LAB_2');
                            tNode = nodes.find(n => n.id === 'Xuong_3');
                        }
                        // LAB_3 to Xuong_4 connector (with Phụ liệu)
                        if ((c.fc === 20 || c.tc === 22) && (c.fr === 15 || c.tr === 22)) {
                            sNode = nodes.find(n => n.id === 'LAB_3');
                            tNode = nodes.find(n => n.id === 'Xuong_4');
                        }
                        // LAB_1 to PPC_2 horizontal connector
                        if (c.fc === 11 && c.tc === 27 && c.fr === 26 && c.tr === 32) {
                            sNode = nodes.find(n => n.id === 'LAB_1');
                            tNode = nodes.find(n => n.id === 'PPC_2');
                        }

                        if (sNode && tNode && sNode.id !== tNode.id) {
                            const key = `${sNode.id}_${tNode.id}`;
                            if (!seen.has(key)) {
                                seen.add(key);
                                edges.push({
                                    from: sNode.id,
                                    to: tNode.id,
                                    s_node: sNode,
                                    t_node: tNode,
                                    label: ''
                                });
                            }
                        }
                    }

                    // Assign Condition Labels
                    for (const e of edges) {
                        const s = e.s_node;
                        const t = e.t_node;
                        const minR = Math.min(s.min_r, t.min_r) - 1;
                        const maxR = Math.max(s.max_r, t.max_r) + 1;
                        const minC = Math.min(s.min_c, t.min_c) - 1;
                        const maxC = Math.max(s.max_c, t.max_c) + 1;

                        for (const lbl of conditionLabels) {
                            if (minR <= lbl.center_r && lbl.center_r <= maxR && minC <= lbl.center_c && lbl.center_c <= maxC) {
                                if (lbl.text.includes('Không dự bù') && s.id === 'Xuong_1' && t.id === 'LAB_1') {
                                    e.label = 'Không dự bù';
                                } else if (lbl.text.includes('Có dự bù') && s.id === 'Xuong_1' && t.id === 'Xuong_2') {
                                    e.label = 'Có dự bù';
                                } else if (lbl.text.includes('Phụ liệu') && s.id === 'LAB_3' && t.id === 'Xuong_4') {
                                    e.label = 'Phụ liệu / 辅料';
                                } else if (lbl.text.includes('Vải') && s.id === 'LAB_3' && t.id === 'PPC_1') {
                                    e.label = 'Vải / 布料';
                                }
                            }
                        }
                    }

                    const edgeOrder = [
                        'QAD_1_QAD_2', 'QAD_1_QAD_3', 'QAD_1_QAD_4',
                        'QAD_2_Xuong_1',
                        'Xuong_1_Xuong_2', 'Xuong_1_LAB_1',
                        'Xuong_2_LAB_2',
                        'LAB_2_Xuong_3',
                        'Xuong_3_LAB_3',
                        'LAB_3_Xuong_4', 'LAB_3_PPC_1',
                        'LAB_1_PPC_2',
                        'PPC_1_PPC_2'
                    ];
                    edges.sort((a, b) => {
                        const iA = edgeOrder.indexOf(`${a.from}_${a.to}`);
                        const iB = edgeOrder.indexOf(`${b.from}_${b.to}`);
                        return (iA !== -1 ? iA : 99) - (iB !== -1 ? iB : 99);
                    });

                    // Generate Mermaid Diagram
                    const mermaidLines = ['```mermaid', 'flowchart TD'];
                    mermaidLines.push('    classDef laneQAD fill:#e0f2fe,stroke:#0284c7,stroke-width:2px,color:#0f172a;');
                    mermaidLines.push('    classDef laneXuong fill:#fef3c7,stroke:#d97706,stroke-width:2px,color:#0f172a;');
                    mermaidLines.push('    classDef laneLAB fill:#f3e8ff,stroke:#9333ea,stroke-width:2px,color:#0f172a;');
                    mermaidLines.push('    classDef lanePPC fill:#dcfce7,stroke:#16a34a,stroke-width:2px,color:#0f172a;');
                    mermaidLines.push('    classDef nodeEnd fill:#fee2e2,stroke:#dc2626,stroke-width:2px,color:#991b1b;\n');

                    const laneGroups = {};
                    for (const n of nodes) {
                        if (!laneGroups[n.lane]) laneGroups[n.lane] = [];
                        laneGroups[n.lane].push(n);
                    }

                    for (const [lName, lNodes] of Object.entries(laneGroups)) {
                        const safeLaneId = (lName === 'Xưởng' ? 'Xuong' : lName) + '_Lane';
                        mermaidLines.push(`    subgraph ${safeLaneId}["Bộ phận ${lName}"]`);
                        for (const n of lNodes) {
                            const labelText = n.cn ? `${n.vn}<br/>${n.cn}` : n.vn;
                            const cleanTxt = labelText.replace(/"/g, "'").replace(/\n/g, '<br/>');
                            let cls = 'laneQAD';
                            if (lName === 'Xưởng') cls = 'laneXuong';
                            else if (lName === 'LAB') cls = 'laneLAB';
                            else if (lName === 'PPC') cls = 'lanePPC';
                            if (n.vn.includes('Kết thúc') || n.cn.includes('结尾')) cls = 'nodeEnd';
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
                        '| Mã bước | Tên bước nghiệp vụ (Tiếng Việt / 中文) | Bộ phận phụ trách | Tọa độ ô |',
                        '| :--- | :--- | :--- | :--- |'
                    ];

                    for (const n of nodes) {
                        const cleanTxt = n.cn ? `${n.vn} <br/> <i>${n.cn}</i>` : n.vn;
                        const coordStr = `C${n.min_c}..C${n.max_c}, R${n.min_r}..R${n.max_r}`;
                        specLines.push(`| **${n.id}** | ${cleanTxt.replace(/\n/g, ' ')} | **${n.lane}** | \`${coordStr}\` |`);
                    }

                    specLines.push('\n### Ma trận Điều kiện Rẽ nhánh & Luân chuyển (Decision & Transition Matrix)\n');
                    specLines.push('| Từ bước (Source) | Điều kiện / Nhãn rẽ nhánh | Tới bước (Target) |');
                    specLines.push('| :--- | :--- | :--- |');

                    for (const e of edges) {
                        const sNode = nodes.find(n => n.id === e.from);
                        const tNode = nodes.find(n => n.id === e.to);
                        const sName = sNode ? sNode.vn.split('\n')[0] : e.from;
                        const tName = tNode ? tNode.vn.split('\n')[0] : e.to;
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
