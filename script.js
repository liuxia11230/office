/**
 * Excel Viewer - 直接解析 Open XML 格式
 * 不依赖 ExcelJS 的样式解析，自己实现完整的样式读取
 */

document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const previewContainer = document.getElementById('preview-container');
    const tableWrapper = document.getElementById('table-wrapper');
    const controls = document.getElementById('controls');
    const sheetTabs = document.getElementById('sheet-tabs');
    const uploadWrapper = document.getElementById('upload-wrapper');
    const fileNameDisplay = document.getElementById('file-name');
    const cellInfo = document.getElementById('cell-info');

    // ==========================================
    // 样式存储
    // ==========================================
    
    let themeColors = {};      // 主题颜色
    let fonts = [];            // 字体定义
    let fills = [];            // 填充定义
    let borders = [];          // 边框定义
    let cellXfs = [];          // 单元格格式定义
    let sharedStrings = [];    // 共享字符串
    let numFmts = {};          // 数字格式
    let currentMaxCol = 0;     // 当前工作表最大列数（用于外侧边框处理）
let currentColWidths = {}; // 当前工作表列宽（用于溢出计算）
let currentMDW = 7;        // 当前默认字体的最大数字宽度（px）

    // 默认主题颜色
    const DEFAULT_THEME = {
        0: 'FFFFFF', 1: '000000', 2: 'E7E6E6', 3: '44546A',
        4: '4472C4', 5: 'ED7D31', 6: 'A5A5A5', 7: 'FFC000',
        8: '5B9BD5', 9: '70AD47', 10: '0563C1', 11: '954F72'
    };

    // ==========================================
    // 文件处理
    // ==========================================

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', e => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', e => {
        if (e.target.files.length) handleFile(e.target.files[0]);
    });

    async function handleFile(file) {
        fileNameDisplay.textContent = file.name;
        
        try {
            // 使用 JSZip 解压 xlsx 文件
            const zip = await JSZip.loadAsync(file);
            
            // 解析所有需要的 XML 文件
            await parseTheme(zip);
            await parseStyles(zip);
            await parseSharedStrings(zip);
            const workbook = await parseWorkbook(zip);
            
            console.log('解析完成:', {
                主题颜色: themeColors,
                字体数量: fonts.length,
                填充数量: fills.length,
                样式数量: cellXfs.length
            });
            
            processWorkbook(workbook, zip);
            
        } catch (err) {
            console.error('解析失败:', err);
            alert('无法解析文件: ' + err.message);
        }
    }

    // ==========================================
    // XML 解析函数
    // ==========================================

    async function parseTheme(zip) {
        themeColors = { ...DEFAULT_THEME };
        
        const themeFile = zip.file('xl/theme/theme1.xml');
        if (!themeFile) return;
        
        const xml = await themeFile.async('string');
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');
        
        const colorScheme = doc.querySelector('clrScheme');
        if (!colorScheme) return;
        
        const colorMap = {
            'dk1': 1, 'lt1': 0, 'dk2': 3, 'lt2': 2,
            'accent1': 4, 'accent2': 5, 'accent3': 6,
            'accent4': 7, 'accent5': 8, 'accent6': 9,
            'hlink': 10, 'folHlink': 11
        };
        
        for (const [name, index] of Object.entries(colorMap)) {
            const elem = colorScheme.querySelector(name);
            if (elem) {
                const srgb = elem.querySelector('srgbClr');
                const sys = elem.querySelector('sysClr');
                if (srgb) {
                    themeColors[index] = srgb.getAttribute('val');
                } else if (sys) {
                    themeColors[index] = sys.getAttribute('lastClr') || DEFAULT_THEME[index];
                }
            }
        }
        
        console.log('主题颜色:', themeColors);
    }

    async function parseStyles(zip) {
        fonts = [];
        fills = [];
        borders = [];
        cellXfs = [];
        numFmts = {};
        
        const stylesFile = zip.file('xl/styles.xml');
        if (!stylesFile) return;
        
        const xml = await stylesFile.async('string');
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');
        
        // 解析数字格式
        doc.querySelectorAll('numFmt').forEach(nf => {
            const id = nf.getAttribute('numFmtId');
            const code = nf.getAttribute('formatCode');
            if (id && code) numFmts[id] = code;
        });
        
        // 解析字体
        doc.querySelectorAll('fonts > font').forEach(fontElem => {
            const font = {};
            
            const name = fontElem.querySelector('name');
            if (name) font.name = name.getAttribute('val');
            
            const sz = fontElem.querySelector('sz');
            if (sz) font.size = parseFloat(sz.getAttribute('val'));
            
            const b = fontElem.querySelector('b');
            if (b) font.bold = b.getAttribute('val') !== 'false';
            
            const i = fontElem.querySelector('i');
            if (i) font.italic = i.getAttribute('val') !== 'false';
            
            const u = fontElem.querySelector('u');
            if (u) font.underline = true;
            
            const strike = fontElem.querySelector('strike');
            if (strike) font.strike = strike.getAttribute('val') !== 'false';
            
            const color = fontElem.querySelector('color');
            if (color) font.color = parseColor(color);
            
            fonts.push(font);
        });
        
        console.log('字体列表:');
        fonts.forEach((f, i) => {
            console.log(`  [${i}] ${f.name || '(默认)'} ${f.size || ''}pt`, 
                f.bold ? '粗体' : '', 
                f.italic ? '斜体' : '',
                f.underline ? '下划线' : '',
                f.color ? `颜色:${JSON.stringify(f.color)}` : '');
        });
        
        // 解析填充
        doc.querySelectorAll('fills > fill').forEach(fillElem => {
            const fill = {};
            const pattern = fillElem.querySelector('patternFill');
            
            if (pattern) {
                fill.type = 'pattern';
                fill.pattern = pattern.getAttribute('patternType') || 'none';
                
                const fgColor = pattern.querySelector('fgColor');
                if (fgColor) fill.fgColor = parseColor(fgColor);
                
                const bgColor = pattern.querySelector('bgColor');
                if (bgColor) fill.bgColor = parseColor(bgColor);
            }
            
            fills.push(fill);
        });
        
        console.log('填充列表:', fills);
        
        // 解析边框
        doc.querySelectorAll('borders > border').forEach(borderElem => {
            const border = {};
            
            ['left', 'right', 'top', 'bottom'].forEach(side => {
                const sideElem = borderElem.querySelector(side);
                if (sideElem) {
                    const style = sideElem.getAttribute('style');
                    if (style) {
                        border[side] = { style };
                        const color = sideElem.querySelector('color');
                        if (color) border[side].color = parseColor(color);
                    }
                }
            });
            
            borders.push(border);
        });
        
        // 解析 cellStyleXfs（基础样式）
        const baseCellXfs = [];
        doc.querySelectorAll('cellStyleXfs > xf').forEach(xf => {
            const style = {};
            const alignment = xf.querySelector('alignment');
            if (alignment) {
                const wrapAttr = alignment.getAttribute('wrapText');
                const shrinkAttr = alignment.getAttribute('shrinkToFit');
                style.alignment = {
                    horizontal: alignment.getAttribute('horizontal'),
                    vertical: alignment.getAttribute('vertical'),
                    wrapText: wrapAttr === 'true' || wrapAttr === '1',
                    shrinkToFit: shrinkAttr === 'true' || shrinkAttr === '1',
                    textRotation: parseInt(alignment.getAttribute('textRotation')) || 0,
                    indent: parseInt(alignment.getAttribute('indent')) || 0
                };
            }
            baseCellXfs.push(style);
        });

        // 解析 cellXfs（单元格样式组合）
        doc.querySelectorAll('cellXfs > xf').forEach(xf => {
            const style = {
                fontId: parseInt(xf.getAttribute('fontId')) || 0,
                fillId: parseInt(xf.getAttribute('fillId')) || 0,
                borderId: parseInt(xf.getAttribute('borderId')) || 0,
                numFmtId: parseInt(xf.getAttribute('numFmtId')) || 0,
                xfId: xf.getAttribute('xfId') !== null ? parseInt(xf.getAttribute('xfId')) : null,
                applyAlignment: xf.getAttribute('applyAlignment')
            };
            
            const alignment = xf.querySelector('alignment');
            if (alignment) {
                const wrapAttr = alignment.getAttribute('wrapText');
                const shrinkAttr = alignment.getAttribute('shrinkToFit');
                style.alignment = {
                    horizontal: alignment.getAttribute('horizontal'),
                    vertical: alignment.getAttribute('vertical'),
                    wrapText: wrapAttr === 'true' || wrapAttr === '1',
                    shrinkToFit: shrinkAttr === 'true' || shrinkAttr === '1',
                    textRotation: parseInt(alignment.getAttribute('textRotation')) || 0,
                    indent: parseInt(alignment.getAttribute('indent')) || 0
                };
            }
            
            cellXfs.push(style);
        });

        // 合并基础样式中的对齐（如果 cellXfs 中没有对齐信息）
        cellXfs.forEach(style => {
            const applyAlign = style.applyAlignment;
            if (applyAlign === '0') {
                style.alignment = null;
                return;
            }
            if (!style.alignment && style.xfId !== null && baseCellXfs[style.xfId]) {
                style.alignment = baseCellXfs[style.xfId].alignment || null;
            }
        });
        
        console.log('cellXfs 数量:', cellXfs.length);
    }

    function parseColor(colorElem) {
        if (!colorElem) return null;
        
        const rgb = colorElem.getAttribute('rgb');
        if (rgb) return { argb: rgb };
        
        const theme = colorElem.getAttribute('theme');
        if (theme !== null) {
            const tint = parseFloat(colorElem.getAttribute('tint')) || 0;
            return { theme: parseInt(theme), tint };
        }
        
        const indexed = colorElem.getAttribute('indexed');
        if (indexed !== null) return { indexed: parseInt(indexed) };
        
        return null;
    }

    async function parseSharedStrings(zip) {
        sharedStrings = [];
        
        const ssFile = zip.file('xl/sharedStrings.xml');
        if (!ssFile) return;
        
        const xml = await ssFile.async('string');
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');
        
        doc.querySelectorAll('si').forEach(si => {
            sharedStrings.push(extractRichText(si));
        });
    }

    async function parseWorkbook(zip) {
        const wbFile = zip.file('xl/workbook.xml');
        if (!wbFile) throw new Error('无效的 Excel 文件');
        
        const xml = await wbFile.async('string');
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');
        
        const sheets = [];
        doc.querySelectorAll('sheet').forEach(sheet => {
            sheets.push({
                name: sheet.getAttribute('name'),
                id: sheet.getAttribute('sheetId'),
                rId: sheet.getAttribute('r:id')
            });
        });
        
        return { sheets };
    }

    function extractRichText(node) {
        if (!node) return '';
        let text = '';
        node.childNodes.forEach(child => {
            if (child.nodeType !== 1) return;
            const tag = child.localName;
            if (tag === 't') {
                text += child.textContent || '';
            } else if (tag === 'r') {
                child.childNodes.forEach(rn => {
                    if (rn.nodeType !== 1) return;
                    const rtag = rn.localName;
                    if (rtag === 't') {
                        text += rn.textContent || '';
                    } else if (rtag === 'br') {
                        text += '\n';
                    }
                });
            } else if (tag === 'br') {
                text += '\n';
            }
        });
        if (!text) {
            const t = node.querySelector('t');
            if (t) text = t.textContent || '';
        }
        return text;
    }

    function parseHyperlinks(sheetDoc, relsDoc) {
        const map = new Map();
        if (!sheetDoc) return map;
        const relMap = new Map();
        if (relsDoc) {
            relsDoc.querySelectorAll('Relationship').forEach(rel => {
                const id = rel.getAttribute('Id');
                const target = rel.getAttribute('Target');
                if (id && target) relMap.set(id, target);
            });
        }

        sheetDoc.querySelectorAll('hyperlink').forEach(link => {
            const ref = link.getAttribute('ref');
            const rId = link.getAttribute('r:id') || link.getAttribute('id');
            const target = rId ? (relMap.get(rId) || '') : (link.getAttribute('location') || '');
            if (!ref) return;
            expandRangeRefs(ref).forEach(cellRefStr => {
                map.set(cellRefStr, { target });
            });
        });
        return map;
    }

    function expandRangeRefs(ref) {
        if (!ref) return [];
        if (!ref.includes(':')) return [ref];
        const [start, end] = ref.split(':');
        const s = cellRef(start);
        const e = cellRef(end);
        const refs = [];
        for (let r = s.r; r <= e.r; r++) {
            for (let c = s.c; c <= e.c; c++) {
                refs.push(`${colName(c)}${r}`);
            }
        }
        return refs;
    }

    function escapeHtml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function buildHyperlinkHtml(text, target) {
        const str = String(text || '');
        const match = str.match(/https?:\/\/\S+/);
        if (!match) return escapeHtml(str);
        const url = match[0];
        const href = target || url;
        const before = str.slice(0, match.index);
        const after = str.slice(match.index + url.length);
        return `${escapeHtml(before)}<a class="cell-hyperlink" href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(url)}</a>${escapeHtml(after)}`;
    }

    // ==========================================
    // 颜色解析
    // ==========================================

    function resolveColor(colorObj) {
        if (!colorObj) return null;
        
        // ARGB 格式
        if (colorObj.argb) {
            const argb = colorObj.argb;
            if (argb.length === 8) {
                return '#' + argb.substring(2);
            }
            return '#' + argb;
        }
        
        // 主题颜色
        if (colorObj.theme !== undefined) {
            let baseColor = themeColors[colorObj.theme] || DEFAULT_THEME[colorObj.theme];
            if (baseColor) {
                let color = '#' + baseColor;
                if (colorObj.tint && colorObj.tint !== 0) {
                    color = applyTint(color, colorObj.tint);
                }
                return color;
            }
        }
        
        // 索引颜色
        if (colorObj.indexed !== undefined) {
            const colors = [
                '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
                '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
                '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
                '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
                '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
                '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
                '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
                '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'
            ];
            return colors[colorObj.indexed] || null;
        }
        
        return null;
    }

    function applyTint(hexColor, tint) {
        const hex = hexColor.replace('#', '');
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        
        // RGB to HSL
        const rn = r / 255, gn = g / 255, bn = b / 255;
        const max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn);
        let h, s, l = (max + min) / 2;
        
        if (max === min) {
            h = s = 0;
        } else {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case rn: h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6; break;
                case gn: h = ((bn - rn) / d + 2) / 6; break;
                case bn: h = ((rn - gn) / d + 4) / 6; break;
            }
        }
        
        // Apply tint
        if (tint < 0) {
            l = l * (1 + tint);
        } else {
            l = l * (1 - tint) + tint;
        }
        l = Math.max(0, Math.min(1, l));
        
        // HSL to RGB
        let r2, g2, b2;
        if (s === 0) {
            r2 = g2 = b2 = l;
        } else {
            const hue2rgb = (p, q, t) => {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1/6) return p + (q - p) * 6 * t;
                if (t < 1/2) return q;
                if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
                return p;
            };
            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            r2 = hue2rgb(p, q, h + 1/3);
            g2 = hue2rgb(p, q, h);
            b2 = hue2rgb(p, q, h - 1/3);
        }
        
        const toHex = n => Math.round(n * 255).toString(16).padStart(2, '0');
        return `#${toHex(r2)}${toHex(g2)}${toHex(b2)}`;
    }

    // ==========================================
    // 工作簿处理
    // ==========================================

    function processWorkbook(workbook, zip) {
        uploadWrapper.classList.add('hide');
        controls.classList.remove('hide');
        previewContainer.classList.remove('hide');

        document.querySelectorAll('.info-bar').forEach(e => e.remove());
        
        setupTabs(workbook.sheets, zip);
        
        if (workbook.sheets.length > 0) {
            loadSheet(workbook.sheets[0], zip);
        }
    }

    function setupTabs(sheets, zip) {
        sheetTabs.innerHTML = '';

        sheets.forEach((sheet, index) => {
            const btn = document.createElement('button');
            btn.textContent = sheet.name;
            btn.className = `tab-btn ${index === 0 ? 'active' : ''}`;
            
            btn.addEventListener('click', () => {
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                loadSheet(sheet, zip);
            });
            
            sheetTabs.appendChild(btn);
        });
    }

    async function loadSheet(sheetInfo, zip) {
        // 找到工作表文件
        const relsFile = zip.file('xl/_rels/workbook.xml.rels');
        let sheetPath = `xl/worksheets/sheet${sheetInfo.id}.xml`;
        
        if (relsFile) {
            const relsXml = await relsFile.async('string');
            const parser = new DOMParser();
            const relsDoc = parser.parseFromString(relsXml, 'text/xml');
            
            const rel = relsDoc.querySelector(`Relationship[Id="${sheetInfo.rId}"]`);
            if (rel) {
                sheetPath = 'xl/' + rel.getAttribute('Target');
            }
        }
        
        const sheetFile = zip.file(sheetPath);
        if (!sheetFile) {
            tableWrapper.innerHTML = '<div style="padding:2rem;text-align:center;">无法加载工作表</div>';
            return;
        }

        // 读取工作表关系（用于超链接）
        let sheetRelsDoc = null;
        const relsPath = sheetPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
        const sheetRelsFile = zip.file(relsPath);
        if (sheetRelsFile) {
            const relsXml = await sheetRelsFile.async('string');
            const relsParser = new DOMParser();
            sheetRelsDoc = relsParser.parseFromString(relsXml, 'text/xml');
        }
        
        const xml = await sheetFile.async('string');
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');
        
        renderSheet(doc, sheetRelsDoc);
    }

    // ==========================================
    // 渲染工作表
    // ==========================================

    function renderSheet(doc, sheetRelsDoc) {
        tableWrapper.innerHTML = '';
        
        // 解析冻结窗格信息
        const freezeInfo = parseFreezePanes(doc);
        console.log('冻结窗格:', freezeInfo);
        
        const table = document.createElement('table');
        table.className = 'spreadsheet-table';

        // 默认字体（通常是 fonts[0]）
        if (fonts[0]) {
            table.style.fontFamily = getFontFamily(fonts[0].name, false);
            table.style.fontSize = (fonts[0].size || 11) + 'pt';
            table.style.fontWeight = getFontWeight(fonts[0].name, fonts[0].bold);
        }
        // 计算默认字体的最大数字宽度（MDW）
        currentMDW = measureMaxDigitWidth();

        // 超链接主题色
        const hlinkColor = resolveColor({ theme: 10 }) || '#0563c1';
        table.style.setProperty('--hyperlink-color', hlinkColor);
        
        // 获取列宽与列样式
        const colWidths = {};
        const colStyles = {};
        doc.querySelectorAll('col').forEach(col => {
            const min = parseInt(col.getAttribute('min'));
            const max = parseInt(col.getAttribute('max'));
            const width = parseFloat(col.getAttribute('width')) || 8.43;
            const styleAttr = col.getAttribute('style') || col.getAttribute('s');
            const styleIndex = styleAttr !== null ? parseInt(styleAttr) : null;
            for (let i = min; i <= max; i++) {
                colWidths[i] = width;
                if (!Number.isNaN(styleIndex) && styleIndex !== null) {
                    colStyles[i] = styleIndex;
                }
            }
        });
        
        // 获取合并单元格
        const merges = new Map();
        // 超链接映射（ref -> { target })
        const hyperlinkMap = parseHyperlinks(doc, sheetRelsDoc);
        const mergeSpanMap = new Map();
        doc.querySelectorAll('mergeCell').forEach(mc => {
            const ref = mc.getAttribute('ref');
            if (ref) {
                const [start, end] = ref.split(':');
                const s = cellRef(start), e = cellRef(end);
                merges.set(`${s.r},${s.c}`, {
                    rowspan: e.r - s.r + 1,
                    colspan: e.c - s.c + 1,
                    isMaster: true
                });
                for (let r = s.r; r <= e.r; r++) {
                    for (let c = s.c; c <= e.c; c++) {
                        mergeSpanMap.set(`${r},${c}`, {
                            startRow: s.r,
                            startCol: s.c,
                            endRow: e.r,
                            endCol: e.c
                        });
                    }
                }
                for (let r = s.r; r <= e.r; r++) {
                    for (let c = s.c; c <= e.c; c++) {
                        if (r !== s.r || c !== s.c) {
                            merges.set(`${r},${c}`, { hidden: true });
                        }
                    }
                }
            }
        });
        
        // 获取所有行
        const rows = doc.querySelectorAll('row');
        if (rows.length === 0) {
            tableWrapper.innerHTML = '<div style="padding:2rem;text-align:center;color:#666;">空工作表</div>';
            return;
        }
        
        // 计算最大列数
        let maxCol = 1;
        rows.forEach(row => {
            row.querySelectorAll('c').forEach(c => {
                const ref = cellRef(c.getAttribute('r'));
                if (ref.c > maxCol) maxCol = ref.c;
            });
        });
        maxCol = Math.min(maxCol, 100);
        currentMaxCol = maxCol;
        currentColWidths = { ...colWidths };

        // 预构建行样式与单元格样式索引映射
        const rowStyleMap = {};
        const cellStyleMap = new Map();
        const cellElemMap = new Map();
        rows.forEach(rowElem => {
            const rowNum = parseInt(rowElem.getAttribute('r'));
            const rowStyleAttr = rowElem.getAttribute('s');
            if (rowStyleAttr !== null) {
                const idx = parseInt(rowStyleAttr);
                if (!Number.isNaN(idx)) rowStyleMap[rowNum] = idx;
            }
            rowElem.querySelectorAll('c').forEach(c => {
                const ref = cellRef(c.getAttribute('r'));
                const sAttr = c.getAttribute('s');
                if (sAttr !== null) {
                    const sIdx = parseInt(sAttr);
                    if (!Number.isNaN(sIdx)) cellStyleMap.set(`${ref.r},${ref.c}`, sIdx);
                }
                cellElemMap.set(`${ref.r},${ref.c}`, c);
            });
        });

        const resolveStyleIndex = (r, c) => {
            const cellKey = `${r},${c}`;
            if (cellStyleMap.has(cellKey)) return cellStyleMap.get(cellKey);
            if (rowStyleMap[r] !== undefined) return rowStyleMap[r];
            if (colStyles[c] !== undefined) return colStyles[c];
            return null;
        };
        
        // 列宽控制（使用 colgroup 保证列宽准确）
        const colgroup = document.createElement('colgroup');
        const rowHeaderCol = document.createElement('col');
        const rowHeaderWidth = 46;
        rowHeaderCol.style.width = `${rowHeaderWidth}px`;
        colgroup.appendChild(rowHeaderCol);
        let totalTableWidth = rowHeaderWidth;
        for (let c = 1; c <= maxCol; c++) {
            const col = document.createElement('col');
            const w = calcColumnWidthPx(colWidths[c] || 8.43);
            col.style.width = `${w}px`;
            colgroup.appendChild(col);
            totalTableWidth += w;
        }
        table.appendChild(colgroup);
        table.style.width = `${totalTableWidth}px`;
        table.style.minWidth = `${totalTableWidth}px`;

        // 表头（列标题 A, B, C...）
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        headerRow.className = 'header-row';
        headerRow.appendChild(document.createElement('th'));

        for (let c = 1; c <= maxCol; c++) {
            const th = document.createElement('th');
            th.textContent = colName(c);
            th.dataset.col = String(c);
            const w = calcColumnWidthPx(colWidths[c] || 8.43);
            th.style.width = th.style.minWidth = `${w}px`;
            headerRow.appendChild(th);
        }
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // 计算冻结行的高度（用于 sticky 定位）
        const defaultRowHeight = 20; // 默认行高 px
        const headerHeight = 22; // 表头行高
        
        // 表体
        const tbody = document.createElement('tbody');
        let lastRow = 0;
        
        // 预先收集所有行的高度（用于计算冻结行的 top 值）
        const rowHeights = {};
        rows.forEach(rowElem => {
            const rowNum = parseInt(rowElem.getAttribute('r'));
            const ht = rowElem.getAttribute('ht');
            const customHeight = rowElem.getAttribute('customHeight') === 'true';
            rowHeights[rowNum] = {
                height: (customHeight && ht) ? (parseFloat(ht) * 1.33) : null,
                calcHeight: ht ? (parseFloat(ht) * 1.33) : defaultRowHeight
            };
        });
        
        // 计算每个冻结行的 top 偏移量
        const frozenRowTops = {};
        if (freezeInfo.isFrozen) {
            let cumTop = headerHeight;
            for (let r = 1; r <= freezeInfo.frozenRows; r++) {
                frozenRowTops[r] = cumTop;
                const rh = rowHeights[r] ? rowHeights[r].calcHeight : defaultRowHeight;
                cumTop += rh;
            }
        }
        
        rows.forEach(rowElem => {
            const rowNum = parseInt(rowElem.getAttribute('r'));
            
            // 填充空行
            while (lastRow < rowNum - 1 && lastRow < 1000) {
                lastRow++;
                const emptyTr = document.createElement('tr');
                
                // 空行也可能是冻结行
                if (freezeInfo.isFrozen && lastRow <= freezeInfo.frozenRows) {
                    emptyTr.classList.add('frozen-row');
                    emptyTr.style.top = (frozenRowTops[lastRow] || headerHeight) + 'px';
                }
                
                const rh = document.createElement('th');
                rh.textContent = lastRow;
                emptyTr.appendChild(rh);
                for (let c = 1; c <= maxCol; c++) {
                    emptyTr.appendChild(document.createElement('td'));
                }
                tbody.appendChild(emptyTr);
            }
            
            lastRow = rowNum;
            if (rowNum > 1000) return;
            
            const tr = document.createElement('tr');

            // 行高
            const rowHeightInfo = rowHeights[rowNum];
            if (rowHeightInfo && rowHeightInfo.height !== null) {
                tr.style.height = rowHeightInfo.height + 'px';
            }
            
            // 检查是否是冻结行
            if (freezeInfo.isFrozen && rowNum <= freezeInfo.frozenRows) {
                tr.classList.add('frozen-row');
                tr.style.top = frozenRowTops[rowNum] + 'px';
                
                // 标记最后一个冻结行
                if (rowNum === freezeInfo.frozenRows) {
                    tr.classList.add('frozen-row-last');
                }
            }
            
            // 行号（始终固定在左侧）
            const rowHeader = document.createElement('th');
            rowHeader.textContent = rowNum;
            rowHeader.dataset.row = String(rowNum);
            tr.appendChild(rowHeader);

            // 计算冻结列的宽度偏移（用于 sticky left 定位）
            let frozenColOffset = 46; // 行号列宽度
            
            // 收集此行的单元格
            const cellMap = {};
            rowElem.querySelectorAll('c').forEach(c => {
                const ref = cellRef(c.getAttribute('r'));
                cellMap[ref.c] = c;
            });
            
            // 渲染每个单元格
            for (let c = 1; c <= maxCol; c++) {
                const mergeKey = `${rowNum},${c}`;
                const merge = merges.get(mergeKey);
                
                if (merge && merge.hidden) continue;
                
                const td = document.createElement('td');

                if (merge && merge.isMaster) {
                    td.rowSpan = merge.rowspan;
                    td.colSpan = merge.colspan;
                }
                
                const cellElem = cellMap[c];
                const resolvedStyleIndex = resolveStyleIndex(rowNum, c);
                const cellStyle = (resolvedStyleIndex !== null && resolvedStyleIndex !== undefined)
                    ? cellXfs[resolvedStyleIndex]
                    : null;

                // 文本溢出（Excel 允许文本溢出到右侧连续空单元格）
                let overflowEndCol = null;
                if (cellElem) {
                    overflowEndCol = getOverflowEndCol(
                        cellElem,
                        resolvedStyleIndex,
                        rowNum,
                        c,
                        cellMap,
                        mergeSpanMap,
                        resolveStyleIndex
                    );
                }

                if (cellElem) {
                    renderCell(td, cellElem, resolvedStyleIndex, rowNum, c, resolveStyleIndex, overflowEndCol, hyperlinkMap);
                } else if (resolvedStyleIndex !== null && !Number.isNaN(resolvedStyleIndex)) {
                    // 空单元格也要应用样式（尤其是边框）
                    renderCell(td, null, resolvedStyleIndex, rowNum, c, resolveStyleIndex, overflowEndCol, hyperlinkMap);
                }

                if (overflowEndCol && overflowEndCol > c) {
                    td.classList.add('cell-overflow');
                }
                td.dataset.row = String(rowNum);
                td.dataset.col = String(c);

                // 合并单元格需要合并范围的边框应用到主单元格
                if (merge && merge.isMaster) {
                    applyMergedBorders(td, rowNum, c, merge, resolveStyleIndex);
                }
                
                // 冻结行的单元格需要确保有背景色（防止内容透视）
                const isFrozenRow = freezeInfo.isFrozen && rowNum <= freezeInfo.frozenRows;
                const isFrozenCol = freezeInfo.isFrozen && c <= freezeInfo.frozenCols;
                
                if (isFrozenRow || isFrozenCol) {
                    if (!td.style.backgroundColor) {
                        td.style.backgroundColor = '#ffffff';
                    }
                }
                
                // 冻结列处理
                if (isFrozenCol) {
                    td.classList.add('frozen-col');
                    td.style.position = 'sticky';
                    td.style.left = frozenColOffset + 'px';
                    td.style.zIndex = isFrozenRow ? '20' : '10';
                    
                    // 最后一个冻结列添加标记
                    if (c === freezeInfo.frozenCols) {
                        td.classList.add('frozen-col-last');
                    }
                }

                tr.appendChild(td);
                
                // 更新冻结列偏移
                if (c <= freezeInfo.frozenCols) {
                    const colWidth = calcColumnWidthPx(colWidths[c] || 8.43);
                    frozenColOffset += colWidth;
            }
        }
            
            tbody.appendChild(tr);
        });
        
        table.appendChild(tbody);
        tableWrapper.appendChild(table);
        adjustAutoRowHeights(table, rowHeights, defaultRowHeight, freezeInfo);

        // 保存单元格元数据，供选择/调试使用
        table._mergeSpanMap = mergeSpanMap;
        table._cellElemMap = cellElemMap;
        table._resolveStyleIndex = resolveStyleIndex;
        table._colWidths = { ...colWidths };

        // 绑定悬停与选中效果
        bindHoverEffect(table);
        bindSelectionEffect(table);
        
        cellInfo.textContent = `${lastRow} 行 × ${maxCol} 列`;
    }

    function renderCell(td, cellElem, styleIndex, rowNum, colNum, resolveStyleIndex, overflowEndCol, hyperlinkMap) {
        const resolvedStyleIndex = (styleIndex !== undefined && styleIndex !== null && !Number.isNaN(styleIndex))
            ? styleIndex
            : (cellElem && cellElem.getAttribute('s') !== null)
                ? parseInt(cellElem.getAttribute('s')) || 0
                : 0;
        const cellStyle = cellXfs[resolvedStyleIndex];
        
        if (cellElem) {
            // 获取值
            const type = cellElem.getAttribute('t');
            const vElem = cellElem.querySelector('v');
            let value = vElem ? vElem.textContent : '';
            
            // 处理共享字符串
            if (type === 's' && value) {
                const idx = parseInt(value);
                value = sharedStrings[idx] || '';
            } else if (type === 'inlineStr') {
                // 处理内联字符串（含富文本/换行）
                const isElem = cellElem.querySelector('is');
                if (isElem) {
                    value = extractRichText(isElem);
                }
            }
            
            // 处理布尔值
            if (type === 'b') {
                value = value === '1' ? 'TRUE' : 'FALSE';
            }
            
            // 处理数字格式
            if (!type || type === 'n') {
                const numVal = parseFloat(value);
                if (!isNaN(numVal) && cellStyle) {
                    const fmtId = cellStyle.numFmtId;
                    value = formatNumber(numVal, fmtId);
                }
                td.classList.add('cell-number');
            }
            
            const ref = `${colName(colNum)}${rowNum}`;
            const hyperlink = hyperlinkMap && hyperlinkMap.get(ref);
            if (hyperlink) {
                td.innerHTML = buildHyperlinkHtml(value, hyperlink.target);
            } else {
                td.textContent = value;
            }

            // 溢出文本覆盖（只对文本类生效）
            if (overflowEndCol && overflowEndCol > colNum) {
                applyOverflowOverlay(td, value, colNum, overflowEndCol);
            }
        }
        
        // === 应用样式（无论是否有值） ===
        applyCellStyle(td, cellStyle, rowNum, colNum, resolveStyleIndex);
    }

    function applyOverflowOverlay(td, text, startCol, endCol) {
        if (!text) return;
        const span = document.createElement('span');
        span.className = 'cell-overflow-text';
        span.textContent = text;
        td.textContent = '';
        td.appendChild(span);
        const width = calcOverflowWidthPx(startCol, endCol);
        span.style.width = width + 'px';
    }

    function calcOverflowWidthPx(startCol, endCol) {
        let w = 0;
        for (let c = startCol; c <= endCol; c++) {
            const cw = calcColumnWidthPx(currentColWidths[c] || 8.43);
            w += cw;
        }
        return w;
    }

    function measureTextWidth(text, cellStyle) {
        if (!text) return 0;
        if (!measureTextWidth._ctx) {
            const canvas = document.createElement('canvas');
            measureTextWidth._ctx = canvas.getContext('2d');
        }
        const ctx = measureTextWidth._ctx;
        const font = getFontSpec(cellStyle);
        ctx.font = font;
        return ctx.measureText(String(text)).width;
    }

    function measureMaxDigitWidth() {
        const ctx = measureTextWidth._ctx || (measureTextWidth._ctx = document.createElement('canvas').getContext('2d'));
        ctx.font = getFontSpec(null);
        const digits = ['0','1','2','3','4','5','6','7','8','9'];
        let max = 0;
        for (const d of digits) {
            const w = ctx.measureText(d).width;
            if (w > max) max = w;
        }
        return max || 7;
    }

    function getFontSizePx(cellStyle) {
        const fontId = cellStyle ? cellStyle.fontId : null;
        const font = (fontId !== null && fontId !== undefined) ? fonts[fontId] : null;
        const fontSizePt = (font && font.size) ? font.size : ((fonts[0] && fonts[0].size) || 11);
        return (fontSizePt * 96) / 72;
    }

    function adjustAutoRowHeights(table, rowHeights, defaultRowHeight, freezeInfo) {
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(tr => {
            const header = tr.querySelector('th[data-row]');
            if (!header) return;
            const rowNum = parseInt(header.dataset.row);
            if (freezeInfo.isFrozen && rowNum <= freezeInfo.frozenRows) return;
            const rowHeightInfo = rowHeights[rowNum];
            if (rowHeightInfo && rowHeightInfo.height !== null) return;

            let maxHeight = defaultRowHeight;
            tr.querySelectorAll('td.cell-wrap').forEach(td => {
                td.style.height = 'auto';
                const h = td.scrollHeight || 0;
                if (h > maxHeight) maxHeight = h;
            });
            if (maxHeight > defaultRowHeight) {
                tr.style.height = `${maxHeight}px`;
                if (rowHeights[rowNum]) rowHeights[rowNum].calcHeight = maxHeight;
            }
        });
    }

    function getFontSpec(cellStyle) {
        const fontId = cellStyle ? cellStyle.fontId : null;
        const font = (fontId !== null && fontId !== undefined) ? fonts[fontId] : null;
        const fontName = (font && font.name) || (fonts[0] && fonts[0].name) || '等线';
        const fontWeight = getFontWeight(fontName, font && font.bold);
        const fontStyle = (font && font.italic) ? 'italic' : 'normal';
        const fontSizePt = (font && font.size) ? font.size : ((fonts[0] && fonts[0].size) || 11);
        const fontSizePx = (fontSizePt * 96) / 72;
        return `${fontStyle} ${fontWeight} ${fontSizePx}px ${getFontFamily(fontName, font && font.bold)}`;
    }

    function getOverflowEndCol(cellElem, styleIndex, rowNum, colNum, cellMap, mergeSpanMap, resolveStyleIndex) {
        if (!cellElem) return null;
        const cellType = cellElem.getAttribute('t');
        const isTextType = cellType === 's' || cellType === 'str' || cellType === 'inlineStr';
        if (!isTextType) return null;
        const cellStyle = (styleIndex !== null && styleIndex !== undefined) ? cellXfs[styleIndex] : null;
        const wrapText = cellStyle && cellStyle.alignment && cellStyle.alignment.wrapText;
        if (wrapText) return null;
        const hAlign = cellStyle && cellStyle.alignment && cellStyle.alignment.horizontal;
        if (hAlign === 'center' || hAlign === 'right' || hAlign === 'justify' || hAlign === 'distributed') {
        return null;
    }
        const shrinkToFit = cellStyle && cellStyle.alignment && cellStyle.alignment.shrinkToFit;
        if (shrinkToFit) return null;
        const rotation = cellStyle && cellStyle.alignment && cellStyle.alignment.textRotation;
        if (rotation && rotation !== 0) return null;

        const text = getCellDisplayText(cellElem, cellStyle);
        if (!text) return null;

        // 如果当前单元格在合并区域内，禁用溢出
        if (mergeSpanMap && mergeSpanMap.has(`${rowNum},${colNum}`)) return null;

        const cellWidth = calcColumnWidthPx(currentColWidths[colNum] || 8.43);
        const textWidth = measureTextWidth(text, cellStyle);
        const paddingX = 6; // 左右各 3px

        // 文本本身没有超过当前单元格宽度，则不需要溢出
        if (textWidth <= Math.max(0, cellWidth - paddingX)) return null;

        // 仅允许溢出到右侧连续空单元格（无值）
        let endCol = colNum;
        for (let c = colNum + 1; c <= currentMaxCol; c++) {
            const key = `${rowNum},${c}`;
            if (mergeSpanMap && mergeSpanMap.has(key)) break;
            const neighbor = cellMap[c];
            if (neighbor) {
                const neighborText = getCellDisplayText(neighbor, resolveStyleIndex ? resolveStyleIndex(rowNum, c) : null);
                if (neighborText) break;
            }
            endCol = c;
        }

        return endCol > colNum ? endCol : null;
    }


    function getCellDisplayText(cellElem, styleIndex) {
        if (!cellElem) return '';
        const type = cellElem.getAttribute('t');
        const vElem = cellElem.querySelector('v');
        let value = vElem ? vElem.textContent : '';
        if (type === 's' && value) {
            const idx = parseInt(value);
            value = sharedStrings[idx] || '';
        } else if (type === 'inlineStr') {
            const isElem = cellElem.querySelector('is');
            if (isElem) {
                value = extractRichText(isElem);
            }
        } else if (type === 'b') {
            value = value === '1' ? 'TRUE' : 'FALSE';
        }
        if ((!type || type === 'n') && value !== '') {
            const numVal = parseFloat(value);
            if (!Number.isNaN(numVal)) {
                const cellStyle = (styleIndex !== null && styleIndex !== undefined) ? cellXfs[styleIndex] : null;
                const fmtId = cellStyle ? cellStyle.numFmtId : null;
                value = formatNumber(numVal, fmtId);
            }
        }
        return String(value || '');
    }

    function bindHoverEffect(table) {
        let lastCell = null;
        table.addEventListener('mouseover', (e) => {
            const cell = e.target.closest('td');
            if (!cell || cell === lastCell) return;
            if (lastCell) lastCell.classList.remove('cell-hover');
            cell.classList.add('cell-hover');
            lastCell = cell;
        });
        table.addEventListener('mouseout', (e) => {
            const toEl = e.relatedTarget;
            if (!toEl || !table.contains(toEl)) {
                if (lastCell) lastCell.classList.remove('cell-hover');
                lastCell = null;
            }
        });
    }

    function bindSelectionEffect(table) {
        let selectedCell = null;
        let selectedColHeaders = [];
        let selectedRowHeaders = [];
        let selectedRowCells = [];
        let selectedColCells = [];

        table.addEventListener('click', (e) => {
            const cell = e.target.closest('td');
            if (!cell) return;

            if (selectedCell) selectedCell.classList.remove('cell-selected');
            selectedCell = cell;
            cell.classList.add('cell-selected');

            const col = cell.dataset.col;
            const row = cell.dataset.row;

            selectedColHeaders.forEach(h => h.classList.remove('col-selected-header'));
            selectedRowHeaders.forEach(h => h.classList.remove('row-selected-header'));
            selectedColHeaders = [];
            selectedRowHeaders = [];

            selectedRowCells.forEach(c => c.classList.remove('cell-row-highlight'));
            selectedColCells.forEach(c => c.classList.remove('cell-col-highlight'));
            selectedRowCells = [];
            selectedColCells = [];

            let colStart = col ? parseInt(col) : null;
            let colEnd = colStart;
            let rowStart = row ? parseInt(row) : null;
            let rowEnd = rowStart;

            // 合并单元格：扩展行列范围
            const span = table._mergeSpanMap ? table._mergeSpanMap.get(`${row},${col}`) : null;
            if (span) {
                colStart = span.startCol;
                colEnd = span.endCol;
                rowStart = span.startRow;
                rowEnd = span.endRow;
            }

            if (colStart !== null && colEnd !== null) {
                for (let c = colStart; c <= colEnd; c++) {
                    const header = table.querySelector(`thead th[data-col="${c}"]`);
                    if (header) {
                        header.classList.add('col-selected-header');
                        selectedColHeaders.push(header);
                    }
                    selectedColCells.push(...Array.from(table.querySelectorAll(`tbody td[data-col="${c}"]`)));
                }
                selectedColCells.forEach(c => c.classList.add('cell-col-highlight'));
            }

            if (rowStart !== null && rowEnd !== null) {
                for (let r = rowStart; r <= rowEnd; r++) {
                    const header = table.querySelector(`tbody th[data-row="${r}"]`);
                    if (header) {
                        header.classList.add('row-selected-header');
                        selectedRowHeaders.push(header);
                    }
                    selectedRowCells.push(...Array.from(table.querySelectorAll(`tbody td[data-row="${r}"]`)));
                }
                selectedRowCells.forEach(c => c.classList.add('cell-row-highlight'));
            }

            // 调试选中单元格的样式/溢出判断
            debugSelectedCell(table, row, col);
        });
    }

    function debugSelectedCell(table, row, col) {
        if (!row || !col) return;
        const key = `${row},${col}`;
        const cellElem = table._cellElemMap ? table._cellElemMap.get(key) : null;
        const resolveStyleIndex = table._resolveStyleIndex;
        const styleIndex = resolveStyleIndex ? resolveStyleIndex(parseInt(row), parseInt(col)) : null;
        const cellStyle = (styleIndex !== null && styleIndex !== undefined) ? cellXfs[styleIndex] : null;
        const text = cellElem ? getCellDisplayText(cellElem, cellStyle) : '';
        const cellWidth = calcColumnWidthPx((table._colWidths && table._colWidths[col]) || 8.43);
        const textWidth = measureTextWidth(text, cellStyle);
        const overflowEnd = cellElem
            ? getOverflowEndCol(cellElem, styleIndex, parseInt(row), parseInt(col), {}, table._mergeSpanMap, resolveStyleIndex)
            : null;

        console.log('[Cell Debug]', {
            cell: `${col}${row}`,
            styleIndex,
            alignment: cellStyle ? cellStyle.alignment : null,
            wrapText: cellStyle && cellStyle.alignment ? cellStyle.alignment.wrapText : null,
            shrinkToFit: cellStyle && cellStyle.alignment ? cellStyle.alignment.shrinkToFit : null,
            textRotation: cellStyle && cellStyle.alignment ? cellStyle.alignment.textRotation : null,
            text,
            textWidth,
            cellWidth,
            overflowEnd
        });
    }

    function applyCellStyle(td, cellStyle, rowNum, colNum, resolveStyleIndex) {
        if (!cellStyle) return;
        
        // 字体
        const font = fonts[cellStyle.fontId];
        if (font) {
            const fontName = font.name || (fonts[0] && fonts[0].name) || '等线';
            td.style.fontFamily = getFontFamily(fontName, font.bold);
            if (font.size) td.style.fontSize = font.size + 'pt';
            td.style.fontWeight = getFontWeight(fontName, font.bold);
            if (font.italic) td.style.fontStyle = 'italic';
            if (font.underline) td.style.textDecoration = 'underline';
            if (font.strike) td.style.textDecoration = (td.style.textDecoration || '') + ' line-through';
            if (font.color) {
                const color = resolveColor(font.color);
                if (color) td.style.color = color;
            }
        }
        
        // 填充（背景色）
        const fill = fills[cellStyle.fillId];
        if (fill && fill.type === 'pattern') {
            if (fill.pattern === 'solid' && fill.fgColor) {
                const bgColor = resolveColor(fill.fgColor);
                if (bgColor) td.style.backgroundColor = bgColor;
            } else if (fill.fgColor) {
                const bgColor = resolveColor(fill.fgColor);
                if (bgColor) td.style.backgroundColor = bgColor;
            }
        }
        
        // 边框
        const border = borders[cellStyle.borderId];
        const applyBorderSide = (side, bs) => {
            if (!bs || !bs.style) return;
            // 忽略外侧列边框（第1列左边、最后一列右边）
            if (side === 'left' && colNum === 1) return;
            if (side === 'right' && currentMaxCol && colNum === currentMaxCol) return;
            let w = '1px', s = 'solid';
            switch (bs.style) {
                case 'thin': w = '1px'; break;
                case 'medium': w = '2px'; break;
                case 'thick': w = '3px'; break;
                case 'hair': w = '0.5px'; break;
                case 'dotted': s = 'dotted'; break;
                case 'dashed': s = 'dashed'; break;
                case 'double': s = 'double'; w = '3px'; break;
            }
            const c = bs.color ? resolveColor(bs.color) : '#000';
            td.style[`border${side.charAt(0).toUpperCase() + side.slice(1)}`] = `${w} ${s} ${c}`;
        };

        if (border) {
            ['top', 'right', 'bottom', 'left'].forEach(side => {
                applyBorderSide(side, border[side]);
            });
        }

        // 如果当前单元格缺少边框，尝试从相邻单元格补齐
        if (typeof resolveStyleIndex === 'function') {
            // 左侧邻居的右边框
            if (!td.style.borderLeft && colNum > 1) {
                const leftStyleIdx = resolveStyleIndex(rowNum, colNum - 1);
                if (leftStyleIdx !== null && leftStyleIdx !== undefined) {
                    const leftStyle = cellXfs[leftStyleIdx];
                    const leftBorder = leftStyle ? borders[leftStyle.borderId] : null;
                    if (leftBorder && leftBorder.right) {
                        applyBorderSide('left', leftBorder.right);
                    }
                }
            }
            // 右侧邻居的左边框
            if (!td.style.borderRight) {
                const rightStyleIdx = resolveStyleIndex(rowNum, colNum + 1);
                if (rightStyleIdx !== null && rightStyleIdx !== undefined) {
                    const rightStyle = cellXfs[rightStyleIdx];
                    const rightBorder = rightStyle ? borders[rightStyle.borderId] : null;
                    if (rightBorder && rightBorder.left) {
                        applyBorderSide('right', rightBorder.left);
                    }
                }
            }
            // 上方邻居的下边框
            if (!td.style.borderTop && rowNum > 1) {
                const topStyleIdx = resolveStyleIndex(rowNum - 1, colNum);
                if (topStyleIdx !== null && topStyleIdx !== undefined) {
                    const topStyle = cellXfs[topStyleIdx];
                    const topBorder = topStyle ? borders[topStyle.borderId] : null;
                    if (topBorder && topBorder.bottom) {
                        applyBorderSide('top', topBorder.bottom);
                    }
                }
            }
            // 下方邻居的上边框
            if (!td.style.borderBottom) {
                const bottomStyleIdx = resolveStyleIndex(rowNum + 1, colNum);
                if (bottomStyleIdx !== null && bottomStyleIdx !== undefined) {
                    const bottomStyle = cellXfs[bottomStyleIdx];
                    const bottomBorder = bottomStyle ? borders[bottomStyle.borderId] : null;
                    if (bottomBorder && bottomBorder.top) {
                        applyBorderSide('bottom', bottomBorder.top);
                    }
                }
            }
        }
        
        // 对齐
        if (cellStyle.alignment) {
            const a = cellStyle.alignment;
            if (a.horizontal) {
                const map = { left: 'left', center: 'center', right: 'right', justify: 'justify' };
                td.style.textAlign = map[a.horizontal] || a.horizontal;
            }
            if (a.vertical) {
                const map = { top: 'top', center: 'middle', bottom: 'bottom' };
                td.style.verticalAlign = map[a.vertical] || 'bottom';
            }
            if (a.wrapText) {
                td.style.whiteSpace = 'pre-wrap';
                td.style.wordBreak = 'break-word';
                td.classList.add('cell-wrap');
            }
            if (a.indent) {
                td.style.paddingLeft = (a.indent * 10) + 'px';
            }
        }
    }

    function applyMergedBorders(td, rowNum, colNum, merge, resolveStyleIndex) {
        const startRow = rowNum;
        const startCol = colNum;
        const endRow = rowNum + merge.rowspan - 1;
        const endCol = colNum + merge.colspan - 1;

        const applyBorderSide = (side, bs) => {
            if (!bs || !bs.style) return;
            // 忽略外侧列边框（第1列左边、最后一列右边）
            if (side === 'left' && startCol === 1) return;
            if (side === 'right' && currentMaxCol && endCol === currentMaxCol) return;
            let w = '1px', s = 'solid';
            switch (bs.style) {
                case 'thin': w = '1px'; break;
                case 'medium': w = '2px'; break;
                case 'thick': w = '3px'; break;
                case 'hair': w = '0.5px'; break;
                case 'dotted': s = 'dotted'; break;
                case 'dashed': s = 'dashed'; break;
                case 'double': s = 'double'; w = '3px'; break;
            }
            const c = bs.color ? resolveColor(bs.color) : '#000';
            td.style[`border${side.charAt(0).toUpperCase() + side.slice(1)}`] = `${w} ${s} ${c}`;
        };

        const getBorderForCell = (r, c) => {
            const styleIdx = resolveStyleIndex(r, c);
            if (styleIdx === null || styleIdx === undefined) return null;
            const style = cellXfs[styleIdx];
            if (!style) return null;
            return borders[style.borderId] || null;
        };

        // 顶边框：取合并区域第一行各单元格的 top
        if (!td.style.borderTop) {
            for (let c = startCol; c <= endCol; c++) {
                const b = getBorderForCell(startRow, c);
                if (b && b.top) {
                    applyBorderSide('top', b.top);
                    break;
                }
            }
        }
        // 底边框：取合并区域最后一行各单元格的 bottom
        if (!td.style.borderBottom) {
            for (let c = startCol; c <= endCol; c++) {
                const b = getBorderForCell(endRow, c);
                if (b && b.bottom) {
                    applyBorderSide('bottom', b.bottom);
                    break;
                }
            }
        }
        // 左边框：取合并区域第一列各单元格的 left
        if (!td.style.borderLeft) {
            for (let r = startRow; r <= endRow; r++) {
                const b = getBorderForCell(r, startCol);
                if (b && b.left) {
                    applyBorderSide('left', b.left);
                    break;
                }
            }
        }
        // 右边框：取合并区域最后一列各单元格的 right
        if (!td.style.borderRight) {
            for (let r = startRow; r <= endRow; r++) {
                const b = getBorderForCell(r, endCol);
                if (b && b.right) {
                    applyBorderSide('right', b.right);
                    break;
                }
            }
        }
    }

    function formatNumber(value, fmtId) {
        // 内置格式
        const builtIn = {
            0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
            9: '0%', 10: '0.00%', 14: 'yyyy/m/d', 22: 'yyyy/m/d h:mm'
        };
        
        let fmt = numFmts[fmtId] || builtIn[fmtId] || 'General';
        
        // 日期格式
        if (/[ymd]/i.test(fmt) && value > 0) {
            const date = new Date((value - 25569) * 86400 * 1000);
            if (!isNaN(date)) {
                return date.toLocaleDateString('zh-CN');
            }
        }
        
        // 百分比
        if (fmt.includes('%')) {
            const dec = (fmt.match(/0/g) || []).length - 1;
            return (value * 100).toFixed(Math.max(0, dec)) + '%';
        }
        
        // 千分位
        if (fmt.includes(',')) {
            const dec = (fmt.match(/\.0+/) || [''])[0].length - 1;
            return value.toLocaleString('zh-CN', { 
                minimumFractionDigits: Math.max(0, dec),
                maximumFractionDigits: Math.max(0, dec)
            });
        }
        
        // 小数
        const decMatch = fmt.match(/\.0+/);
        if (decMatch) {
            return value.toFixed(decMatch[0].length - 1);
        }
        
        // 整数或默认
        if (Number.isInteger(value)) return String(value);
        return parseFloat(value.toPrecision(10)).toString();
    }

    // ==========================================
    // 工具函数
    // ==========================================

    /**
     * 解析冻结窗格信息
     */
    function parseFreezePanes(doc) {
        const result = {
            frozenRows: 0,      // 冻结的行数
            frozenCols: 0,      // 冻结的列数
            isFrozen: false
        };
        
        // 查找 pane 元素
        const pane = doc.querySelector('sheetView pane');
        if (!pane) return result;
        
        const state = pane.getAttribute('state');
        if (state !== 'frozen' && state !== 'frozenSplit') return result;
        
        result.isFrozen = true;
        
        // ySplit = 冻结的行数
        const ySplit = pane.getAttribute('ySplit');
        if (ySplit) {
            result.frozenRows = parseInt(ySplit) || 0;
        }
        
        // xSplit = 冻结的列数
        const xSplit = pane.getAttribute('xSplit');
        if (xSplit) {
            result.frozenCols = parseInt(xSplit) || 0;
        }
        
        return result;
    }

    function cellRef(ref) {
        if (!ref) return { r: 1, c: 1 };
        let col = 0, row = '';
        for (const ch of ref) {
            if (ch >= 'A' && ch <= 'Z') {
                col = col * 26 + (ch.charCodeAt(0) - 64);
            } else {
                row += ch;
            }
        }
        return { r: parseInt(row) || 1, c: col || 1 };
    }

    function colName(n) {
        let s = '';
        while (n > 0) {
            s = String.fromCharCode(65 + (n - 1) % 26) + s;
            n = Math.floor((n - 1) / 26);
        }
        return s;
    }

    function calcColumnWidthPx(excelWidth) {
        // Excel column width to pixels based on max digit width (MDW)
        const w = Number(excelWidth);
        if (!Number.isFinite(w)) return 64;
        const mdw = currentMDW || 7;
        if (w <= 0) return 0;
        const pixels = Math.floor(((256 * w + Math.floor(128 / mdw)) / 256) * mdw) + 5;
        return pixels;
    }

    function getFontFamily(name, bold) {
        // 处理字体名称中的变体标识
        const baseName = name ? name.replace(/ Light$| Bold$/i, '').trim() : '';
        const isLightVariant = name && / Light$/i.test(name);
        
        const map = {
            '等线': {
                regular: '"DengXian", "等线", "Microsoft YaHei", sans-serif',
                bold: '"DengXian Bold", "等线 Bold", "Microsoft YaHei Bold", "Microsoft YaHei", sans-serif',
                light: '"DengXian Light", "等线 Light", "Microsoft YaHei Light", "Microsoft YaHei", sans-serif'
            },
            '宋体': {
                regular: '"SimSun", "宋体", "Noto Serif CJK SC", serif',
                bold: '"SimSun", "宋体", "Noto Serif CJK SC", serif'
            },
            '黑体': {
                regular: '"SimHei", "黑体", "Noto Sans CJK SC", sans-serif',
                bold: '"SimHei", "黑体", "Noto Sans CJK SC", sans-serif'
            },
            '微软雅黑': {
                regular: '"Microsoft YaHei", "微软雅黑", "PingFang SC", sans-serif',
                bold: '"Microsoft YaHei Bold", "微软雅黑 Bold", "PingFang SC Semibold", sans-serif',
                light: '"Microsoft YaHei Light", "微软雅黑 Light", "PingFang SC Light", sans-serif'
            },
            '楷体': {
                regular: '"KaiTi", "楷体", "STKaiti", serif',
                bold: '"KaiTi", "楷体", "STKaiti", serif'
            },
            'Calibri': {
                regular: 'Calibri, "Segoe UI", sans-serif',
                bold: '"Calibri Bold", Calibri, "Segoe UI Semibold", sans-serif',
                light: '"Calibri Light", Calibri, "Segoe UI Light", sans-serif'
            },
            'Arial': {
                regular: 'Arial, "Helvetica Neue", sans-serif',
                bold: '"Arial Bold", Arial, "Helvetica Neue", sans-serif'
            }
        };
        
        const fontDef = map[baseName];
        if (fontDef) {
            if (isLightVariant) return fontDef.light || fontDef.regular;
            if (bold) return fontDef.bold || fontDef.regular;
            return fontDef.regular;
        }
        
        // 未知字体，直接使用名称
        return `"${name}", "Microsoft YaHei", sans-serif`;
    }
    
    /**
     * 获取字体粗细值
     * Excel 的粗体不是简单的 bold，需要根据字体类型调整
     */
    function getFontWeight(fontName, isBold) {
        if (!isBold) {
            // 检查是否是 Light 变体
            if (fontName && / Light$/i.test(fontName)) {
                return '300';
            }
            return '400'; // normal
        }
        
        // 粗体：不同字体使用不同的粗细值
        // 中文字体通常粗体效果需要更高的 weight
        const heavyBoldFonts = ['等线', 'DengXian', '微软雅黑', 'Microsoft YaHei', 'Calibri'];
        const baseName = fontName ? fontName.replace(/ Light$| Bold$/i, '').trim() : '';
        
        if (heavyBoldFonts.some(f => baseName.includes(f) || f.includes(baseName))) {
            return '700'; // bold
        }
        
        return '700'; // bold
    }
});
