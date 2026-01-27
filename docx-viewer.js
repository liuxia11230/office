/**
 * DOCX Viewer - 纯前端预览
 * 使用 docx-preview 渲染 .docx，无需服务端，仅 CDN + JS
 * 左侧导航（标题 / 页面）与 PDF 预览一致
 */

class DOCXViewer {
    constructor(options) {
        this.previewContainer = options.previewContainer;
        this.uploadWrapper = options.uploadWrapper;
        this.controls = options.controls;
        this.container = null;
        this.bodyEl = null;
        this.sidebar = null;
        this.sidebarVisible = false;
        this.outlineContainer = null;
        this.thumbnailContainer = null;
        this.viewerWrap = null;
        this._docxReady = false;
        this._mammothReady = false;
        this._outlineData = null;
        this._docInfo = { totalWords: 0, totalChars: 0 };
        this._infoLabel = null;
    }

    async ensureMammothLib() {
        if (this._mammothReady && window.mammoth) return;
        if (window.mammoth) {
            this._mammothReady = true;
            return;
        }
        return new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = 'https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js';
            s.onload = () => {
                this._mammothReady = true;
                resolve();
            };
            s.onerror = () => reject(new Error('mammoth.js 加载失败'));
            document.head.appendChild(s);
        });
    }

    async ensureDocxLib() {
        if (this._docxReady && window.docx) return;
        if (window.docx) {
            this._docxReady = true;
            return;
        }
        const base = 'https://cdn.jsdelivr.net/npm/docx-preview@0.3.7/dist';
        if (!document.querySelector('link[href*="docx-preview"]')) {
            const link = document.createElement('link');
            link.rel = 'stylesheet';
            link.href = base + '/docx-preview.css';
            document.head.appendChild(link);
        }

        return new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = base + '/docx-preview.min.js';
            s.onload = () => {
                this._docxReady = true;
                resolve();
            };
            s.onerror = () => reject(new Error('docx-preview 加载失败'));
            document.head.appendChild(s);
        });
    }

    async loadFile(file) {
        await this.ensureDocxLib();

        this.uploadWrapper.classList.add('hide');
        this.controls.classList.remove('hide');
        this.previewContainer.classList.remove('hide');

        const cellInfoBar = this.previewContainer.querySelector('.cell-info-bar');
        if (cellInfoBar) cellInfoBar.style.display = 'none';

        const tableWrapper = this.previewContainer.querySelector('.table-wrapper');
        if (tableWrapper) tableWrapper.style.display = 'none';

        const pdfContainer = this.previewContainer.querySelector('.pdf-viewer-container');
        if (pdfContainer) pdfContainer.style.display = 'none';

        let root = this.previewContainer.querySelector('.docx-viewer-container');
        if (!root) {
            root = document.createElement('div');
            root.className = 'docx-viewer-container';
            this.previewContainer.appendChild(root);
        }
        root.style.display = 'flex';
        root.innerHTML = '';

        const mainLayout = document.createElement('div');
        mainLayout.className = 'docx-main-layout';
        mainLayout.style.cssText = 'display: flex; flex: 1; overflow: hidden;';

        this.createSidebar();
        mainLayout.appendChild(this.sidebar);

        const contentArea = document.createElement('div');
        contentArea.className = 'docx-content-area';
        contentArea.style.cssText = 'flex: 1; display: flex; flex-direction: column; overflow: hidden;';

        this.createToolbar();
        contentArea.appendChild(this.toolbar);

        const wrap = document.createElement('div');
        wrap.className = 'docx-viewer-wrap';
        wrap.style.cssText = `
            flex: 1;
            overflow: auto;
            background: #525252;
            display: flex;
            justify-content: center;
            padding: 24px;
        `;
        const body = document.createElement('div');
        body.className = 'docx-viewer-body';
        body.style.cssText = `
            width: 100%;
            min-width: 0;
            max-width: 210mm;
            background: #fff;
            box-shadow: 0 2px 12px rgba(0,0,0,0.2);
            padding: 24px;
        `;
        wrap.appendChild(body);
        contentArea.appendChild(wrap);

        mainLayout.appendChild(contentArea);
        root.appendChild(mainLayout);

        this.container = root;
        this.bodyEl = body;
        this.viewerWrap = wrap;

        const ext = (file.name.split('.').pop() || '').toLowerCase();
        if (ext !== 'docx') {
            body.innerHTML = '<p style="color:#666;">仅支持 .docx 格式</p>';
            this.updateStatus();
            return;
        }

        body.innerHTML = '<p style="color:#888;">加载中…</p>';

        try {
            const buf = await file.arrayBuffer();
            await this.ensureMammothLib();
            await this.extractOutlineWithMammoth(buf);
            await window.docx.renderAsync(buf, body, null, {
                className: 'docx',
                inWrapper: true,
                ignoreWidth: false,
                ignoreHeight: false,
                ignoreFonts: false,
                breakPages: true,
                renderHeaders: true,
                renderFooters: true,
            });
            await new Promise(resolve => setTimeout(resolve, 100));
            this.mapHeadingsToDOM();
            this.calculateDocInfo();
        } catch (e) {
            console.error('DOCX 渲染失败', e);
            body.innerHTML = `<p style="color:#c00;">预览失败：${e.message || '未知错误'}</p>`;
        }

        this.loadOutline();
        this.updateStatus();
        this.setupScrollListener();
    }

    createSidebar() {
        const sidebar = document.createElement('div');
        sidebar.className = 'docx-sidebar';
        sidebar.style.cssText = `
            width: 0;
            background: #2b2b2b;
            border-right: 1px solid #444;
            overflow: hidden;
            transition: width 0.3s;
            display: flex;
            flex-direction: column;
        `;

        const titleLabel = document.createElement('div');
        titleLabel.textContent = '标题导航';
        titleLabel.style.cssText = `
            padding: 12px 16px; background: #333; border-bottom: 1px solid #444;
            color: #fff; font-size: 13px; font-weight: 500;
        `;
        sidebar.appendChild(titleLabel);

        const outlineContainer = document.createElement('div');
        outlineContainer.className = 'docx-outline-container';
        outlineContainer.style.cssText = `
            flex: 1; overflow-y: auto; padding: 16px; color: #fff;
            display: flex; flex-direction: column;
        `;
        this.outlineContainer = outlineContainer;
        sidebar.appendChild(outlineContainer);

        this.sidebar = sidebar;
    }

    createToolbar() {
        const toolbar = document.createElement('div');
        toolbar.className = 'docx-toolbar';
        toolbar.style.cssText = `
            background: #2b2b2b; padding: 8px 16px; display: flex; align-items: center;
            gap: 12px; border-bottom: 1px solid #444; flex-shrink: 0;
        `;

        const toggleBtn = document.createElement('button');
        toggleBtn.textContent = '☰';
        toggleBtn.style.cssText = `
            background: #4a4a4a; border: 1px solid #666; color: #fff;
            padding: 6px 12px; cursor: pointer; border-radius: 3px; font-size: 16px; width: 36px;
        `;
        toggleBtn.addEventListener('click', () => this.toggleSidebar());
        toolbar.appendChild(toggleBtn);

        const sep1 = document.createElement('div');
        sep1.style.cssText = 'width:1px;height:24px;background:#666;margin:0 4px;';
        toolbar.appendChild(sep1);

        const infoLabel = document.createElement('span');
        infoLabel.style.cssText = 'color:#fff;font-size:12px;flex:1;';
        infoLabel.textContent = '加载中...';
        this._infoLabel = infoLabel;
        toolbar.appendChild(infoLabel);

        this.toolbar = toolbar;
    }

    toggleSidebar() {
        this.sidebarVisible = !this.sidebarVisible;
        this.sidebar.style.width = this.sidebarVisible ? '250px' : '0';
    }

    /**
     * 使用 mammoth.js 提取 DOCX 标题结构（导航标签信息）。
     */
    async extractOutlineWithMammoth(arrayBuffer) {
        if (!window.mammoth) {
            await this.ensureMammothLib();
        }
        try {
            const result = await window.mammoth.convertToHtml({ arrayBuffer }, {
                styleMap: [
                    "p[style-name='Heading 1'] => h1:fresh",
                    "p[style-name='Heading 2'] => h2:fresh",
                    "p[style-name='Heading 3'] => h3:fresh",
                    "p[style-name='Heading 4'] => h4:fresh",
                    "p[style-name='Heading 5'] => h5:fresh",
                    "p[style-name='Heading 6'] => h6:fresh",
                    "p[style-name='标题 1'] => h1:fresh",
                    "p[style-name='标题 2'] => h2:fresh",
                    "p[style-name='标题 3'] => h3:fresh",
                    "p[style-name='标题 4'] => h4:fresh",
                    "p[style-name='标题 5'] => h5:fresh",
                    "p[style-name='标题 6'] => h6:fresh",
                ]
            });
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = result.value;
            const headings = tempDiv.querySelectorAll('h1, h2, h3, h4, h5, h6');
            this._outlineData = Array.from(headings).map(h => ({
                text: (h.textContent || '').trim(),
                level: parseInt(h.tagName.charAt(1), 10) - 1,
                tagName: h.tagName
            }));
            console.log('mammoth 提取的标题:', this._outlineData.length, this._outlineData);
            const fullText = tempDiv.textContent || '';
            this._docInfo.totalChars = fullText.length;
            this._docInfo.totalWords = fullText.trim().split(/\s+/).filter(w => w.length > 0).length;
        } catch (e) {
            console.warn('mammoth.js 提取标题失败:', e);
            this._outlineData = null;
        }
    }

    /**
     * 加载导航：优先使用 mammoth.js 提取的结构，否则从渲染后的 HTML 提取。
     */
    loadOutline() {
        if (!this.outlineContainer || !this.bodyEl) return;
        if (this._outlineData && this._outlineData.length > 0) {
            this.renderOutlineFromMammoth();
            return;
        }
        let headings = this.bodyEl.querySelectorAll('h1, h2, h3, h4, h5, h6');
        if (!headings.length) {
            headings = this.bodyEl.querySelectorAll('[class*="Heading1"],[class*="Heading2"],[class*="Heading3"],[class*="Heading4"],[class*="Heading5"],[class*="Heading6"]');
        }
        if (!headings.length) {
            this.outlineContainer.innerHTML = '<div style="color:#999;font-size:12px;">未识别到标题，请使用滚动浏览</div>';
            return;
        }
        const ul = document.createElement('ul');
        ul.style.cssText = 'list-style:none;padding:0;margin:0;';
        const levelMap = { H1: 0, H2: 1, H3: 2, H4: 3, H5: 4, H6: 5 };
        const reHeading = /Heading([1-6])/i;
        headings.forEach((el, i) => {
            el.id = el.id || `docx-heading-${i}`;
            let level = levelMap[el.tagName];
            if (level == null && el.className) {
                const m = String(el.className).match(reHeading);
                level = m ? Math.min(5, parseInt(m[1], 10) - 1) : 0;
            }
            level = level ?? 0;
            const li = document.createElement('li');
            li.style.cssText = `margin:4px 0;padding-left:${level * 16}px;`;
            const a = document.createElement('a');
            a.href = '#';
            a.textContent = (el.textContent || '').trim() || '(无标题)';
            a.style.cssText = `
                color:#fff;text-decoration:none;display:block;padding:6px 8px;
                border-radius:3px;font-size:13px;cursor:pointer;line-height:1.4;
            `;
            a.addEventListener('mouseenter', () => { a.style.background = '#3a3a3a'; });
            a.addEventListener('mouseleave', () => { a.style.background = 'transparent'; });
            a.addEventListener('click', (e) => {
                e.preventDefault();
                el.scrollIntoView({ behavior: 'smooth', block: 'start' });
            });
            li.appendChild(a);
            ul.appendChild(li);
        });
        this.outlineContainer.innerHTML = '';
        this.outlineContainer.appendChild(ul);
    }

    /**
     * 使用 mammoth.js 提取的标题数据渲染导航。
     */
    renderOutlineFromMammoth() {
        if (!this._outlineData || !this._outlineData.length) return;
        const ul = document.createElement('ul');
        ul.style.cssText = 'list-style:none;padding:0;margin:0;';
        this._outlineData.forEach((item, i) => {
            const li = document.createElement('li');
            li.style.cssText = `margin:4px 0;padding-left:${item.level * 16}px;`;
            const a = document.createElement('a');
            a.href = '#';
            a.textContent = item.text || '(无标题)';
            a.style.cssText = `
                color:#fff;text-decoration:none;display:block;padding:6px 8px;
                border-radius:3px;font-size:13px;cursor:pointer;line-height:1.4;
            `;
            a.addEventListener('mouseenter', () => { a.style.background = '#3a3a3a'; });
            a.addEventListener('mouseleave', () => { a.style.background = 'transparent'; });
            a.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                console.log('导航点击:', { text: item.text, tagName: item.tagName, item });
                this.scrollToHeading(item.text, item.tagName);
            });
            li.appendChild(a);
            ul.appendChild(li);
        });
        this.outlineContainer.innerHTML = '';
        this.outlineContainer.appendChild(ul);
    }

    /**
     * 将 mammoth 提取的标题映射到渲染后的 DOM 元素。
     */
    mapHeadingsToDOM() {
        if (!this._outlineData || !this.bodyEl) {
            console.warn('mapHeadingsToDOM: 缺少数据', { hasData: !!this._outlineData, hasBody: !!this.bodyEl });
            return;
        }
        const allHeadings = [];
        const standardHeadings = this.bodyEl.querySelectorAll('h1, h2, h3, h4, h5, h6');
        console.log('mapHeadingsToDOM: 标准标题数量', standardHeadings.length);
        standardHeadings.forEach(el => allHeadings.push(el));
        const classHeadings = this.bodyEl.querySelectorAll('[class*="Heading1"],[class*="Heading2"],[class*="Heading3"],[class*="Heading4"],[class*="Heading5"],[class*="Heading6"]');
        console.log('mapHeadingsToDOM: class 标题数量', classHeadings.length);
        classHeadings.forEach(el => {
            if (!allHeadings.includes(el)) allHeadings.push(el);
        });
        if (!allHeadings.length) {
            console.warn('mapHeadingsToDOM: DOM 中未找到标准标题，尝试查找所有可能的标题元素');
            const allElements = this.bodyEl.querySelectorAll('p, div, span, h1, h2, h3, h4, h5, h6');
            console.log('DOM 中所有元素数量:', allElements.length);
            const sampleElements = Array.from(allElements).slice(0, 20);
            console.log('前20个元素的标签和类名:', sampleElements.map(el => ({
                tag: el.tagName,
                class: el.className,
                text: (el.textContent || '').trim().substring(0, 30)
            })));
        }
        console.log('mapHeadingsToDOM: 总标题数量', allHeadings.length, 'mammoth 提取数量', this._outlineData.length);
        if (!allHeadings.length) {
            console.warn('mapHeadingsToDOM: DOM 中未找到标题，将使用文本匹配方式');
            return;
        }
        const usedElements = new Set();
        this._outlineData.forEach((item, idx) => {
            const normalizedText = item.text.trim().toLowerCase();
            const expectedLevel = item.level;
            let bestMatch = null;
            let bestScore = 0;
            for (const el of allHeadings) {
                if (usedElements.has(el)) continue;
                const elLevel = this.getHeadingLevel(el);
                if (elLevel !== expectedLevel) continue;
                const elText = (el.textContent || '').trim().toLowerCase();
                if (elText === normalizedText) {
                    bestMatch = el;
                    bestScore = 100;
                    break;
                }
                const similarity = this.calculateTextSimilarity(normalizedText, elText);
                if (similarity > bestScore && similarity > 0.7) {
                    bestScore = similarity;
                    bestMatch = el;
                }
            }
            if (bestMatch) {
                item.domElement = bestMatch;
                usedElements.add(bestMatch);
                if (!bestMatch.id) {
                    bestMatch.id = `docx-heading-mammoth-${idx}`;
                }
                console.log(`映射成功 [${idx}]: "${item.text}" -> ${bestMatch.tagName}`, bestMatch);
            } else {
                console.warn(`映射失败 [${idx}]: "${item.text}" (${item.tagName}) 未找到匹配的 DOM 元素`);
            }
        });
        console.log('mapHeadingsToDOM 完成，成功映射:', this._outlineData.filter(d => d.domElement).length, '/', this._outlineData.length);
    }

    /**
     * 获取元素的标题层级（0-5 对应 h1-h6）。
     */
    getHeadingLevel(el) {
        const tagMatch = el.tagName.match(/^H([1-6])$/i);
        if (tagMatch) return parseInt(tagMatch[1], 10) - 1;
        const classMatch = String(el.className).match(/Heading([1-6])/i);
        if (classMatch) return parseInt(classMatch[1], 10) - 1;
        return -1;
    }

    /**
     * 计算两个文本的相似度（简单实现）。
     */
    calculateTextSimilarity(text1, text2) {
        if (text1 === text2) return 1.0;
        if (text1.includes(text2) || text2.includes(text1)) return 0.9;
        const len1 = text1.length;
        const len2 = text2.length;
        if (len1 === 0 || len2 === 0) return 0;
        const minLen = Math.min(len1, len2);
        const maxLen = Math.max(len1, len2);
        let matches = 0;
        for (let i = 0; i < minLen; i++) {
            if (text1[i] === text2[i]) matches++;
        }
        return matches / maxLen;
    }

    /**
     * 滚动到指定标题（优先使用已映射的 DOM 元素）。
     */
    scrollToHeading(text, tagName) {
        console.log('scrollToHeading 调用:', { text, tagName, hasBodyEl: !!this.bodyEl, hasViewerWrap: !!this.viewerWrap });
        if (!this.bodyEl || !text) {
            console.warn('scrollToHeading: 缺少必要参数', { bodyEl: !!this.bodyEl, text });
            return;
        }
        const normalizedText = text.trim().toLowerCase();
        let targetElement = null;
        const item = this._outlineData?.find(d => d.text === text && d.tagName === tagName);
        console.log('找到的 item:', item);
        if (item && item.domElement && this.bodyEl.contains(item.domElement)) {
            targetElement = item.domElement;
            console.log('使用映射的元素:', targetElement);
        } else {
            console.log('映射元素不存在，尝试在 DOM 中查找匹配的元素');
            const candidates = [];
            const standardHeadings = this.bodyEl.querySelectorAll(tagName.toLowerCase());
            console.log('标准标题数量:', standardHeadings.length);
            standardHeadings.forEach(el => candidates.push(el));
            const level = parseInt(tagName.charAt(1), 10);
            const classHeadings = this.bodyEl.querySelectorAll(`[class*="Heading${level}"]`);
            console.log('class 标题数量:', classHeadings.length);
            classHeadings.forEach(el => {
                if (!candidates.includes(el)) candidates.push(el);
            });
            if (candidates.length === 0) {
                console.log('未找到标准标题，尝试在所有元素中搜索匹配文本:', normalizedText);
                const allElements = Array.from(this.bodyEl.querySelectorAll('p, div, span, h1, h2, h3, h4, h5, h6, li, td, th'));
                let exactMatch = null;
                let partialMatch = null;
                for (const el of allElements) {
                    const elText = (el.textContent || '').trim();
                    if (!elText || elText.length > 500) continue;
                    const elTextLower = elText.toLowerCase();
                    if (elTextLower === normalizedText) {
                        exactMatch = el;
                        console.log('精确匹配找到:', { tag: el.tagName, class: el.className, text: elText.substring(0, 50) });
                        break;
                    }
                    if (!partialMatch && (elTextLower.includes(normalizedText) || normalizedText.includes(elTextLower))) {
                        if (elText.length > normalizedText.length * 0.8 && elText.length < normalizedText.length * 1.5) {
                            partialMatch = el;
                            console.log('部分匹配找到:', { tag: el.tagName, class: el.className, text: elText.substring(0, 50) });
                        }
                    }
                }
                if (exactMatch) {
                    candidates.push(exactMatch);
                } else if (partialMatch) {
                    candidates.push(partialMatch);
                }
            }
            console.log('总候选数量:', candidates.length);
            if (candidates.length > 0) {
                targetElement = candidates[0];
                console.log('使用找到的候选元素:', { tag: targetElement.tagName, class: targetElement.className, text: (targetElement.textContent || '').trim().substring(0, 50) });
            }
        }
        if (targetElement) {
            console.log('准备滚动到:', targetElement, { 
                hasViewerWrap: !!this.viewerWrap,
                viewerWrapScrollHeight: this.viewerWrap?.scrollHeight,
                viewerWrapClientHeight: this.viewerWrap?.clientHeight
            });
            requestAnimationFrame(() => {
                if (this.viewerWrap) {
                    const targetRect = targetElement.getBoundingClientRect();
                    const wrapRect = this.viewerWrap.getBoundingClientRect();
                    const currentScrollTop = this.viewerWrap.scrollTop;
                    const targetOffsetTop = targetElement.offsetTop;
                    const wrapOffsetTop = this.viewerWrap.offsetTop || 0;
                    const scrollTop = currentScrollTop + (targetRect.top - wrapRect.top) - 20;
                    console.log('滚动计算:', { 
                        currentScrollTop, 
                        targetOffsetTop, 
                        targetRectTop: targetRect.top, 
                        wrapRectTop: wrapRect.top,
                        calculatedScrollTop: scrollTop
                    });
                    this.viewerWrap.scrollTo({
                        top: Math.max(0, scrollTop),
                        behavior: 'smooth'
                    });
                } else {
                    targetElement.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    console.log('使用 scrollIntoView (无 viewerWrap)');
                }
            });
        } else {
            console.warn('scrollToHeading: 未找到目标元素', { text, tagName, outlineDataCount: this._outlineData?.length });
        }
    }

    /**
     * 计算文档信息（总字数等）。
     */
    calculateDocInfo() {
        if (!this.bodyEl) return;
        const fullText = this.bodyEl.textContent || '';
        this._docInfo.totalChars = fullText.length;
        this._docInfo.totalWords = fullText.trim().split(/\s+/).filter(w => w.length > 0).length;
        this.updateInfoDisplay();
    }

    /**
     * 更新顶部信息显示。
     */
    updateInfoDisplay() {
        if (!this._infoLabel) return;
        const info = this._docInfo;
        const wordsText = info.totalWords > 0 ? `总字数: ${info.totalWords.toLocaleString()}` : '';
        const charsText = info.totalChars > 0 ? `总字符: ${info.totalChars.toLocaleString()}` : '';
        const parts = [wordsText, charsText].filter(Boolean);
        this._infoLabel.textContent = parts.length > 0 ? parts.join(' | ') : 'DOCX 查看器';
    }

    /**
     * 设置滚动监听，更新当前页信息（DOCX 无真实分页，基于内容高度估算页数）。
     */
    setupScrollListener() {
        if (!this.viewerWrap || !this.bodyEl) return;
        const A4_HEIGHT_MM = 297;
        const MM_TO_PX = 3.779527559; // 1mm ≈ 3.78px (96 DPI)
        const A4_HEIGHT_PX = A4_HEIGHT_MM * MM_TO_PX;
        
        const updateScrollInfo = () => {
            if (!this.viewerWrap || !this._infoLabel || !this.bodyEl) return;
            const scrollTop = this.viewerWrap.scrollTop;
            const scrollHeight = this.viewerWrap.scrollHeight;
            const clientHeight = this.viewerWrap.clientHeight;
            
            const estimatedTotalPages = Math.max(1, Math.ceil(scrollHeight / A4_HEIGHT_PX));
            const currentPage = Math.max(1, Math.min(estimatedTotalPages, Math.ceil((scrollTop + clientHeight / 2) / A4_HEIGHT_PX)));
            
            const info = this._docInfo;
            const pageText = `第 ${currentPage} 页 / 共 ${estimatedTotalPages} 页`;
            const wordsText = info.totalWords > 0 ? `总字数: ${info.totalWords.toLocaleString()}` : '';
            const charsText = info.totalChars > 0 ? `总字符: ${info.totalChars.toLocaleString()}` : '';
            const parts = [pageText, wordsText, charsText].filter(Boolean);
            this._infoLabel.textContent = parts.length > 0 ? parts.join(' | ') : 'DOCX 查看器';
        };
        this.viewerWrap.addEventListener('scroll', updateScrollInfo, { passive: true });
        setTimeout(updateScrollInfo, 200);
    }

    updateStatus() {
        const bar = this.controls?.querySelector('.status-bar');
        if (bar) bar.innerHTML = '<span>DOCX 查看器（docx-preview）</span>';
    }

    destroy() {
        const root = this.previewContainer?.querySelector('.docx-viewer-container');
        if (root) root.style.display = 'none';
        this.container = null;
        this.bodyEl = null;
    }
}
