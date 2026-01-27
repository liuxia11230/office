/**
 * PDF Viewer - 混合方案
 * - 使用 PDFium.js 渲染 PDF 页面（高质量显示）
 * - 使用 PDF.js 提取书签/目录（原生支持）
 * 支持直接打开 HTML 文件进行本地预览
 */

class PDFViewer {
    constructor(options) {
        // DOM 元素
        this.previewContainer = options.previewContainer;
        this.uploadWrapper = options.uploadWrapper;
        this.controls = options.controls;
        
        // PDF 相关
        this.pdfDoc = null; // PDFium.js 文档
        this.pdfjsDoc = null; // PDF.js 文档（用于书签）
        this.pdfiumLibrary = null;
        this.pdfjsLib = null;
        this.currentPage = 1;
        this.scale = 1.0; // 默认 100%
        this.pdfContainer = null;
        this.outline = null; // 书签/目录（来自 PDF.js）
        
        // 工具栏元素
        this.toolbar = null;
        this.pageInput = null;
        this.pageCountSpan = null;
        this.zoomSelect = null;
        this.sidebarToggle = null;
        
        // 侧边栏
        this.sidebar = null;
        this.sidebarVisible = false;
        
        // 初始化 PDFium.js 和 PDF.js
        this.initPDFium();
        this.initPDFJS();
    }

    async initPDFium() {
        // 如果已经初始化，直接返回
        if (this.pdfiumLibrary) {
            return;
        }
        
        // 动态导入并初始化 PDFiumLibrary
        try {
            // 在浏览器环境中，尝试使用 CDN 版本
            // 首先尝试使用浏览器 CDN 导入
            let pdfiumModule;
            try {
                // 尝试使用浏览器 CDN 版本
                pdfiumModule = await import('https://cdn.jsdelivr.net/npm/@hyzyla/pdfium@2.1.9/dist/index.esm.cdn.js');
            } catch (importErr) {
                console.warn('CDN 导入失败，尝试标准导入:', importErr);
                // 回退到标准导入
                pdfiumModule = await import('https://cdn.jsdelivr.net/npm/@hyzyla/pdfium@2.1.9/dist/index.esm.js');
            }
            
            // 检查是否有 PDFiumLibrary 导出
            if (!pdfiumModule.PDFiumLibrary) {
                // 尝试其他可能的导出名称
                if (pdfiumModule.default && pdfiumModule.default.PDFiumLibrary) {
                    pdfiumModule = pdfiumModule.default;
                } else {
                    console.error('可用的导出:', Object.keys(pdfiumModule));
                    throw new Error('PDFiumLibrary 未找到，请检查导入的模块');
                }
            }
            
            // 初始化 PDFiumLibrary
            // 在浏览器中，可能需要提供 wasm URL
            const wasmUrl = 'https://cdn.jsdelivr.net/npm/@hyzyla/pdfium@2.1.9/dist/pdfium.wasm';
            this.pdfiumLibrary = await pdfiumModule.PDFiumLibrary.init({ 
                wasmUrl,
                disableCDNWarning: true // 禁用 CDN 警告
            });
            console.log('PDFium.js 初始化成功', this.pdfiumLibrary);
        } catch (err) {
            console.error('PDFium.js 初始化失败:', err);
            console.error('错误详情:', err.stack);
            throw new Error('PDFium.js 库未加载: ' + err.message);
        }
    }

    async initPDFJS() {
        // 等待 PDF.js 加载完成
        await this.waitForPDFJS();
        this.pdfjsLib = window.pdfjsLib;
        
        if (!this.pdfjsLib) {
            console.warn('PDF.js 库未加载，书签功能将不可用');
        }
    }

    async waitForPDFJS() {
        // 如果已经加载，直接返回
        if (window.pdfjsLib) {
            return;
        }
        
        // 等待 PDF.js 加载完成
        return new Promise((resolve) => {
            if (window.pdfjsReady) {
                resolve();
            } else {
                window.addEventListener('pdfjsReady', () => {
                    resolve();
                }, { once: true });
            }
        });
    }

    async loadFile(file) {
        // 确保 PDFium.js 已初始化
        if (!this.pdfiumLibrary) {
            await this.initPDFium();
        }
        
        try {
            // 显示预览区域，隐藏上传区域
            this.uploadWrapper.classList.add('hide');
            this.controls.classList.remove('hide');
            this.previewContainer.classList.remove('hide');
            
            // 隐藏 Excel 专用的信息栏
            const cellInfoBar = this.previewContainer.querySelector('.cell-info-bar');
            if (cellInfoBar) {
                cellInfoBar.style.display = 'none';
            }
            
            // 隐藏 table-wrapper
            const tableWrapper = this.previewContainer.querySelector('.table-wrapper');
            if (tableWrapper) {
                tableWrapper.style.display = 'none';
            }
            
            // 查找或创建 PDF viewer 容器
            let existingContainer = this.previewContainer.querySelector('.pdf-viewer-container');
            if (!existingContainer) {
                existingContainer = document.createElement('div');
                existingContainer.className = 'pdf-viewer-container';
                this.previewContainer.appendChild(existingContainer);
            } else {
                existingContainer.innerHTML = '';
                existingContainer.style.display = 'flex';
            }
            
            this.pdfContainer = existingContainer;
            
            // 创建主布局
            const mainLayout = document.createElement('div');
            mainLayout.className = 'pdf-main-layout';
            mainLayout.style.cssText = `
                display: flex;
                flex: 1;
                overflow: hidden;
            `;
            
            // 创建侧边栏
            this.createSidebar();
            mainLayout.appendChild(this.sidebar);
            
            // 创建主内容区
            const contentArea = document.createElement('div');
            contentArea.className = 'pdf-content-area';
            contentArea.style.cssText = `
                flex: 1;
                display: flex;
                flex-direction: column;
                overflow: hidden;
            `;
            
            // 创建工具栏
            this.createToolbar();
            contentArea.appendChild(this.toolbar);
            
            // 创建查看区域
            const viewerArea = document.createElement('div');
            viewerArea.className = 'pdf-viewer-area';
            viewerArea.id = 'pdf-container';
            viewerArea.style.cssText = `
                flex: 1;
                overflow: auto;
                background: #525252;
                display: flex;
                flex-direction: column;
                align-items: center;
                padding: 20px;
            `;
            this.viewerArea = viewerArea;
            contentArea.appendChild(viewerArea);
            
            mainLayout.appendChild(contentArea);
            this.pdfContainer.appendChild(mainLayout);
            
            // 加载 PDF 文件
            const arrayBuffer = await file.arrayBuffer();
            
            // 将 ArrayBuffer 转换为 Uint8Array（PDFium.js 需要）
            const uint8Array = new Uint8Array(arrayBuffer);
            
            // 验证 PDF 文件格式（检查 PDF 文件头）
            if (uint8Array.length < 4 || 
                String.fromCharCode(uint8Array[0], uint8Array[1], uint8Array[2], uint8Array[3]) !== '%PDF') {
                throw new Error('文件不是有效的 PDF 格式');
            }
            
            // 使用 PDFium.js 加载文档
            // PDFium.js 在浏览器中需要使用 Uint8Array
            let loadSuccess = false;
            try {
                console.log('尝试加载 PDF，文件大小:', uint8Array.length, 'bytes');
                console.log('PDF 文件头:', Array.from(uint8Array.slice(0, 10)).map(b => String.fromCharCode(b)).join(''));
                
                // PDFium.js 的 loadDocument 在浏览器中需要使用 Uint8Array
                this.pdfDoc = await this.pdfiumLibrary.loadDocument(uint8Array);
                
                if (!this.pdfDoc) {
                    throw new Error('loadDocument 返回 null');
                }
                
                // PDFium.js 没有 pageCount 属性，需要通过 pages() 迭代器获取
                let pageCount = 0;
                for (const page of this.pdfDoc.pages()) {
                    pageCount = page.number + 1; // page.number 是 0-based
                }
                this.pdfDoc.pageCount = pageCount; // 缓存页数
                console.log('PDF 文档加载成功，页数:', pageCount);
                this.currentPage = 1;
                loadSuccess = true;
            } catch (err) {
                console.error('PDFium.js loadDocument 错误:', err);
                console.error('错误类型:', err.constructor.name);
                console.error('错误消息:', err.message);
                
                // 如果 Uint8Array 失败，尝试 ArrayBuffer
                if (err.message && err.message.includes('PDF format')) {
                    try {
                        console.log('尝试使用 ArrayBuffer 加载...');
                        this.pdfDoc = await this.pdfiumLibrary.loadDocument(arrayBuffer);
                        if (this.pdfDoc) {
                            console.log('使用 ArrayBuffer 加载成功');
                            this.currentPage = 1;
                            loadSuccess = true;
                        } else {
                            throw new Error('loadDocument 返回 null');
                        }
                    } catch (err2) {
                        console.error('使用 ArrayBuffer 也失败:', err2);
                        // 提供更详细的错误信息
                        let errorMessage = '无法加载 PDF 文件';
                        if (err2.message) {
                            errorMessage += ': ' + err2.message;
                        } else if (err2.toString) {
                            errorMessage += ': ' + err2.toString();
                        }
                        throw new Error(errorMessage);
                    }
                } else {
                    // 提供更详细的错误信息
                    let errorMessage = '无法加载 PDF 文件';
                    if (err.message) {
                        errorMessage += ': ' + err.message;
                    } else if (err.toString) {
                        errorMessage += ': ' + err.toString();
                    }
                    throw new Error(errorMessage);
                }
            }
            
            // 只有加载成功才继续执行后续操作
            if (loadSuccess && this.pdfDoc) {
                // 同时使用 PDF.js 加载文档（用于提取书签）
                await this.loadPDFJSForOutline(arrayBuffer);
                
                // 更新页面信息
                this.updatePageInfo();
                
                // 加载书签导航（使用 PDF.js，异步，不阻塞主渲染）
                this.loadOutline().catch(err => {
                    console.error('加载书签导航失败:', err);
                });
                
                // 加载缩略图（异步，不阻塞主渲染）
                this.loadThumbnails().catch(err => {
                    console.error('加载缩略图失败:', err);
                });
                
                // 渲染所有页面（使用 PDFium.js）
                await this.renderAllPages();
                
                // 设置控制栏
                this.setupControls();
                // 右键菜单：复制
                this.setupContextMenu();
            }
            
        } catch (err) {
            console.error('PDF 加载失败:', err);
            if (this.pdfContainer) {
                this.pdfContainer.innerHTML = `
                    <div style="padding: 2rem; text-align: center; color: #333;">
                        <h3>PDF 加载失败</h3>
                        <p>${err.message}</p>
                    </div>
                `;
            }
            throw new Error('无法加载 PDF 文件: ' + err.message);
        }
    }

    createSidebar() {
        const sidebar = document.createElement('div');
        sidebar.className = 'pdf-sidebar';
        sidebar.style.cssText = `
            width: 0;
            background: #2b2b2b;
            border-right: 1px solid #444;
            overflow: hidden;
            transition: width 0.3s;
            display: flex;
            flex-direction: column;
        `;
        
        // 标签页切换
        const tabContainer = document.createElement('div');
        tabContainer.style.cssText = `
            display: flex;
            border-bottom: 1px solid #444;
            background: #333;
        `;
        
        const titleTab = document.createElement('button');
        titleTab.textContent = '标题';
        titleTab.className = 'sidebar-tab active';
        titleTab.style.cssText = `
            flex: 1;
            padding: 8px;
            background: #2b2b2b;
            border: none;
            border-bottom: 2px solid #217346;
            color: #fff;
            cursor: pointer;
            font-size: 12px;
        `;
        
        const thumbnailTab = document.createElement('button');
        thumbnailTab.textContent = '页面';
        thumbnailTab.className = 'sidebar-tab';
        thumbnailTab.style.cssText = `
            flex: 1;
            padding: 8px;
            background: #333;
            border: none;
            border-bottom: 2px solid transparent;
            color: #999;
            cursor: pointer;
            font-size: 12px;
        `;
        
        // 标签切换事件
        titleTab.addEventListener('click', () => {
            titleTab.classList.add('active');
            titleTab.style.background = '#2b2b2b';
            titleTab.style.borderBottomColor = '#217346';
            titleTab.style.color = '#fff';
            thumbnailTab.classList.remove('active');
            thumbnailTab.style.background = '#333';
            thumbnailTab.style.borderBottomColor = 'transparent';
            thumbnailTab.style.color = '#999';
            
            if (this.outlineContainer) {
                this.outlineContainer.style.display = 'flex';
            }
            if (this.thumbnailContainer) {
                this.thumbnailContainer.style.display = 'none';
            }
        });
        
        thumbnailTab.addEventListener('click', () => {
            thumbnailTab.classList.add('active');
            thumbnailTab.style.background = '#2b2b2b';
            thumbnailTab.style.borderBottomColor = '#217346';
            thumbnailTab.style.color = '#fff';
            titleTab.classList.remove('active');
            titleTab.style.background = '#333';
            titleTab.style.borderBottomColor = 'transparent';
            titleTab.style.color = '#999';
            
            if (this.outlineContainer) {
                this.outlineContainer.style.display = 'none';
            }
            if (this.thumbnailContainer) {
                this.thumbnailContainer.style.display = 'flex';
            }
        });
        
        tabContainer.appendChild(titleTab);
        tabContainer.appendChild(thumbnailTab);
        sidebar.appendChild(tabContainer);
        
        // 标题导航容器
        const outlineContainer = document.createElement('div');
        outlineContainer.className = 'pdf-outline-container';
        outlineContainer.style.cssText = `
            flex: 1;
            overflow-y: auto;
            padding: 16px;
            color: #fff;
            display: flex;
            flex-direction: column;
        `;
        
        this.outlineContainer = outlineContainer;
        sidebar.appendChild(outlineContainer);
        
        // 页面缩略图容器
        const thumbnailContainer = document.createElement('div');
        thumbnailContainer.className = 'pdf-thumbnail-container';
        thumbnailContainer.style.cssText = `
            flex: 1;
            overflow-y: auto;
            padding: 12px;
            color: #fff;
            display: none;
            flex-direction: column;
            align-items: center;
            gap: 8px;
        `;
        
        this.thumbnailContainer = thumbnailContainer;
        sidebar.appendChild(thumbnailContainer);
        
        this.sidebar = sidebar;
    }

    async loadPDFJSForOutline(arrayBuffer) {
        // 使用 PDF.js 加载文档，仅用于提取书签
        if (!this.pdfjsLib) {
            await this.waitForPDFJS();
            this.pdfjsLib = window.pdfjsLib;
        }
        
        if (!this.pdfjsLib) {
            console.warn('PDF.js 未加载，无法提取书签');
            return;
        }
        
        try {
            const loadingTask = this.pdfjsLib.getDocument({
                data: arrayBuffer,
                verbosity: 0
            });
            
            this.pdfjsDoc = await loadingTask.promise;
            console.log('PDF.js 文档加载成功（用于书签提取）');
        } catch (err) {
            console.warn('PDF.js 文档加载失败（不影响渲染）:', err);
        }
    }

    async loadOutline() {
        if (!this.outlineContainer) return;
        
        try {
            this.outlineContainer.innerHTML = '<div style="color: #999; font-size: 12px; margin-bottom: 12px;">正在加载书签...</div>';
            
            // 优先使用 PDF.js 提取书签
            if (this.pdfjsDoc) {
                try {
                    const outline = await this.pdfjsDoc.getOutline();
                    if (outline && outline.length > 0) {
                        console.log('从 PDF.js 提取到书签:', outline);
                        this.renderPDFJSOutline(outline);
                        return;
                    }
                } catch (err) {
                    console.warn('PDF.js 提取书签失败:', err);
                }
            }
            
            // 如果 PDF.js 没有书签，显示提示
            this.outlineContainer.innerHTML = '<div style="color: #999; font-size: 12px;">此 PDF 没有书签，请使用页面导航</div>';
        } catch (err) {
            console.error('加载书签导航失败:', err);
            this.outlineContainer.innerHTML = '<div style="color: #999; font-size: 12px;">无法加载书签导航</div>';
        }
    }

    /**
     * 解析书签目标并跳转到对应页面。
     * PDF.js 的 outline item.dest 可能是：字符串（命名目标）或数组（显式目标）。
     * getDestination() 仅接受字符串，传入数组会报错导致无法定位。此处按 PDF.js 官方实现处理。
     */
    async goToOutlineDest(dest) {
        if (!this.pdfjsDoc) return;
        let explicitDest;
        if (typeof dest === 'string') {
            explicitDest = await this.pdfjsDoc.getDestination(dest);
        } else if (Array.isArray(dest)) {
            explicitDest = dest;
        } else {
            console.warn('书签目标格式不支持:', dest);
            return;
        }
        if (!Array.isArray(explicitDest) || explicitDest.length === 0) {
            console.warn('书签目标解析结果无效:', explicitDest);
            return;
        }
        const destRef = explicitDest[0];
        let pageNumber;
        if (destRef && typeof destRef === 'object') {
            const pageIndex = await this.pdfjsDoc.getPageIndex(destRef);
            pageNumber = pageIndex + 1;
        } else if (Number.isInteger(destRef)) {
            pageNumber = destRef + 1;
        } else {
            console.warn('书签目标无法解析为页码:', destRef);
            return;
        }
        let pageCount = this.pdfDoc?.pageCount ?? 0;
        if (!pageCount && this.pdfDoc) {
            for (const p of this.pdfDoc.pages()) pageCount = p.number + 1;
            this.pdfDoc.pageCount = pageCount;
        }
        if (pageNumber < 1 || pageNumber > pageCount) {
            console.warn('书签页码超出范围:', pageNumber, '总页数:', pageCount);
            return;
        }
        await this.goToPage(pageNumber, explicitDest);
    }

    /**
     * 从 PDF 目标数组解析「距页面顶部的像素偏移」，用于贴顶滚动。
     * PDF 坐标系 y=0 在底部，需转为自顶向下的偏移。
     * 支持 XYZ / FitH / FitR；Fit、FitV 等无 top 的视为 0（页面顶部）。
     */
    getDestOffsetFromTop(destArray, pageDiv) {
        if (!destArray || !Array.isArray(destArray) || !pageDiv) return null;
        const fitType = destArray[1];
        const name = typeof fitType === 'object' && fitType?.name ? fitType.name : fitType;
        const h = parseFloat(pageDiv.dataset.pageHeightPoints);
        if (!Number.isFinite(h) || h <= 0) return null;
        let pdfTop = undefined;
        if (name === 'XYZ' && destArray.length >= 4) pdfTop = destArray[3];
        else if (name === 'FitH' && destArray.length >= 3) pdfTop = destArray[2];
        else if (name === 'FitR' && destArray.length >= 6) pdfTop = destArray[5];
        if (pdfTop == null || !Number.isFinite(pdfTop)) return 0;
        const pageHeightPx = h * this.scale;
        const offsetFromTop = Math.max(0, Math.min(pageHeightPx, (h - pdfTop) * this.scale));
        return offsetFromTop;
    }

    async renderPDFJSOutline(outline) {
        if (!outline || outline.length === 0) {
            this.outlineContainer.innerHTML = '<div style="color: #999; font-size: 12px;">此 PDF 没有书签</div>';
            return;
        }
        
        const ul = document.createElement('ul');
        ul.style.cssText = 'list-style: none; padding: 0; margin: 0;';
        
        for (const item of outline) {
            const li = this.createPDFJSOutlineItem(item);
            ul.appendChild(li);
        }
        
        this.outlineContainer.innerHTML = '';
        this.outlineContainer.appendChild(ul);
    }

    createPDFJSOutlineItem(item, level = 0) {
        const li = document.createElement('li');
        li.style.cssText = `
            margin: 4px 0;
            padding-left: ${level * 16}px;
        `;
        
        const a = document.createElement('a');
        a.href = '#';
        a.textContent = item.title;
        a.style.cssText = `
            color: #fff;
            text-decoration: none;
            display: block;
            padding: 6px 8px;
            border-radius: 3px;
            font-size: 13px;
            cursor: pointer;
            line-height: 1.4;
        `;
        
        a.addEventListener('mouseenter', () => {
            a.style.background = '#3a3a3a';
        });
        a.addEventListener('mouseleave', () => {
            a.style.background = 'transparent';
        });
        
        if (item.dest) {
            a.addEventListener('click', async (e) => {
                e.preventDefault();
                try {
                    await this.goToOutlineDest(item.dest);
                } catch (err) {
                    console.error('书签跳转失败:', err);
                }
            });
        }
        
        li.appendChild(a);
        
        // 处理子书签
        if (item.items && item.items.length > 0) {
            const subUl = document.createElement('ul');
            subUl.style.cssText = 'list-style: none; padding: 0; margin: 0;';
            item.items.forEach(subItem => {
                const subLi = this.createPDFJSOutlineItem(subItem, level + 1);
                subUl.appendChild(subLi);
            });
            li.appendChild(subUl);
        }
        
        return li;
    }


    async loadThumbnails() {
        if (!this.pdfDoc || !this.thumbnailContainer) return;
        
        try {
            // 清空容器
            const titleDiv = document.createElement('div');
            titleDiv.textContent = '页面导航';
            titleDiv.style.cssText = `
                color: #999;
                font-size: 12px;
                margin-bottom: 8px;
                width: 100%;
                text-align: center;
                padding-bottom: 8px;
                border-bottom: 1px solid #444;
            `;
            this.thumbnailContainer.innerHTML = '';
            this.thumbnailContainer.appendChild(titleDiv);
            
            // 获取页数
            let pageCount = this.pdfDoc.pageCount || 0;
            if (!pageCount) {
                for (const page of this.pdfDoc.pages()) {
                    pageCount = page.number + 1;
                }
                this.pdfDoc.pageCount = pageCount;
            }
            
            // 为每页创建缩略图
            let pageIndex = 0;
            for (const page of this.pdfDoc.pages()) {
                await this.createThumbnail(page, pageIndex, pageCount);
                pageIndex++;
            }
        } catch (err) {
            console.error('加载缩略图失败:', err);
            this.thumbnailContainer.innerHTML = '<div style="color: #999; font-size: 12px;">无法加载页面缩略图</div>';
        }
    }

    async createThumbnail(page, pageIndex, totalPages) {
        try {
            // 渲染小尺寸缩略图（scale=0.2，约 14 DPI）
            const thumbnail = await page.render({
                scale: 0.2,
                render: 'bitmap'
            });
            
            // 创建缩略图容器
            const thumbDiv = document.createElement('div');
            thumbDiv.className = 'pdf-thumbnail';
            thumbDiv.setAttribute('data-page', pageIndex + 1);
            thumbDiv.style.cssText = `
                width: 100%;
                max-width: 120px;
                cursor: pointer;
                padding: 6px;
                border-radius: 4px;
                background: #333;
                border: 2px solid transparent;
                transition: all 0.2s;
                display: flex;
                flex-direction: column;
                align-items: center;
            `;
            
            // 鼠标悬停效果
            thumbDiv.addEventListener('mouseenter', () => {
                thumbDiv.style.background = '#444';
                thumbDiv.style.borderColor = '#666';
            });
            thumbDiv.addEventListener('mouseleave', () => {
                if (this.currentPage !== pageIndex + 1) {
                    thumbDiv.style.background = '#333';
                    thumbDiv.style.borderColor = 'transparent';
                }
            });
            
            // 点击跳转
            thumbDiv.addEventListener('click', () => {
                this.goToPage(pageIndex + 1);
            });
            
            // 创建缩略图 canvas
            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d', { alpha: false });
            
            canvas.width = thumbnail.width;
            canvas.height = thumbnail.height;
            
            // 计算缩略图显示尺寸，保持宽高比，垂直方向
            const maxWidth = 110; // 最大宽度
            const aspectRatio = thumbnail.width / thumbnail.height;
            let displayWidth = maxWidth;
            let displayHeight = maxWidth / aspectRatio;
            
            // 如果高度太大，限制最大高度
            const maxHeight = 150;
            if (displayHeight > maxHeight) {
                displayHeight = maxHeight;
                displayWidth = maxHeight * aspectRatio;
            }
            
            canvas.style.cssText = `
                width: ${displayWidth}px;
                height: ${displayHeight}px;
                display: block;
                border-radius: 2px;
                object-fit: contain;
            `;
            
            // 绘制缩略图
            const imageData = context.createImageData(thumbnail.width, thumbnail.height);
            if (thumbnail.data.length === imageData.data.length) {
                imageData.data.set(thumbnail.data);
            } else {
                const copyLength = Math.min(thumbnail.data.length, imageData.data.length);
                imageData.data.set(thumbnail.data.subarray(0, copyLength));
            }
            context.putImageData(imageData, 0, 0);
            
            // 页面编号
            const pageLabel = document.createElement('div');
            pageLabel.textContent = `第 ${pageIndex + 1} 页`;
            pageLabel.style.cssText = `
                text-align: center;
                font-size: 11px;
                color: #ccc;
                margin-top: 6px;
                width: 100%;
            `;
            
            thumbDiv.appendChild(canvas);
            thumbDiv.appendChild(pageLabel);
            this.thumbnailContainer.appendChild(thumbDiv);
            
            // 高亮当前页
            if (this.currentPage === pageIndex + 1) {
                thumbDiv.style.background = '#4a4a4a';
                thumbDiv.style.borderColor = '#217346';
            }
            
        } catch (err) {
            console.error(`创建第 ${pageIndex + 1} 页缩略图失败:`, err);
        }
    }

    updateThumbnailHighlight() {
        if (!this.thumbnailContainer) return;
        
        // 更新所有缩略图的高亮状态
        const thumbnails = this.thumbnailContainer.querySelectorAll('.pdf-thumbnail');
        thumbnails.forEach(thumb => {
            const pageNum = parseInt(thumb.getAttribute('data-page'));
            if (pageNum === this.currentPage) {
                thumb.style.background = '#4a4a4a';
                thumb.style.borderColor = '#217346';
            } else {
                thumb.style.background = '#333';
                thumb.style.borderColor = 'transparent';
            }
        });
    }

    toggleSidebar() {
        this.sidebarVisible = !this.sidebarVisible;
        if (this.sidebarVisible) {
            this.sidebar.style.width = '250px';
        } else {
            this.sidebar.style.width = '0';
        }
    }

    createToolbar() {
        const toolbar = document.createElement('div');
        toolbar.className = 'pdf-toolbar';
        toolbar.style.cssText = `
            background: #2b2b2b;
            padding: 8px 16px;
            display: flex;
            align-items: center;
            gap: 12px;
            border-bottom: 1px solid #444;
            flex-wrap: wrap;
            flex-shrink: 0;
        `;
        
        // 侧边栏切换按钮
        this.sidebarToggle = this.createButton('☰', () => this.toggleSidebar());
        this.sidebarToggle.style.cssText += 'width: 36px; font-size: 16px;';
        toolbar.appendChild(this.sidebarToggle);
        
        // 分隔线
        toolbar.appendChild(this.createSeparator());
        
        // 上一页按钮
        const prevBtn = this.createButton('上一页', () => this.prevPage());
        prevBtn.id = 'pdf-prev-btn';
        toolbar.appendChild(prevBtn);
        
        // 下一页按钮
        const nextBtn = this.createButton('下一页', () => this.nextPage());
        nextBtn.id = 'pdf-next-btn';
        toolbar.appendChild(nextBtn);
        
        // 页面输入
        const pageGroup = document.createElement('div');
        pageGroup.style.cssText = 'display: flex; align-items: center; gap: 4px;';
        this.pageInput = document.createElement('input');
        this.pageInput.type = 'number';
        this.pageInput.min = 1;
        this.pageInput.value = 1;
        this.pageInput.style.cssText = `
            width: 50px;
            padding: 4px 8px;
            border: 1px solid #666;
            background: #fff;
            border-radius: 3px;
            text-align: center;
        `;
        this.pageInput.addEventListener('change', (e) => {
            const page = parseInt(e.target.value);
            if (this.pdfDoc) {
                // 获取页数
                let pageCount = this.pdfDoc.pageCount;
                if (!pageCount) {
                    pageCount = 0;
                    for (const p of this.pdfDoc.pages()) {
                        pageCount = p.number + 1;
                    }
                    this.pdfDoc.pageCount = pageCount;
                }
                if (page >= 1 && page <= pageCount) {
                    this.goToPage(page);
                }
            }
        });
        this.pageInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                e.target.blur();
            }
        });
        
        this.pageCountSpan = document.createElement('span');
        this.pageCountSpan.style.cssText = 'color: #fff; font-size: 12px;';
        this.pageCountSpan.textContent = ' / 1';
        
        pageGroup.appendChild(this.pageInput);
        pageGroup.appendChild(this.pageCountSpan);
        toolbar.appendChild(pageGroup);
        
        // 分隔线
        toolbar.appendChild(this.createSeparator());
        
        // 缩放控制
        const zoomGroup = document.createElement('div');
        zoomGroup.style.cssText = 'display: flex; align-items: center; gap: 4px;';
        
        const zoomOutBtn = this.createButton('−', () => this.zoomOut());
        zoomOutBtn.style.cssText += 'width: 32px;';
        zoomGroup.appendChild(zoomOutBtn);
        
        this.zoomSelect = document.createElement('select');
        this.zoomSelect.style.cssText = `
            padding: 4px 8px;
            border: 1px solid #666;
            background: #fff;
            border-radius: 3px;
            cursor: pointer;
        `;
        this.zoomSelect.innerHTML = `
            <option value="0.5">50%</option>
            <option value="0.75">75%</option>
            <option value="1" selected>100%</option>
            <option value="1.25">125%</option>
            <option value="1.5">150%</option>
            <option value="2">200%</option>
            <option value="2.5">250%</option>
            <option value="3">300%</option>
        `;
        this.zoomSelect.addEventListener('change', (e) => {
            this.scale = parseFloat(e.target.value);
            this.renderAllPages();
        });
        zoomGroup.appendChild(this.zoomSelect);
        
        const zoomInBtn = this.createButton('+', () => this.zoomIn());
        zoomInBtn.style.cssText += 'width: 32px;';
        zoomGroup.appendChild(zoomInBtn);
        
        toolbar.appendChild(zoomGroup);
        
        // 分隔线
        toolbar.appendChild(this.createSeparator());
        
        // 下载按钮
        const downloadBtn = this.createButton('下载', () => this.download());
        toolbar.appendChild(downloadBtn);
        
        // 打印按钮
        const printBtn = this.createButton('打印', () => this.print());
        toolbar.appendChild(printBtn);
        
        this.toolbar = toolbar;
    }

    createButton(text, onClick) {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.style.cssText = `
            background: #4a4a4a;
            border: 1px solid #666;
            color: #fff;
            padding: 6px 12px;
            cursor: pointer;
            border-radius: 3px;
            font-size: 12px;
        `;
        btn.addEventListener('mouseenter', () => {
            btn.style.background = '#5a5a5a';
        });
        btn.addEventListener('mouseleave', () => {
            btn.style.background = '#4a4a4a';
        });
        btn.addEventListener('click', onClick);
        return btn;
    }

    createSeparator() {
        const sep = document.createElement('div');
        sep.style.cssText = `
            width: 1px;
            height: 24px;
            background: #666;
            margin: 0 4px;
        `;
        return sep;
    }

    updatePageInfo() {
        if (this.pageInput && this.pageCountSpan && this.pdfDoc) {
            // 获取页数（如果已缓存则使用，否则重新计算）
            let pageCount = this.pdfDoc.pageCount;
            if (!pageCount) {
                pageCount = 0;
                for (const page of this.pdfDoc.pages()) {
                    pageCount = page.number + 1;
                }
                this.pdfDoc.pageCount = pageCount; // 缓存
            }
            
            this.pageInput.max = pageCount;
            this.pageInput.value = this.currentPage;
            this.pageCountSpan.textContent = ` / ${pageCount}`;
            
            // 更新按钮状态
            const prevBtn = document.getElementById('pdf-prev-btn');
            const nextBtn = document.getElementById('pdf-next-btn');
            if (prevBtn) prevBtn.disabled = this.currentPage <= 1;
            if (nextBtn) nextBtn.disabled = this.currentPage >= pageCount;
        }
    }

    async renderAllPages() {
        if (!this.pdfDoc || !this.viewerArea) return;
        
        // 清空容器
        this.viewerArea.innerHTML = '';
        
        // 使用 pages() 迭代器渲染所有页面
        console.log('开始渲染所有页面...');
        let pageIndex = 0;
        for (const page of this.pdfDoc.pages()) {
            await this.renderPage(page, pageIndex);
            pageIndex++;
        }
        console.log(`所有页面渲染完成，共 ${pageIndex} 页`);
    }

    async renderPage(page, pageIndex) {
        if (!page || pageIndex < 0) return;
        
        try {
            // page 是从 pages() 迭代器获取的页面对象
            const pageNumber = page.number; // 0-based
            console.log(`开始渲染页面 ${pageNumber + 1} (索引 ${pageIndex})...`);
            
            // 获取设备像素比，用于高清晰度渲染
            const devicePixelRatio = window.devicePixelRatio || 1;
            
            // 计算渲染缩放（考虑用户缩放和设备像素比）
            // PDFium.js 的 scale 参数：1 = 72 DPI，3 = 216 DPI（推荐用于高质量）
            const renderScale = Math.max(1, Math.floor(this.scale * devicePixelRatio * 2));
            
            console.log(`渲染参数: scale=${renderScale}, userScale=${this.scale}, devicePixelRatio=${devicePixelRatio}`);
            
            // 渲染页面为图像
            const image = await page.render({
                scale: renderScale,
                render: 'bitmap' // 返回 RGBA 数据
            });
            
            if (!image || !image.data) {
                console.error(`页面 ${pageIndex + 1} 渲染返回空数据`);
                return;
            }
            
            console.log(`页面 ${pageIndex + 1} 渲染成功:`, {
                width: image.width,
                height: image.height,
                originalWidth: image.originalWidth,
                originalHeight: image.originalHeight,
                dataLength: image.data.length
            });
            
            // 创建页面容器（存页面尺寸 point，供书签贴顶滚动用）
            const pageDiv = document.createElement('div');
            pageDiv.className = 'pdf-page';
            pageDiv.setAttribute('data-page', pageNumber + 1);
            pageDiv.dataset.pageHeightPoints = String(image.originalHeight);
            pageDiv.dataset.pageWidthPoints = String(image.originalWidth);
            pageDiv.style.cssText = `
                position: relative;
                margin-bottom: 20px;
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
                background: #fff;
                display: inline-block;
            `;
            
            // 创建 canvas
            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d', { 
                alpha: false
            });
            
            // 设置 canvas 尺寸
            // 显示尺寸（用户看到的）
            const displayWidth = image.originalWidth * this.scale;
            const displayHeight = image.originalHeight * this.scale;
            
            // Canvas 实际分辨率（高分辨率）
            canvas.width = image.width;
            canvas.height = image.height;
            
            // Canvas 显示尺寸
            canvas.style.width = displayWidth + 'px';
            canvas.style.height = displayHeight + 'px';
            
            // 将 RGBA 数据绘制到 canvas
            const imageData = context.createImageData(image.width, image.height);
            
            // 确保数据长度匹配
            if (image.data.length !== imageData.data.length) {
                console.warn(`数据长度不匹配: image.data.length=${image.data.length}, imageData.data.length=${imageData.data.length}`);
                // 只复制匹配的部分
                const copyLength = Math.min(image.data.length, imageData.data.length);
                imageData.data.set(image.data.subarray(0, copyLength));
            } else {
                imageData.data.set(image.data);
            }
            
            context.putImageData(imageData, 0, 0);
            
            pageDiv.appendChild(canvas);
            this.viewerArea.appendChild(pageDiv);
            
            await this.renderTextLayer(pageNumber + 1, pageDiv, displayWidth, displayHeight, image.originalWidth, image.originalHeight);
            
            console.log(`页面 ${pageNumber + 1} 显示完成`);
            
        } catch (err) {
            console.error(`页面 ${pageIndex + 1} 渲染失败:`, err);
            console.error('错误堆栈:', err.stack);
        }
    }

    /**
     * 文字选中功能：参考 getTextContent + 透明文本层方案
     * - 使用 PDF.js getTextContent 获取每页文本及位置
     * - 创建透明 div 覆盖在 canvas 上，按 transform 定位 span
     * - 用户可拖选文字，复制时通过 copy 事件写入剪贴板
     * @see https://juejin.cn/post/7047022519294885924
     * @see https://www.cnblogs.com/jiayouba/p/14969611.html
     */
    async renderTextLayer(pageNum, pageDiv, displayWidth, displayHeight, pageWidthPoints, pageHeightPoints) {
        if (!this.pdfjsDoc) return;
        try {
            const pdfPage = await this.pdfjsDoc.getPage(pageNum);
            const textContent = await pdfPage.getTextContent({ normalizeWhitespace: false });
            const items = textContent?.items;
            if (!items || items.length === 0) return;

            const layer = document.createElement('div');
            layer.className = 'textLayer';
            layer.setAttribute('aria-hidden', 'true');
            layer.style.cssText = `
                position: absolute; left: 0; top: 0;
                width: ${displayWidth}px; height: ${displayHeight}px;
                pointer-events: auto;
            `;

            const scale = this.scale;
            const seen = new Set();
            for (const item of items) {
                if (item.str === undefined) continue;
                const t = item.transform;
                if (!t || t.length < 6) continue;
                const x = t[4], y = t[5];
                const fs = Math.max(1, (Math.abs(t[0]) + Math.abs(t[3])) / 2 || 12);
                const fontSizePx = Math.max(4, fs * scale);
                const chars = Array.from(item.str);
                const n = chars.length;
                if (n === 0) continue;
                const w = typeof item.width === 'number' && item.width > 0 ? item.width : fs * n * 0.6;
                const advance = w / n;
                const advancePx = advance * scale;
                const ascent = 0.82;
                const topPx = (pageHeightPoints - y - fs * ascent) * scale;
                const grid = Math.max(2, advancePx * 0.75);
                const topKey = Math.round(topPx / grid) * grid;
                for (let i = 0; i < n; i++) {
                    const ch = chars[i];
                    if (this._isInvisibleChar(ch)) continue;
                    const leftPx = (x + i * advance) * scale;
                    const leftKey = Math.round(leftPx / grid) * grid;
                    const key = `${leftKey}|${topKey}|${ch}`;
                    if (seen.has(key)) continue;
                    seen.add(key);
                    const span = document.createElement('span');
                    span.textContent = ch;
                    span.style.cssText = `
                        left: ${leftPx}px; top: ${topPx}px;
                        width: ${advancePx}px; height: ${fontSizePx}px;
                        font-size: ${fontSizePx}px; font-family: inherit; line-height: 1;
                        display: inline-block; overflow: hidden; box-sizing: border-box;
                    `;
                    layer.appendChild(span);
                }
            }

            layer.addEventListener('copy', (e) => {
                const sel = document.getSelection();
                if (sel && sel.toString().length > 0) {
                    const text = this._normalizeCopyText(sel.toString());
                    if (text) e.clipboardData.setData('text/plain', text);
                    e.preventDefault();
                }
            });

            pageDiv.appendChild(layer);
        } catch (err) {
            console.warn(`文本层 第 ${pageNum} 页 渲染失败:`, err);
        }
    }

    async prevPage() {
        if (this.pdfDoc && this.currentPage > 1) {
            this.currentPage--;
            this.updatePageInfo();
            this.updateThumbnailHighlight();
            // 滚动到对应页面
            const pageDiv = this.viewerArea.querySelector(`.pdf-page[data-page="${this.currentPage}"]`);
            if (pageDiv) {
                pageDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    }

    async nextPage() {
        if (!this.pdfDoc) return;
        
        // 获取页数
        let pageCount = this.pdfDoc.pageCount;
        if (!pageCount) {
            pageCount = 0;
            for (const page of this.pdfDoc.pages()) {
                pageCount = page.number + 1;
            }
            this.pdfDoc.pageCount = pageCount;
        }
        
        if (this.currentPage < pageCount) {
            this.currentPage++;
            this.updatePageInfo();
            this.updateThumbnailHighlight();
            // 滚动到对应页面
            const pageDiv = this.viewerArea.querySelector(`.pdf-page[data-page="${this.currentPage}"]`);
            if (pageDiv) {
                pageDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    }

    /**
     * @param {number} pageNum - 1-based 页码
     * @param {Array|null} [destArray] - PDF 目标数组（书签跳转时传入），用于贴顶滚动
     */
    async goToPage(pageNum, destArray) {
        if (!this.pdfDoc) return;
        
        let pageCount = this.pdfDoc.pageCount;
        if (!pageCount) {
            pageCount = 0;
            for (const page of this.pdfDoc.pages()) {
                pageCount = page.number + 1;
            }
            this.pdfDoc.pageCount = pageCount;
        }
        
        if (pageNum < 1 || pageNum > pageCount) return;
        
        this.currentPage = pageNum;
        this.updatePageInfo();
        this.updateThumbnailHighlight();
        
        const pageDiv = this.viewerArea?.querySelector(`.pdf-page[data-page="${pageNum}"]`);
        if (!pageDiv) return;
        
        const offsetFromTop = destArray ? this.getDestOffsetFromTop(destArray, pageDiv) : null;
        const useAnchor = offsetFromTop != null && offsetFromTop > 0;
        
        requestAnimationFrame(() => {
            if (useAnchor) {
                const anchor = document.createElement('div');
                anchor.style.cssText = `
                    position: absolute; left: 0; top: ${offsetFromTop}px;
                    width: 1px; height: 1px; pointer-events: none;
                `;
                anchor.setAttribute('aria-hidden', 'true');
                pageDiv.appendChild(anchor);
                anchor.scrollIntoView({ behavior: 'smooth', block: 'start' });
                setTimeout(() => anchor.remove(), 600);
            } else {
                pageDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        });
    }

    zoomIn() {
        this.scale = Math.min(5.0, this.scale * 1.2);
        this.updateZoomSelect();
        this.renderAllPages();
    }

    zoomOut() {
        this.scale = Math.max(0.25, this.scale / 1.2);
        this.updateZoomSelect();
        this.renderAllPages();
    }

    updateZoomSelect() {
        if (this.zoomSelect) {
            // 找到最接近的选项
            const options = Array.from(this.zoomSelect.options);
            let closest = options[0];
            let minDiff = Math.abs(parseFloat(closest.value) - this.scale);
            options.forEach(opt => {
                const diff = Math.abs(parseFloat(opt.value) - this.scale);
                if (diff < minDiff) {
                    minDiff = diff;
                    closest = opt;
                }
            });
            this.zoomSelect.value = closest.value;
        }
    }

    download() {
        alert('下载功能：需要保存原始文件对象才能实现完整下载功能');
    }

    print() {
        window.print();
    }

    setupControls() {
        if (this.controls) {
            const statusBar = this.controls.querySelector('.status-bar');
            if (statusBar) {
                statusBar.innerHTML = '<span>PDF 查看器（使用 PDFium.js）</span>';
            }
        }
    }

    /**
     * 右键菜单：复制。在查看区域右击显示「复制」，有选区时点击即复制到剪贴板。
     */
    setupContextMenu() {
        if (!this.viewerArea) return;
        const handler = (e) => {
            e.preventDefault();
            this._closeContextMenu();
            const sel = document.getSelection();
            const hasSel = sel && sel.toString().length > 0;
            const selectedText = hasSel ? this._normalizeCopyText(sel.toString()) : '';
            const menu = document.createElement('div');
            menu.className = 'pdf-ctx-menu';
            menu.style.cssText = `
                position: fixed; z-index: 10000;
                left: ${e.clientX}px; top: ${e.clientY}px;
                min-width: 100px; padding: 4px 0;
                background: #fff; border: 1px solid #ddd;
                border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.15);
                font-size: 13px; font-family: inherit;
            `;
            const item = document.createElement('div');
            item.textContent = '复制';
            item.style.cssText = `
                padding: 6px 14px; cursor: ${hasSel ? 'pointer' : 'default'};
                color: ${hasSel ? '#333' : '#999'};
            `;
            item.addEventListener('mouseenter', () => {
                if (hasSel) item.style.background = 'var(--selection-bg, rgba(33,115,70,0.1))';
            });
            item.addEventListener('mouseleave', () => { item.style.background = 'transparent'; });
            item.addEventListener('mousedown', (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
            });
            item.addEventListener('click', (ev) => {
                ev.preventDefault();
                if (!hasSel || !selectedText) return;
                const doCopy = () => {
                    if (navigator.clipboard?.writeText) {
                        navigator.clipboard.writeText(selectedText).then(() => {}, () => fallbackCopy(selectedText));
                    } else {
                        fallbackCopy(selectedText);
                    }
                };
                doCopy();
                this._closeContextMenu();
            });
            menu.appendChild(item);
            document.body.appendChild(menu);
            this._ctxMenu = menu;

            const closeOnOutside = (ev) => {
                if (menu.parentNode && !menu.contains(ev.target)) {
                    this._closeContextMenu();
                    document.removeEventListener('mousedown', closeOnOutside);
                }
            };
            const raf = requestAnimationFrame(() => {
                document.addEventListener('mousedown', closeOnOutside);
            });
            this._ctxMenuClose = () => {
                cancelAnimationFrame(raf);
                document.removeEventListener('mousedown', closeOnOutside);
            };
        };
        const fallbackCopy = (text) => {
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.cssText = 'position:fixed;left:-9999px;top:0;';
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); } catch (_) {}
            ta.remove();
        };
        if (this._ctxMenuHandler) {
            this.viewerArea.removeEventListener('contextmenu', this._ctxMenuHandler);
        }
        this._ctxMenuHandler = handler;
        this.viewerArea.addEventListener('contextmenu', handler);
    }

    _closeContextMenu() {
        if (this._ctxMenuClose) {
            this._ctxMenuClose();
            this._ctxMenuClose = null;
        }
        if (this._ctxMenu?.parentNode) {
            this._ctxMenu.parentNode.removeChild(this._ctxMenu);
            this._ctxMenu = null;
        }
    }

    /** 是否为零宽/不可见字符（不建 span，避免 i、d 间多余 span 及复制异常）。 */
    _isInvisibleChar(c) {
        if (!c || c.length > 1) return false;
        const code = c.charCodeAt(0);
        if (code >= 0x200b && code <= 0x200f) return true;
        if (code === 0x2028 || code === 0x2029 || code === 0xfeff || code === 0x00ad) return true;
        if (code >= 0x2060 && code <= 0x2064) return true;
        return false;
    }

    /**
     * 复制前规范化：去掉 \\0，若为「每字重复一次」则去重（部分 PDF 同文多 item 导致 AAnnddrrooiidd）。
     */
    _normalizeCopyText(str) {
        if (!str || typeof str !== 'string') return '';
        const s = str.replace(/\u0000/g, '');
        if (s.length % 2 !== 0) return s;
        let out = '';
        for (let i = 0; i < s.length; i += 2) {
            if (s[i] !== s[i + 1]) return s;
            out += s[i];
        }
        return out;
    }

    destroy() {
        // 销毁 PDFium.js 文档
        if (this.pdfDoc) {
            try {
                this.pdfDoc.destroy();
            } catch (err) {
                console.warn('销毁 PDFium 文档时出错:', err);
            }
            this.pdfDoc = null;
        }
        
        // 销毁 PDF.js 文档
        if (this.pdfjsDoc) {
            try {
                this.pdfjsDoc.destroy();
            } catch (err) {
                console.warn('销毁 PDF.js 文档时出错:', err);
            }
            this.pdfjsDoc = null;
        }
        
        this.currentPage = 1;
        
        this._closeContextMenu();
        this._ctxMenuHandler = null;
        
        // 隐藏 PDF 容器
        const pdfContainer = this.previewContainer.querySelector('.pdf-viewer-container');
        if (pdfContainer) {
            pdfContainer.style.display = 'none';
        }
    }
}
