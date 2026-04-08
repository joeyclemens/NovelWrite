document.addEventListener('DOMContentLoaded', () => {
    let glProject = '';
    let glToken = '';
    let glBranch = 'main';
    let chapters = new Map();
    let currentFilename = null;
    let unsavedChanges = false;
    let suppressEditorChanges = false;
    let editor = null;
    const SESSION_TIMEOUT_MS = 30 * 60 * 1000;
    
    // Metadata State
    let novelMetadata = { title: '', author: '', subtitle: '', copyright: '', chapter_order: [] };
    let hasMetadataFile = false;
    let hasCoverFile = false;
    let pendingCoverBase64 = null;

    // DOM Elements - Setup
    const welcomeScreen = document.getElementById('welcome-screen');
    const inputToken = document.getElementById('gl-token');
    const authBtn = document.getElementById('auth-gitlab-btn');
    const authErrorMsg = document.getElementById('auth-error-msg');
    const projectSelectionBox = document.getElementById('project-selection');
    const glProjectSelect = document.getElementById('gl-project-select');
    const inputBranch = document.getElementById('gl-branch');
    const loadNovelBtn = document.getElementById('load-novel-btn');
    
    // DOM Elements - Editor UI
    const chapterListEl = document.getElementById('chapter-list');
    const bookTitleDisplay = document.getElementById('book-title-display');
    const chapterTitleInput = document.getElementById('chapter-title-input');
    const saveStatusEl = document.getElementById('save-status');
    const addChapterBtn = document.getElementById('add-chapter-btn');
    const addPartBtn = document.getElementById('add-part-btn');
    const manualSaveBtn = document.getElementById('manual-save-btn');
    const switchNovelBtn = document.getElementById('switch-novel-btn');
    const editMetadataBtn = document.getElementById('edit-metadata-btn');
    const editorMain = document.getElementById('editor-main');
    const focusModeBtn = document.getElementById('focus-mode-btn');
    const focusModeExitBtn = document.getElementById('focus-mode-exit');

    // DOM Elements - Metadata Modal
    const metadataModal = document.getElementById('metadata-modal');
    const cancelMetaBtn = document.getElementById('cancel-metadata-btn');
    const saveMetaBtn = document.getElementById('save-metadata-btn');
    const inputMetaTitle = document.getElementById('meta-title');
    const inputMetaAuthor = document.getElementById('meta-author');
    const inputMetaSubtitle = document.getElementById('meta-subtitle');
    const inputMetaCopyright = document.getElementById('meta-copyright');
    const metaCover = document.getElementById('meta-cover');
    
    // DOM Elements - Export & Font
    const fontSelect = document.getElementById('editor-font-select');
    const exportPdfBtn = document.getElementById('export-pdf-btn');
    const exportDocxBtn = document.getElementById('export-docx-btn');
    const exportEpubBtn = document.getElementById('export-epub-btn');

    function getSessionSnapshot() {
        return {
            token: localStorage.getItem('glToken') || '',
            project: localStorage.getItem('glProject') || '',
            branch: localStorage.getItem('glBranch') || 'main',
            lastActive: Number(localStorage.getItem('glLastActiveAt') || '0')
        };
    }

    function markSessionActivity() {
        if (!glToken || !glProject) return;
        localStorage.setItem('glLastActiveAt', String(Date.now()));
    }

    function persistSession() {
        localStorage.setItem('glToken', glToken);
        localStorage.setItem('glProject', glProject);
        localStorage.setItem('glBranch', glBranch);
        markSessionActivity();
    }

    function clearSession() {
        localStorage.removeItem('glToken');
        localStorage.removeItem('glProject');
        localStorage.removeItem('glBranch');
        localStorage.removeItem('glLastActiveAt');
    }

    function isSessionActive() {
        const { token, project, lastActive } = getSessionSnapshot();
        return !!token && !!project && (Date.now() - lastActive) < SESSION_TIMEOUT_MS;
    }

    function resetToWelcomeScreen() {
        welcomeScreen.classList.remove('hidden');
        authBtn.style.display = 'block';
        authBtn.textContent = 'Authenticate';
        inputToken.disabled = false;
        projectSelectionBox.style.display = 'none';
    }

    function activateEditorShell() {
        welcomeScreen.classList.add('hidden');
        addChapterBtn.style.display = 'flex';
        addPartBtn.style.display = 'flex';
        switchNovelBtn.style.display = 'block';
        editMetadataBtn.style.display = 'block';
        document.getElementById('export-dropdown-block').style.display = 'inline-block';
        bookTitleDisplay.textContent = novelMetadata.title || glProject.split('/').pop();
        markSessionActivity();
    }

    const editorReady = new Promise((resolve) => {
        tinymce.init({
            selector: '#editor-container',
            inline: true,
            menubar: false,
            promotion: false,
            branding: false,
            fixed_toolbar_container: '#editor-toolbar',
            toolbar_persist: true,
            resize: false,
            statusbar: false,
            min_height: 500,
            toolbar: 'undo redo | blocks | bold italic underline strikethrough | blockquote | bullist numlist outdent indent | removeformat',
            plugins: 'lists link',
            readonly: true,
            setup: (ed) => {
                ed.on('input change undo redo', () => {
                    if (!suppressEditorChanges) {
                        onContentChange();
                    }
                });
            },
            init_instance_callback: (ed) => {
                editor = ed;
                updateEditorFont(fontSelect.value || localStorage.getItem('editor-font') || "'Merriweather', serif");
                applyEditorTheme();
                resolve(ed);
            }
        });
    });

    function setEditorMode(isEditable) {
        if (editor) {
            editor.mode.set(isEditable ? 'design' : 'readonly');
        }
    }

    function setEditorContent(content) {
        if (!editor) return;
        suppressEditorChanges = true;
        editor.setContent(content || '');
        suppressEditorChanges = false;
    }

    function getEditorContent() {
        return editor ? editor.getContent() : '';
    }

    function updateEditorFont(fontFamily) {
        document.documentElement.style.setProperty('--editor-font', fontFamily);
        if (editor && editor.getBody()) {
            editor.getBody().style.fontFamily = fontFamily;
        }
    }

    function applyEditorTheme() {
        if (!editor || !editor.getBody()) return;

        const isDark = document.body.classList.contains('dark-theme');
        const body = editor.getBody();
        body.style.backgroundColor = isDark ? '#0B0E14' : '#FFFFFF';
        body.style.color = isDark ? '#F0F4F8' : '#1A202C';
    }

    function setFocusMode(enabled) {
        document.body.classList.toggle('focus-mode', enabled);
        localStorage.setItem('focus-mode', enabled ? 'true' : 'false');
        focusModeBtn.classList.toggle('active', enabled);
        focusModeBtn.title = enabled ? 'Disable Focus Mode' : 'Enable Focus Mode';
        focusModeExitBtn.title = 'Exit Focus Mode';
    }

    const savedFocusMode = localStorage.getItem('focus-mode') === 'true';
    setFocusMode(savedFocusMode);

    focusModeBtn.addEventListener('click', () => {
        setFocusMode(!document.body.classList.contains('focus-mode'));
    });

    focusModeExitBtn.addEventListener('click', () => {
        setFocusMode(false);
    });

    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape' && document.body.classList.contains('focus-mode')) {
            setFocusMode(false);
        }
    });

    // --- Drag and Drop Logic (SortableJS) ---
    Sortable.create(chapterListEl, {
        animation: 150,
        ghostClass: 'sortable-ghost',
        onEnd: async function () {
            const newOrder = [];
            document.querySelectorAll('#chapter-list li').forEach(el => {
                const name = el.querySelector('.chapter-name').textContent;
                if (el.classList.contains('divider-item')) {
                    newOrder.push(name);
                } else {
                    newOrder.push(name + '.html');
                }
            });

            novelMetadata.chapter_order = newOrder;
            setSaveStatus('Saving order...', true);
            
            try {
                const actionType = hasMetadataFile ? "update" : "create";
                await reqGL(`/repository/commits`, true, {
                    method: 'POST',
                    body: JSON.stringify({
                        branch: glBranch,
                        commit_message: `Reordered chapters via Web Editor`,
                        actions: [{
                            action: actionType,
                            file_path: "_metadata.json",
                            content: JSON.stringify(novelMetadata, null, 2)
                        }]
                    })
                });
                hasMetadataFile = true;
                setSaveStatus('Order Saved!', true);
                setTimeout(() => setSaveStatus('', false), 3000);
            } catch (err) {
                console.error(err);
                alert("Failed to save chapter order: " + err.message);
            }
        }
    });

    // --- Font Selection Logic ---
    const savedFont = localStorage.getItem('editor-font') || "'Merriweather', serif";
    fontSelect.value = savedFont;
    updateEditorFont(savedFont);

    fontSelect.addEventListener('change', () => {
        updateEditorFont(fontSelect.value);
        localStorage.setItem('editor-font', fontSelect.value);
    });

    const initialSession = getSessionSnapshot();
    if (initialSession.token) {
        inputToken.value = initialSession.token;
        inputBranch.value = initialSession.branch;
        if (!isSessionActive()) {
            clearSession();
            authErrorMsg.textContent = 'Your session expired after 30 minutes of inactivity.';
            inputToken.value = '';
            inputBranch.value = 'main';
        }
    }

    const glHeaders = () => ({
        'PRIVATE-TOKEN': glToken,
        'Content-Type': 'application/json'
    });
    
    const getEncProject = () => encodeURIComponent(glProject);

    async function reqGL(urlPath, isProjectScoped = true, options = {}) {
        markSessionActivity();
        const base = `https://gitlab.com/api/v4`;
        const url = isProjectScoped ? `${base}/projects/${getEncProject()}${urlPath}` : `${base}${urlPath}`;
        
        const res = await fetch(url, { ...options, headers: glHeaders() });
        if (!res.ok) {
            let errText = res.statusText;
            try {
                const json = await res.json();
                errText = json.message || json.error || res.statusText;
            } catch (e) {}
            throw new Error(`${res.status} - ${errText}`);
        }
        return res;
    }

    authBtn.addEventListener('click', async () => {
        glToken = inputToken.value.trim();
        if (!glToken) return;

        authBtn.textContent = 'Authenticating...';
        authErrorMsg.textContent = '';
        
        try {
            const res = await reqGL(`/projects?membership=true&simple=true&order_by=updated_at&per_page=100`, false);
            const projects = await res.json();
            
            glProjectSelect.innerHTML = '';
            projects.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p.path_with_namespace;
                opt.textContent = p.name_with_namespace;
                glProjectSelect.appendChild(opt);
            });
            
            if (projects.length === 0) {
                authErrorMsg.textContent = "No repositories found for this user.";
            } else {
                authBtn.style.display = 'none';
                inputToken.disabled = true;
                projectSelectionBox.style.display = 'flex';
                
                const prev = localStorage.getItem('glProject');
                if (prev) {
                    const exists = Array.from(glProjectSelect.options).some(o => o.value === prev);
                    if (exists) glProjectSelect.value = prev;
                }

                markSessionActivity();
            }
        } catch (err) {
            console.error(err);
            authErrorMsg.textContent = err.message;
        } finally {
            authBtn.textContent = 'Authenticate';
        }
    });

    loadNovelBtn.addEventListener('click', async () => {
        glProject = glProjectSelect.value;
        glBranch = inputBranch.value.trim() || 'main';
        
        loadNovelBtn.textContent = 'Loading...';
        
        try {
            await loadTree();
            persistSession();
            activateEditorShell();
            
        } catch (err) {
            console.error(err);
            alert("Failed to load novel: " + err.message);
        } finally {
            loadNovelBtn.textContent = 'Load Novel';
        }
    });

    switchNovelBtn.addEventListener('click', () => {
        clearSession();
        resetToWelcomeScreen();
    });

    // --- Part Divider ---
    addPartBtn.addEventListener('click', async () => {
        const partName = prompt('Enter Part Name (e.g., Part 1)');
        if (!partName || !partName.trim()) return;
        
        const partId = `DIVIDER:${partName.trim()}`;
        novelMetadata.chapter_order = novelMetadata.chapter_order || [];
        novelMetadata.chapter_order.push(partId);
        
        addChapterToList(partId);
        
        try {
            const actionType = hasMetadataFile ? "update" : "create";
            await reqGL(`/repository/commits`, true, {
                method: 'POST',
                body: JSON.stringify({
                    branch: glBranch,
                    commit_message: `Add Part Divider via Web Editor`,
                    actions: [{
                        action: actionType,
                        file_path: "_metadata.json",
                        content: JSON.stringify(novelMetadata, null, 2)
                    }]
                })
            });
            hasMetadataFile = true;
            
            const scrollBox = document.querySelector('.chapter-list-container');
            if(scrollBox) scrollBox.scrollTop = scrollBox.scrollHeight;
        } catch (e) {
            console.error(e);
            alert("Failed to save Part Divider to GitLab");
        }
    });

    // --- Metadata Modals ---
    editMetadataBtn.addEventListener('click', () => {
        inputMetaTitle.value = novelMetadata.title || '';
        inputMetaAuthor.value = novelMetadata.author || '';
        inputMetaSubtitle.value = novelMetadata.subtitle || '';
        inputMetaCopyright.value = novelMetadata.copyright || '';
        metaCover.value = ''; // clear input
        pendingCoverBase64 = null;
        metadataModal.classList.remove('hidden');
    });

    cancelMetaBtn.addEventListener('click', () => {
        metadataModal.classList.add('hidden');
    });

    metaCover.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            pendingCoverBase64 = ev.target.result; // data:image/jpeg;base64,...
        };
        reader.readAsDataURL(file);
    });

    saveMetaBtn.addEventListener('click', async () => {
        saveMetaBtn.textContent = "Syncing...";
        saveMetaBtn.disabled = true;
        
        novelMetadata.title = inputMetaTitle.value.trim();
        novelMetadata.author = inputMetaAuthor.value.trim();
        novelMetadata.subtitle = inputMetaSubtitle.value.trim();
        novelMetadata.copyright = inputMetaCopyright.value.trim();
        
        try {
            const actions = [];
            actions.push({
                action: hasMetadataFile ? "update" : "create",
                file_path: "_metadata.json",
                content: JSON.stringify(novelMetadata, null, 2)
            });

            if (pendingCoverBase64) {
                const b64Data = pendingCoverBase64.split(",")[1];
                actions.push({
                    action: hasCoverFile ? "update" : "create",
                    file_path: "_cover.jpg",
                    encoding: "base64",
                    content: b64Data
                });
            }

            await reqGL(`/repository/commits`, true, {
                method: 'POST',
                body: JSON.stringify({
                    branch: glBranch,
                    commit_message: `Update Title Page Metadata & Cover via Web Editor`,
                    actions: actions
                })
            });
            
            hasMetadataFile = true;
            hasCoverFile = hasCoverFile || !!pendingCoverBase64;
            pendingCoverBase64 = null;

            bookTitleDisplay.textContent = novelMetadata.title || glProject.split('/').pop();
            metadataModal.classList.add('hidden');
            
        } catch (err) {
            console.error(err);
            alert("Error saving metadata: " + err.message);
        } finally {
            saveMetaBtn.textContent = "Save & Sync";
            saveMetaBtn.disabled = false;
        }
    });

    async function loadTree() {
        await editorReady;

        const res = await reqGL(`/repository/tree?ref=${glBranch}&per_page=100`, true);
        const files = await res.json();
        
        chapters.clear();
        chapterCache.clear();
        chapterFetches.clear();
        chapterListEl.innerHTML = '';
        currentFilename = null;
        hasMetadataFile = false;
        hasCoverFile = false;
        
        let hasChapters = false;
        for (const file of files) {
            if (file.type === 'blob') {
                if (file.name.endsWith('.html')) {
                    hasChapters = true;
                    chapters.set(file.name, { isNew: false });
                } else if (file.name === '_metadata.json') {
                    hasMetadataFile = true;
                    try {
                        const mRes = await reqGL(`/repository/files/_metadata.json/raw?ref=${glBranch}`, true);
                        const mData = await mRes.json();
                        novelMetadata = { ...novelMetadata, ...mData };
                    } catch(e) { console.error("Failed parsing metadata", e); }
                } else if (file.name === '_cover.jpg') {
                    hasCoverFile = true;
                }
            }
        }
        
        novelMetadata.chapter_order = novelMetadata.chapter_order || [];
        const filesArray = Array.from(chapters.keys());
        
        const sortedFiles = filesArray.sort((a, b) => {
            const idxA = novelMetadata.chapter_order.indexOf(a);
            const idxB = novelMetadata.chapter_order.indexOf(b);
            if (idxA === -1 && idxB === -1) return a.localeCompare(b);
            if (idxA === -1) return 1;
            if (idxB === -1) return -1;
            return idxA - idxB;
        });

        // Insert dividers which exist in order but not in files map
        const allItemsList = [];
        novelMetadata.chapter_order.forEach(item => {
            if (item.startsWith('DIVIDER:')) allItemsList.push(item);
            else if (sortedFiles.includes(item)) allItemsList.push(item);
        });
        
        sortedFiles.forEach(file => {
            if(!allItemsList.includes(file)) allItemsList.push(file);
        });

        novelMetadata.chapter_order = allItemsList;

        allItemsList.forEach(filename => {
            addChapterToList(filename);
        });
        
        setEditorContent('');
        chapterTitleInput.value = '';
        editorMain.style.opacity = '0.3';
        editorMain.style.pointerEvents = 'none';
        chapterTitleInput.disabled = true;
        manualSaveBtn.disabled = true;
        setEditorMode(false);
        
        if (hasChapters) {
            const firstChapter = sortedFiles[0];
            if(firstChapter) {
                await switchChapter(firstChapter);
            }
        }
    }

    function addChapterToList(filename) {
        const li = document.createElement('li');
        if (filename.startsWith('DIVIDER:')) {
            const title = filename.replace('DIVIDER:', '');
            li.className = 'divider-item';
            li.innerHTML = `<span class="chapter-name" style="display:none;">${filename}</span>${title}`;
        } else {
            const title = filename.replace('.html', '');
            li.className = 'chapter-item';
            if (filename === currentFilename) li.classList.add('active');
            li.innerHTML = `<span class="chapter-name">${title}</span>`;
            li.addEventListener('click', () => switchChapter(filename));
        }
        chapterListEl.appendChild(li);
    }

    let chapterCache = new Map();
    let chapterFetches = new Map();
    let activeSwitchToken = 0;
    const PREFETCH_CONCURRENCY = 2;
    const PREFETCH_LIMIT = 4;

    function getChapterCacheKey(filename) {
        return `chapter-cache:${glProject}:${glBranch}:${filename}`;
    }

    function getCachedChapterContent(filename) {
        if (chapterCache.has(filename)) {
            return chapterCache.get(filename);
        }

        try {
            const cached = localStorage.getItem(getChapterCacheKey(filename));
            if (cached !== null) {
                chapterCache.set(filename, cached);
                return cached;
            }
        } catch (e) {
            // If storage is unavailable or full, keep using in-memory cache only.
        }

        return null;
    }

    function persistChapterContent(filename, content) {
        chapterCache.set(filename, content);

        try {
            localStorage.setItem(getChapterCacheKey(filename), content);
        } catch (e) {
            // Storage writes are best-effort.
        }
    }

    function removePersistedChapterContent(filename) {
        chapterCache.delete(filename);

        try {
            localStorage.removeItem(getChapterCacheKey(filename));
        } catch (e) {
            // Storage cleanup is best-effort.
        }
    }

    async function fetchChapterContent(filename) {
        const cachedContent = getCachedChapterContent(filename);
        if (cachedContent !== null) {
            return cachedContent;
        }

        if (chapterFetches.has(filename)) {
            return chapterFetches.get(filename);
        }

        const fetchPromise = (async () => {
            const encName = encodeURIComponent(filename);
            const res = await reqGL(`/repository/files/${encName}/raw?ref=${glBranch}`, true);
            const content = await res.text();
            persistChapterContent(filename, content);
            return content;
        })();

        chapterFetches.set(filename, fetchPromise);

        try {
            return await fetchPromise;
        } finally {
            chapterFetches.delete(filename);
        }
    }

    function getPrefetchCandidates(currentFile) {
        const orderedFiles = (novelMetadata.chapter_order || []).filter(item =>
            !item.startsWith('DIVIDER:') && item !== currentFile
        );
        const currentIndex = orderedFiles.indexOf(currentFile);

        if (currentIndex === -1) {
            return orderedFiles;
        }

        const prioritized = [];
        for (let offset = 1; offset < orderedFiles.length; offset++) {
            const nextIndex = currentIndex + offset;
            const prevIndex = currentIndex - offset;

            if (nextIndex < orderedFiles.length) {
                prioritized.push(orderedFiles[nextIndex]);
            }
            if (prevIndex >= 0) {
                prioritized.push(orderedFiles[prevIndex]);
            }
        }

        return prioritized;
    }

    async function prefetchChapters(currentFile) {
        const queue = getPrefetchCandidates(currentFile).filter(filename => {
            const chapterData = chapters.get(filename);
            return chapterData && !chapterData.isNew && getCachedChapterContent(filename) === null && !chapterFetches.has(filename);
        }).slice(0, PREFETCH_LIMIT);

        let nextIndex = 0;
        const workers = Array.from({ length: PREFETCH_CONCURRENCY }, async () => {
            while (nextIndex < queue.length) {
                const filename = queue[nextIndex];
                nextIndex += 1;

                try {
                    await fetchChapterContent(filename);
                } catch (e) {
                    // Background prefetch should never block the editor.
                }
            }
        });

        await Promise.all(workers);
    }

    async function switchChapter(filename) {
        if (currentFilename === filename) return;
        
        if (unsavedChanges) {
            const confirmLeave = confirm("You have unsaved changes! Do you want to discard them? Click Cancel to go back and Commit.");
            if (!confirmLeave) return;
        }

        const switchToken = ++activeSwitchToken;

        try {
            document.querySelectorAll('.chapter-item').forEach(el => {
                el.classList.remove('active');
                if (el.querySelector('.chapter-name').textContent === filename.replace('.html', '')) {
                    el.classList.add('active');
                }
            });

            const chapterData = chapters.get(filename);
            let content = '';
            const cachedContent = chapterData.isNew ? '' : getCachedChapterContent(filename);
            
            if (cachedContent === null) {
                editorMain.style.opacity = '0.5';
                editorMain.style.pointerEvents = 'none';
                setSaveStatus('Loading chapter...', true);
            }

            if (chapterData.isNew) {
                content = '';
            } else if (cachedContent !== null) {
                content = cachedContent;
            } else {
                content = await fetchChapterContent(filename);
            }

            if (switchToken !== activeSwitchToken) {
                return;
            }

            currentFilename = filename;
            chapterTitleInput.value = filename.replace('.html', '');
            
            editorMain.style.opacity = '1';
            editorMain.style.pointerEvents = 'all';
            chapterTitleInput.disabled = false;
            manualSaveBtn.disabled = true;
            setEditorMode(true);
            setEditorContent(content || '');
            unsavedChanges = false;
            setSaveStatus('', false);

            setTimeout(() => {
                prefetchChapters(filename);
            }, 0);
            
        } catch (error) {
            console.error(error);
            alert("Failed to load chapter.");
        }
    }

    function setSaveStatus(text, visible) {
        saveStatusEl.textContent = text;
        if (visible) saveStatusEl.classList.add('visible');
        else saveStatusEl.classList.remove('visible');
    }

    function onContentChange() {
        unsavedChanges = true;
        manualSaveBtn.disabled = false;
        setSaveStatus('Unsaved Changes', true);
    }

    chapterTitleInput.addEventListener('input', onContentChange);

    manualSaveBtn.addEventListener('click', async () => {
        if (!currentFilename) return;
        
        manualSaveBtn.disabled = true;
        setSaveStatus('Committing to GitLab...', true);
        
        const newTitle = chapterTitleInput.value.trim() || 'Untitled';
        const newFilename = `${newTitle}.html`;
        const htmlContent = getEditorContent();
        const chapterData = chapters.get(currentFilename);
        
        try {
            const actions = [];
            
            if (newFilename !== currentFilename) {
                if (!chapterData.isNew) {
                    actions.push({ action: "delete", file_path: currentFilename });
                }
                actions.push({ action: "create", file_path: newFilename, content: htmlContent });
                
                const idx = novelMetadata.chapter_order.indexOf(currentFilename);
                if (idx !== -1) novelMetadata.chapter_order[idx] = newFilename;
                else novelMetadata.chapter_order.push(newFilename);
            } else {
                actions.push({ 
                    action: chapterData.isNew ? "create" : "update", 
                    file_path: currentFilename, 
                    content: htmlContent 
                });
                if (chapterData.isNew && !novelMetadata.chapter_order.includes(newFilename)) {
                    novelMetadata.chapter_order.push(newFilename);
                }
            }

            actions.push({
                action: hasMetadataFile ? "update" : "create",
                file_path: "_metadata.json",
                content: JSON.stringify(novelMetadata, null, 2)
            });

            await reqGL(`/repository/commits`, true, {
                method: 'POST',
                body: JSON.stringify({
                    branch: glBranch,
                    commit_message: `Update ${newFilename} via Web Editor`,
                    actions: actions
                })
            });

            hasMetadataFile = true;

            if (newFilename !== currentFilename) {
                chapters.delete(currentFilename);
                removePersistedChapterContent(currentFilename);
                // Also update the DOM list
                chapterListEl.innerHTML = '';
                novelMetadata.chapter_order.forEach(f => addChapterToList(f));
            }
            
            currentFilename = newFilename;
            chapters.set(newFilename, { isNew: false });
            persistChapterContent(newFilename, htmlContent);
            unsavedChanges = false;
            setSaveStatus('Committed!', true);
            setTimeout(() => setSaveStatus('', false), 3000);
            
        } catch (err) {
            console.error(err);
            setSaveStatus('Commit failed.', true);
            manualSaveBtn.disabled = false;
            alert("Failed to commit changes: " + err.message);
        }
    });

    addChapterBtn.addEventListener('click', () => {
        if (unsavedChanges) {
            alert("Please save your current chapter before creating a new one!");
            return;
        }

        let baseName = 'New Chapter';
        let filename = `${baseName}.html`;
        let counter = 1;
        while (chapters.has(filename)) {
            filename = `${baseName} ${counter}.html`;
            counter++;
        }

        chapters.set(filename, { isNew: true });
        
        // Add instantly to end of sequence
        novelMetadata.chapter_order.push(filename);
        addChapterToList(filename);
        switchChapter(filename);
        
        // Add to DOM instantly
        const scrollBox = document.querySelector('.chapter-list-container');
        if(scrollBox) scrollBox.scrollTop = scrollBox.scrollHeight;
    });

    // --- Export Engine ---
    
    async function compileManuscript() {
        if (unsavedChanges) {
            alert("You must commit your current chapter before exporting!");
            return null;
        }

        const sortedFiles = novelMetadata.chapter_order || [];
        let compiledData = [];

        for (const filename of sortedFiles) {
            if (filename.startsWith('DIVIDER:')) {
                compiledData.push({ 
                    title: filename.replace('DIVIDER:',''), 
                    content: '', 
                    isDivider: true 
                });
            } else if (chapters.has(filename)) {
                const chapterData = chapters.get(filename);
                if (!chapterData.isNew) {
                    const encName = encodeURIComponent(filename);
                    const res = await reqGL(`/repository/files/${encName}/raw?ref=${glBranch}`, true);
                    const content = await res.text();
                    compiledData.push({ title: filename.replace('.html',''), content: content });
                }
            }
        }
        return compiledData;
    }

    function generateHTMLString(compiledChapters) {
        let htmlBlock = `<div style="text-align: center; margin-top: 40%; font-family: serif; page-break-after: always;"><h1>${novelMetadata.title || 'Untitled'}</h1>`;
        if (novelMetadata.subtitle) htmlBlock += `<h2>${novelMetadata.subtitle}</h2>`;
        if (novelMetadata.author) htmlBlock += `<br><h3>By ${novelMetadata.author}</h3>`;
        if (novelMetadata.copyright) htmlBlock += `<br><br><small>© ${novelMetadata.copyright}</small>`;
        htmlBlock += `</div>`;

        compiledChapters.forEach(ch => {
            if (ch.isDivider) {
                htmlBlock += `<div style="page-break-after: always; display:flex; flex-direction:column; align-items:center; justify-content:center; margin-top: 40vh;">`;
                htmlBlock += `<h1 style="font-size: 3rem; text-align: center;">${ch.title}</h1>`;
                htmlBlock += `</div>`;
            } else {
                htmlBlock += `<div style="page-break-after: always; font-family: ${fontSelect.value}; margin-top: 2em;">`;
                htmlBlock += `<h2 style="text-align: center; margin-bottom: 2em;">${ch.title}</h2>`;
                htmlBlock += ch.content;
                htmlBlock += `</div>`;
            }
        });
        return htmlBlock;
    }

    exportPdfBtn.addEventListener('click', async (e) => {
        try {
            e.preventDefault();
            const compiled = await compileManuscript();
            if(!compiled) return;
            
            setSaveStatus('Generating PDF...', true);
            const htmlStr = generateHTMLString(compiled);
            
            const opt = {
                margin:       1,
                filename:     `${novelMetadata.title || 'Novel'}.pdf`,
                image:        { type: 'jpeg', quality: 0.98 },
                html2canvas:  { scale: 2 },
                jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
            };
            
            const element = document.createElement('div');
            element.innerHTML = htmlStr;
            
            html2pdf().set(opt).from(element).save().then(() => {
                setSaveStatus('', false);
            });
        } catch (err) {
            console.error(err);
            alert("Export Error: " + err.message);
        }
    });

    exportDocxBtn.addEventListener('click', async (e) => {
        try {
            e.preventDefault();
            const compiled = await compileManuscript();
            if(!compiled) return;
            
            setSaveStatus('Generating DOCX...', true);
            const htmlStr = generateHTMLString(compiled);
            
            const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
            const footer = "</body></html>";
            const docHtml = header + htmlStr + footer;
            
            const blob = new Blob(['\ufeff', docHtml], { type: 'application/msword' });
            saveAs(blob, `${novelMetadata.title || 'Novel'}.doc`);
            setSaveStatus('', false);
        } catch (err) {
            console.error(err);
            alert("Export Error: " + err.message);
        }
    });

    exportEpubBtn.addEventListener('click', async (e) => {
        try {
            e.preventDefault();
            const compiledChapters = await compileManuscript();
            if(!compiledChapters) return;
            
            setSaveStatus('Fetching Cover & Compiling EPUB...', true);
            
            const zip = new JSZip();
            zip.file("mimetype", "application/epub+zip");
            
            const metaInf = zip.folder("META-INF");
            metaInf.file("container.xml", `<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
    <rootfiles>
        <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
    </rootfiles>
</container>`);

            const oebps = zip.folder("OEBPS");
            const title = novelMetadata.title || "Untitled Novel";
            const author = novelMetadata.author || "Unknown Author";
            
            let manifestItems = '';
            let spineItems = '';
            let ncxNavPoints = '';
            let playOrder = 1;

            if (hasCoverFile) {
                const coverRes = await reqGL(`/repository/files/_cover.jpg/raw?ref=${glBranch}`, true);
                const coverBlob = await coverRes.blob();
                oebps.file("Images/cover.jpg", coverBlob);
                manifestItems += `<item id="cover-image" href="Images/cover.jpg" media-type="image/jpeg" properties="cover-image"/>\n`;
                oebps.file("cover.html", `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>Cover</title></head>
<body style="margin:0;padding:0;text-align:center;">
<img src="Images/cover.jpg" alt="Cover" style="height:100%;max-width:100%;" />
</body></html>`);
                manifestItems += `<item id="cover-page" href="cover.html" media-type="application/xhtml+xml"/>\n`;
                spineItems += `<itemref idref="cover-page"/>\n`;
            }
            
            oebps.file("title.html", `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>${title}</title></head>
<body><div style="text-align: center; margin-top: 30%;">
<h1>${title}</h1><h2>By ${author}</h2>
</div></body></html>`);
            
            manifestItems += `<item id="title" href="title.html" media-type="application/xhtml+xml"/>\n`;
            spineItems += `<itemref idref="title"/>\n`;
            ncxNavPoints += `<navPoint id="navPoint-title" playOrder="${playOrder}"><navLabel><text>Title Page</text></navLabel><content src="title.html"/></navPoint>\n`;
            playOrder++;

            compiledChapters.forEach((ch, idx) => {
                const fileId = `chapter_${idx}`;
                const fileName = `${fileId}.html`;
                
                if (ch.isDivider) {
                    oebps.file(fileName, `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>${ch.title}</title></head>
<body><div style="text-align:center; margin-top:30%;"><h1>${ch.title}</h1></div></body></html>`);
                } else {
                    const parser = new DOMParser();
                    const doc = parser.parseFromString(ch.content, 'text/html');
                    let xmlContent = new XMLSerializer().serializeToString(doc.body);
                    xmlContent = xmlContent.replace(/^<body[^>]*>/i, '').replace(/<\/body>$/i, '');
                    
                    oebps.file(fileName, `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>${ch.title}</title>
<style>body { font-family: ${fontSelect.value.split(',')[0].replace(/'/g, "")}, serif; }</style>
</head>
<body><h2 style="text-align: center; margin-bottom: 2em;">${ch.title}</h2>${xmlContent}</body></html>`);
                }

                manifestItems += `<item id="${fileId}" href="${fileName}" media-type="application/xhtml+xml"/>\n`;
                spineItems += `<itemref idref="${fileId}"/>\n`;
                ncxNavPoints += `<navPoint id="nav-${fileId}" playOrder="${playOrder}"><navLabel><text>${ch.title}</text></navLabel><content src="${fileName}"/></navPoint>\n`;
                playOrder++;
            });

            oebps.file("content.opf", `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="BookID" version="2.0">
    <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
        <dc:title>${title}</dc:title>
        <dc:creator>${author}</dc:creator>
        <dc:language>en</dc:language>
        <dc:identifier id="BookID">urn:uuid:123456789</dc:identifier>
        ${hasCoverFile ? '<meta name="cover" content="cover-image"/>' : ''}
    </metadata>
    <manifest>
        <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
        ${manifestItems}
    </manifest>
    <spine toc="ncx">
        ${spineItems}
    </spine>
</package>`);

            oebps.file("toc.ncx", `<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
    <head><meta name="dtb:uid" content="urn:uuid:123456789"/><meta name="dtb:depth" content="1"/><meta name="dtb:totalPageCount" content="0"/><meta name="dtb:maxPageNumber" content="0"/></head>
    <docTitle><text>${title}</text></docTitle>
    <navMap>${ncxNavPoints}</navMap>
</ncx>`);

            const content = await zip.generateAsync({ type: "blob" });
            saveAs(content, `${title}.epub`);
            setSaveStatus('', false);

        } catch (err) {
            console.error(err);
            setSaveStatus('', false);
            alert("Export Error: " + err.message);
        }
    });

    const savedTheme = localStorage.getItem('theme') || 'dark-theme';
    document.body.className = savedTheme;
    document.getElementById('theme-toggle').addEventListener('click', () => {
        const isDark = document.body.classList.contains('dark-theme');
        document.body.className = isDark ? 'light-theme' : 'dark-theme';
        localStorage.setItem('theme', document.body.className);
        applyEditorTheme();
    });

    ['pointerdown', 'keydown', 'scroll'].forEach((eventName) => {
        document.addEventListener(eventName, () => {
            if (isSessionActive()) {
                markSessionActivity();
            }
        }, { passive: true });
    });

    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible' && isSessionActive()) {
            markSessionActivity();
        }
    });

    async function restorePreviousSession() {
        if (!isSessionActive()) return;

        const session = getSessionSnapshot();
        glToken = session.token;
        glProject = session.project;
        glBranch = session.branch;
        inputToken.value = glToken;
        inputBranch.value = glBranch;

        loadNovelBtn.textContent = 'Restoring...';
        authErrorMsg.textContent = '';

        try {
            await loadTree();
            persistSession();
            activateEditorShell();
        } catch (err) {
            console.error(err);
            clearSession();
            resetToWelcomeScreen();
            authErrorMsg.textContent = 'Session restore failed. Please sign in again.';
        } finally {
            loadNovelBtn.textContent = 'Load Novel';
        }
    }

    restorePreviousSession();
});
