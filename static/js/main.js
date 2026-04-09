document.addEventListener('DOMContentLoaded', () => {
    let glProject = '';
    let glToken = '';
    let glBranch = 'main';
    let chapters = new Map();
    let currentFilename = null;
    let unsavedChanges = false;
    let suppressEditorChanges = false;
    let editor = null;
    let glAuthMode = 'pat';
    let glRefreshToken = '';
    let glTokenExpiresAt = 0;
    let chapterWordCounts = new Map();
    let currentPersistedWordCount = 0;
    const SESSION_TIMEOUT_MS = 30 * 60 * 1000;
    const OAUTH_STORAGE_KEYS = {
        state: 'glOAuthState',
        verifier: 'glOAuthCodeVerifier'
    };
    const GITLAB_OAUTH_CONFIG = {
        instanceUrl: 'https://gitlab.com',
        clientId: 'f3df7239cba7b30d24865fe6073965fa9f02362057d5b4f51c3f946f2e635240',
        redirectUri: `${window.location.origin}${window.location.pathname}`,
        scopes: 'api read_user openid profile'
    };
    
    // Metadata State
    let novelMetadata = { title: '', author: '', subtitle: '', copyright: '', chapter_order: [] };
    let hasMetadataFile = false;
    let hasCoverFile = false;
    let pendingCoverBase64 = null;

    // DOM Elements - Setup
    const welcomeScreen = document.getElementById('welcome-screen');
    const inputToken = document.getElementById('gl-token');
    const authBtn = document.getElementById('auth-gitlab-btn');
    const oauthBtn = document.getElementById('oauth-gitlab-btn');
    const oauthHint = document.getElementById('oauth-hint');
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
    const chapterWordCountEl = document.getElementById('chapter-word-count');
    const addChapterBtn = document.getElementById('add-chapter-btn');
    const addPartBtn = document.getElementById('add-part-btn');
    const manualSaveBtn = document.getElementById('manual-save-btn');
    const switchNovelBtn = document.getElementById('switch-novel-btn');
    const logoutBtn = document.getElementById('logout-btn');
    const editMetadataBtn = document.getElementById('edit-metadata-btn');
    const editorMain = document.getElementById('editor-main');
    const focusModeToggleBtn = document.getElementById('focus-mode-toggle');
    const mobileSidebarToggleBtn = document.getElementById('mobile-sidebar-toggle');
    const mobileSidebarBackdrop = document.getElementById('mobile-sidebar-backdrop');

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
    const fontSizeSelect = document.getElementById('editor-font-size-select');
    const exportPdfBtn = document.getElementById('export-pdf-btn');
    const exportDocxBtn = document.getElementById('export-docx-btn');
    const exportEpubBtn = document.getElementById('export-epub-btn');
    const overallWordCountEl = document.getElementById('overall-word-count');
    const exportProgressOverlay = document.getElementById('export-progress-overlay');
    const exportProgressDetailEl = document.getElementById('export-progress-detail');
    const exportProgressFillEl = document.getElementById('export-progress-fill');
    const exportProgressPercentEl = document.getElementById('export-progress-percent');

    function getSessionSnapshot() {
        return {
            authMode: localStorage.getItem('glAuthMode') || 'pat',
            token: localStorage.getItem('glToken') || '',
            refreshToken: localStorage.getItem('glRefreshToken') || '',
            tokenExpiresAt: Number(localStorage.getItem('glTokenExpiresAt') || '0'),
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
        localStorage.setItem('glAuthMode', glAuthMode);
        localStorage.setItem('glToken', glToken);
        localStorage.setItem('glRefreshToken', glRefreshToken || '');
        localStorage.setItem('glTokenExpiresAt', String(glTokenExpiresAt || 0));
        localStorage.setItem('glProject', glProject);
        localStorage.setItem('glBranch', glBranch);
        markSessionActivity();
    }

    function clearSession() {
        localStorage.removeItem('glAuthMode');
        localStorage.removeItem('glToken');
        localStorage.removeItem('glRefreshToken');
        localStorage.removeItem('glTokenExpiresAt');
        localStorage.removeItem('glProject');
        localStorage.removeItem('glBranch');
        localStorage.removeItem('glLastActiveAt');
        localStorage.removeItem(OAUTH_STORAGE_KEYS.state);
        localStorage.removeItem(OAUTH_STORAGE_KEYS.verifier);
    }

    function isSessionActive() {
        const { token, lastActive } = getSessionSnapshot();
        return !!token && (Date.now() - lastActive) < SESSION_TIMEOUT_MS;
    }

    function resetToWelcomeScreen() {
        welcomeScreen.classList.remove('hidden');
        authBtn.style.display = 'block';
        authBtn.textContent = 'Authenticate';
        oauthBtn.disabled = !GITLAB_OAUTH_CONFIG.clientId;
        inputToken.disabled = false;
        inputToken.placeholder = 'Personal Access Token';
        projectSelectionBox.style.display = 'none';
    }

    function deactivateEditorShell() {
        welcomeScreen.classList.remove('hidden');
        addChapterBtn.style.display = 'none';
        addPartBtn.style.display = 'none';
        switchNovelBtn.style.display = 'none';
        logoutBtn.style.display = 'none';
        editMetadataBtn.style.display = 'none';
        document.getElementById('export-dropdown-block').style.display = 'none';
        bookTitleDisplay.textContent = 'Novel Workspace';
        metadataModal.classList.add('hidden');
        chapterListEl.innerHTML = '';
        chapters.clear();
        chapterCache.clear();
        chapterFetches.clear();
        chapterWordCounts.clear();
        novelMetadata = { title: '', author: '', subtitle: '', copyright: '', chapter_order: [] };
        hasMetadataFile = false;
        hasCoverFile = false;
        pendingCoverBase64 = null;
        currentFilename = null;
        currentPersistedWordCount = 0;
        unsavedChanges = false;
        chapterTitleInput.value = '';
        chapterTitleInput.disabled = true;
        manualSaveBtn.disabled = true;
        setEditorMode(false);
        setEditorContent('');
        setSaveStatus('', false);
        editorMain.style.opacity = '0.3';
        editorMain.style.pointerEvents = 'none';
        updateChapterWordCountDisplay(0);
        updateOverallWordCountDisplay(0);
    }

    function activateEditorShell() {
        welcomeScreen.classList.add('hidden');
        addChapterBtn.style.display = 'flex';
        addPartBtn.style.display = 'flex';
        switchNovelBtn.style.display = 'block';
        logoutBtn.style.display = 'block';
        editMetadataBtn.style.display = 'block';
        document.getElementById('export-dropdown-block').style.display = 'inline-block';
        bookTitleDisplay.textContent = novelMetadata.title || glProject.split('/').pop();
        markSessionActivity();
        renderOverallWordCount();
    }

    function isOAuthConfigured() {
        return !!GITLAB_OAUTH_CONFIG.clientId;
    }

    function getGitLabOAuthUrl(path) {
        return `${GITLAB_OAUTH_CONFIG.instanceUrl}${path}`;
    }

    function randomString(length = 64) {
        const bytes = new Uint8Array(length);
        crypto.getRandomValues(bytes);
        return Array.from(bytes, b => ('0' + (b % 36).toString(36)).slice(-1)).join('');
    }

    function toBase64Url(bytes) {
        return btoa(String.fromCharCode(...bytes))
            .replace(/\+/g, '-')
            .replace(/\//g, '_')
            .replace(/=+$/g, '');
    }

    async function sha256(input) {
        const encoded = new TextEncoder().encode(input);
        return new Uint8Array(await crypto.subtle.digest('SHA-256', encoded));
    }

    async function createPkceChallenge(verifier) {
        return toBase64Url(await sha256(verifier));
    }

    async function exchangeOAuthToken(params) {
        const response = await fetch(getGitLabOAuthUrl('/oauth/token'), {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams(params)
        });

        if (!response.ok) {
            let errorMessage = response.statusText;
            try {
                const payload = await response.json();
                errorMessage = payload.error_description || payload.error || response.statusText;
            } catch (e) {}
            throw new Error(errorMessage);
        }

        return response.json();
    }

    async function ensureValidAccessToken() {
        if (glAuthMode !== 'oauth') return;
        if (!glToken) throw new Error('Missing GitLab access token.');

        const expiresSoon = !glTokenExpiresAt || (Date.now() + 30_000) >= glTokenExpiresAt;
        if (!expiresSoon) return;
        if (!glRefreshToken) throw new Error('GitLab session expired. Please sign in again.');

        const tokenData = await exchangeOAuthToken({
            grant_type: 'refresh_token',
            client_id: GITLAB_OAUTH_CONFIG.clientId,
            refresh_token: glRefreshToken,
            redirect_uri: GITLAB_OAUTH_CONFIG.redirectUri
        });

        glToken = tokenData.access_token;
        glRefreshToken = tokenData.refresh_token || glRefreshToken;
        glTokenExpiresAt = Date.now() + ((tokenData.expires_in || 7200) * 1000);
        persistSession();
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
                updateEditorFontSize(fontSizeSelect.value || localStorage.getItem('editor-font-size') || '1.15rem');
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

    function countWordsFromHtml(html) {
        const scratch = document.createElement('div');
        scratch.innerHTML = html || '';
        const text = (scratch.textContent || scratch.innerText || '').replace(/\s+/g, ' ').trim();
        return text ? text.split(' ').length : 0;
    }

    function formatWordCount(count) {
        return new Intl.NumberFormat().format(count);
    }

    function updateChapterWordCountDisplay(count) {
        chapterWordCountEl.textContent = `Chapter: ${formatWordCount(count)} words`;
    }

    function updateOverallWordCountDisplay(count) {
        overallWordCountEl.textContent = `Total: ${formatWordCount(count)} words`;
    }

    function renderCurrentChapterWordCount() {
        updateChapterWordCountDisplay(countWordsFromHtml(getEditorContent()));
    }

    function renderOverallWordCount() {
        const total = Array.from(chapterWordCounts.values()).reduce((sum, value) => sum + value, 0);
        updateOverallWordCountDisplay(total);
    }

    async function refreshOverallWordCount() {
        overallWordCountEl.textContent = 'Total: Calculating...';

        const files = (novelMetadata.chapter_order || []).filter(filename =>
            !filename.startsWith('DIVIDER:') && chapters.has(filename)
        );

        const counts = await Promise.all(files.map(async (filename) => {
            const chapterData = chapters.get(filename);
            if (!chapterData || chapterData.isNew) {
                return [filename, 0];
            }

            const content = await fetchChapterContent(filename);
            return [filename, countWordsFromHtml(content)];
        }));

        chapterWordCounts = new Map(counts);
        currentPersistedWordCount = chapterWordCounts.get(currentFilename) || 0;
        renderOverallWordCount();
    }

    function updateEditorFont(fontFamily) {
        document.documentElement.style.setProperty('--editor-font', fontFamily);
        if (editor && editor.getBody()) {
            editor.getBody().style.fontFamily = fontFamily;
        }
    }

    function updateEditorFontSize(fontSize) {
        document.documentElement.style.setProperty('--editor-font-size', fontSize);
        if (editor && editor.getBody()) {
            editor.getBody().style.fontSize = fontSize;
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
        focusModeToggleBtn.classList.toggle('active', enabled);
        focusModeToggleBtn.title = enabled ? 'Disable Focus Mode' : 'Enable Focus Mode';
    }

    function isMobileViewport() {
        return window.innerWidth <= 900;
    }

    function setSidebarOpen(enabled) {
        document.body.classList.toggle('sidebar-open', enabled && isMobileViewport());
        mobileSidebarToggleBtn.classList.toggle('active', enabled && isMobileViewport());
        mobileSidebarToggleBtn.title = enabled && isMobileViewport() ? 'Close Chapters' : 'Open Chapters';
    }

    const savedFocusMode = localStorage.getItem('focus-mode') === 'true';
    setFocusMode(savedFocusMode);
    setSidebarOpen(false);

    focusModeToggleBtn.addEventListener('click', () => {
        setFocusMode(!document.body.classList.contains('focus-mode'));
    });

    mobileSidebarToggleBtn.addEventListener('click', () => {
        setSidebarOpen(!document.body.classList.contains('sidebar-open'));
    });

    mobileSidebarBackdrop.addEventListener('click', () => {
        setSidebarOpen(false);
    });

    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape' && document.body.classList.contains('focus-mode')) {
            setFocusMode(false);
        }

        if (event.key === 'Escape' && document.body.classList.contains('sidebar-open')) {
            setSidebarOpen(false);
        }
    });

    window.addEventListener('resize', () => {
        if (!isMobileViewport()) {
            setSidebarOpen(false);
        }
    });

    document.addEventListener('keydown', (event) => {
        const isSaveShortcut = (event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 's';
        if (!isSaveShortcut) return;

        event.preventDefault();

        if (currentFilename && !manualSaveBtn.disabled) {
            manualSaveBtn.click();
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
    const savedFontSize = localStorage.getItem('editor-font-size') || '1.15rem';
    fontSelect.value = savedFont;
    fontSizeSelect.value = savedFontSize;
    updateEditorFont(savedFont);
    updateEditorFontSize(savedFontSize);

    fontSelect.addEventListener('change', () => {
        updateEditorFont(fontSelect.value);
        localStorage.setItem('editor-font', fontSelect.value);
    });

    fontSizeSelect.addEventListener('change', () => {
        updateEditorFontSize(fontSizeSelect.value);
        localStorage.setItem('editor-font-size', fontSizeSelect.value);
    });

    const glHeaders = () => ({
        ...(glAuthMode === 'oauth'
            ? { 'Authorization': `Bearer ${glToken}` }
            : { 'PRIVATE-TOKEN': glToken }),
        'Content-Type': 'application/json'
    });
    
    const getEncProject = () => encodeURIComponent(glProject);

    async function reqGL(urlPath, isProjectScoped = true, options = {}) {
        await ensureValidAccessToken();
        markSessionActivity();
        const base = `${GITLAB_OAUTH_CONFIG.instanceUrl}/api/v4`;
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

    async function populateProjectSelection() {
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
            return false;
        }

        authBtn.style.display = 'none';
        inputToken.disabled = glAuthMode === 'oauth';
        projectSelectionBox.style.display = 'flex';

        const prev = localStorage.getItem('glProject');
        if (prev) {
            const exists = Array.from(glProjectSelect.options).some(o => o.value === prev);
            if (exists) glProjectSelect.value = prev;
        }

        markSessionActivity();
        return true;
    }

    async function beginGitLabOAuth() {
        if (!isOAuthConfigured()) {
            authErrorMsg.textContent = 'Add your GitLab OAuth client ID to enable Sign in with GitLab.';
            return;
        }

        const state = randomString(32);
        const verifier = randomString(96);
        const challenge = await createPkceChallenge(verifier);

        localStorage.setItem(OAUTH_STORAGE_KEYS.state, state);
        localStorage.setItem(OAUTH_STORAGE_KEYS.verifier, verifier);

        const authUrl = new URL(getGitLabOAuthUrl('/oauth/authorize'));
        authUrl.searchParams.set('client_id', GITLAB_OAUTH_CONFIG.clientId);
        authUrl.searchParams.set('redirect_uri', GITLAB_OAUTH_CONFIG.redirectUri);
        authUrl.searchParams.set('response_type', 'code');
        authUrl.searchParams.set('scope', GITLAB_OAUTH_CONFIG.scopes);
        authUrl.searchParams.set('state', state);
        authUrl.searchParams.set('code_challenge', challenge);
        authUrl.searchParams.set('code_challenge_method', 'S256');

        window.location.assign(authUrl.toString());
    }

    async function handleOAuthCallback() {
        const url = new URL(window.location.href);
        const code = url.searchParams.get('code');
        const state = url.searchParams.get('state');
        const oauthError = url.searchParams.get('error');

        if (oauthError) {
            authErrorMsg.textContent = url.searchParams.get('error_description') || oauthError;
            url.search = '';
            window.history.replaceState({}, document.title, url.toString());
            return false;
        }

        if (!code) return false;

        const storedState = localStorage.getItem(OAUTH_STORAGE_KEYS.state);
        const verifier = localStorage.getItem(OAUTH_STORAGE_KEYS.verifier);
        url.search = '';
        window.history.replaceState({}, document.title, url.toString());

        if (!storedState || !verifier || storedState !== state) {
            authErrorMsg.textContent = 'GitLab sign-in could not be verified. Please try again.';
            localStorage.removeItem(OAUTH_STORAGE_KEYS.state);
            localStorage.removeItem(OAUTH_STORAGE_KEYS.verifier);
            return true;
        }

        oauthBtn.textContent = 'Signing in...';
        authErrorMsg.textContent = '';

        try {
            const tokenData = await exchangeOAuthToken({
                grant_type: 'authorization_code',
                client_id: GITLAB_OAUTH_CONFIG.clientId,
                code,
                code_verifier: verifier,
                redirect_uri: GITLAB_OAUTH_CONFIG.redirectUri
            });

            glAuthMode = 'oauth';
            glToken = tokenData.access_token;
            glRefreshToken = tokenData.refresh_token || '';
            glTokenExpiresAt = Date.now() + ((tokenData.expires_in || 7200) * 1000);
            persistSession();
            await populateProjectSelection();
        } catch (err) {
            console.error(err);
            authErrorMsg.textContent = err.message;
        } finally {
            localStorage.removeItem(OAUTH_STORAGE_KEYS.state);
            localStorage.removeItem(OAUTH_STORAGE_KEYS.verifier);
            oauthBtn.textContent = 'Sign In with GitLab';
        }

        return true;
    }

    authBtn.addEventListener('click', async () => {
        glAuthMode = 'pat';
        glRefreshToken = '';
        glTokenExpiresAt = 0;
        glToken = inputToken.value.trim();
        inputToken.placeholder = 'Personal Access Token';
        if (!glToken) return;

        authBtn.textContent = 'Authenticating...';
        authErrorMsg.textContent = '';
        
        try {
            await populateProjectSelection();
        } catch (err) {
            console.error(err);
            authErrorMsg.textContent = err.message;
        } finally {
            authBtn.textContent = 'Authenticate';
        }
    });

    oauthBtn.addEventListener('click', async () => {
        authErrorMsg.textContent = '';
        try {
            await beginGitLabOAuth();
        } catch (err) {
            console.error(err);
            authErrorMsg.textContent = err.message;
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

    switchNovelBtn.addEventListener('click', async () => {
        if (unsavedChanges) {
            const confirmSwitch = confirm("You have unsaved changes. Switch novels anyway?");
            if (!confirmSwitch) return;
        }

        setSidebarOpen(false);
        glProject = '';
        glBranch = 'main';
        localStorage.removeItem('glProject');
        localStorage.removeItem('glBranch');
        localStorage.removeItem('glLastActiveAt');
        inputBranch.value = 'main';
        deactivateEditorShell();
        resetToWelcomeScreen();

        try {
            await populateProjectSelection();
        } catch (err) {
            console.error(err);
            authErrorMsg.textContent = 'Could not reload your projects. Please sign in again.';
            clearSession();
            glToken = '';
            glRefreshToken = '';
            glTokenExpiresAt = 0;
            glAuthMode = 'pat';
            inputToken.value = '';
            resetToWelcomeScreen();
        }
    });

    logoutBtn.addEventListener('click', () => {
        if (unsavedChanges) {
            const confirmLogout = confirm("You have unsaved changes. Log out anyway?");
            if (!confirmLogout) return;
        }

        setSidebarOpen(false);
        deactivateEditorShell();
        clearSession();
        glProject = '';
        glToken = '';
        glRefreshToken = '';
        glTokenExpiresAt = 0;
        glAuthMode = 'pat';
        glBranch = 'main';
        inputToken.value = '';
        inputToken.disabled = false;
        inputToken.placeholder = 'Personal Access Token';
        inputBranch.value = 'main';
        authErrorMsg.textContent = '';
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
        chapterWordCounts.clear();
        chapterListEl.innerHTML = '';
        currentFilename = null;
        currentPersistedWordCount = 0;
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

            refreshOverallWordCount();
        } else {
            updateChapterWordCountDisplay(0);
            updateOverallWordCountDisplay(0);
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
            li.addEventListener('click', () => {
                switchChapter(filename);
                setSidebarOpen(false);
            });
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
        chapterWordCounts.set(filename, countWordsFromHtml(content));

        try {
            localStorage.setItem(getChapterCacheKey(filename), content);
        } catch (e) {
            // Storage writes are best-effort.
        }
    }

    function removePersistedChapterContent(filename) {
        chapterCache.delete(filename);
        chapterWordCounts.delete(filename);

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
            currentPersistedWordCount = countWordsFromHtml(content || '');
            chapterTitleInput.value = filename.replace('.html', '');
            
            editorMain.style.opacity = '1';
            editorMain.style.pointerEvents = 'all';
            chapterTitleInput.disabled = false;
            manualSaveBtn.disabled = true;
            setEditorMode(true);
            setEditorContent(content || '');
            unsavedChanges = false;
            setSaveStatus('', false);
            updateChapterWordCountDisplay(currentPersistedWordCount);

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

    function showExportProgress(detailText, percent = 0) {
        exportProgressDetailEl.textContent = detailText;
        exportProgressFillEl.style.width = `${Math.max(0, Math.min(100, percent))}%`;
        exportProgressPercentEl.textContent = `${Math.round(Math.max(0, Math.min(100, percent)))}%`;
        exportProgressOverlay.classList.remove('hidden');
    }

    function hideExportProgress() {
        exportProgressOverlay.classList.add('hidden');
        exportProgressDetailEl.textContent = 'Preparing your manuscript...';
        exportProgressFillEl.style.width = '0%';
        exportProgressPercentEl.textContent = '0%';
    }

    function onContentChange() {
        unsavedChanges = true;
        manualSaveBtn.disabled = false;
        setSaveStatus('Unsaved Changes', true);
        renderCurrentChapterWordCount();

        if (currentFilename) {
            chapterWordCounts.set(currentFilename, countWordsFromHtml(getEditorContent()));
            renderOverallWordCount();
        }
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
            currentPersistedWordCount = countWordsFromHtml(htmlContent);
            unsavedChanges = false;
            setSaveStatus('Committed!', true);
            updateChapterWordCountDisplay(currentPersistedWordCount);
            renderOverallWordCount();
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

    function htmlToPdfParagraphs(html) {
        const container = document.createElement('div');
        container.innerHTML = html || '';
        const blocks = [];
        const blockSelector = 'h1, h2, h3, h4, h5, h6, p, li, blockquote, div';

        container.querySelectorAll(blockSelector).forEach((node) => {
            if (node.children.length > 0 && node.tagName.toLowerCase() === 'div') return;

            const text = (node.textContent || '').replace(/\s+/g, ' ').trim();
            if (!text) return;

            const tag = node.tagName.toLowerCase();
            if (tag === 'li') {
                blocks.push(`• ${text}`);
            } else {
                blocks.push(text);
            }
        });

        if (blocks.length === 0) {
            const fallback = (container.textContent || '').replace(/\s+/g, ' ').trim();
            if (fallback) blocks.push(fallback);
        }

        return blocks;
    }

    function addPdfWrappedText(doc, text, x, y, maxWidth, lineHeight) {
        const lines = doc.splitTextToSize(text, maxWidth);
        doc.text(lines, x, y);
        return y + (lines.length * lineHeight);
    }

    function getJsPdfConstructor() {
        if (window.jspdf && window.jspdf.jsPDF) {
            return window.jspdf.jsPDF;
        }

        if (window.jsPDF) {
            return window.jsPDF;
        }

        throw new Error('PDF export library is not available on this page.');
    }

    function escapeXml(value) {
        return String(value ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    function createEpubUuid() {
        if (window.crypto && typeof window.crypto.randomUUID === 'function') {
            return `urn:uuid:${window.crypto.randomUUID()}`;
        }

        return `urn:uuid:${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
    }

    async function detectImageFormat(blob) {
        const bytes = new Uint8Array(await blob.slice(0, 16).arrayBuffer());

        if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) {
            return { extension: 'png', mediaType: 'image/png' };
        }

        if (bytes[0] === 0xFF && bytes[1] === 0xD8 && bytes[2] === 0xFF) {
            return { extension: 'jpg', mediaType: 'image/jpeg' };
        }

        return { extension: 'jpg', mediaType: 'image/jpeg' };
    }

    function buildEpubXhtml(title, bodyMarkup, extraHeadMarkup = '') {
        return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>${escapeXml(title)}</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
${extraHeadMarkup}
</head>
<body>${bodyMarkup}</body>
</html>`;
    }

    function convertHtmlToXhtmlFragment(html) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html || '', 'text/html');
        let xmlContent = new XMLSerializer().serializeToString(doc.body);
        return xmlContent.replace(/^<body[^>]*>/i, '').replace(/<\/body>$/i, '');
    }

    exportPdfBtn.addEventListener('click', async (e) => {
        try {
            e.preventDefault();
            const compiled = await compileManuscript();
            if(!compiled) return;
            
            showExportProgress('Rendering PDF pages...', 20);
            setSaveStatus('Generating PDF...', true);
            const jsPDF = getJsPdfConstructor();
            const doc = new jsPDF({ unit: 'in', format: 'letter', orientation: 'portrait' });
            const pageWidth = 8.5;
            const pageHeight = 11;
            const marginX = 1;
            const marginTop = 1;
            const marginBottom = 1;
            const maxWidth = pageWidth - (marginX * 2);
            const bodyFont = fontSelect.value.includes('sans') ? 'helvetica' : 'times';
            const title = novelMetadata.title || 'Untitled';
            const author = novelMetadata.author || '';
            let y = 2.5;

            doc.setFont(bodyFont, 'bold');
            doc.setFontSize(24);
            y = addPdfWrappedText(doc, title, marginX, y, maxWidth, 0.32);

            if (novelMetadata.subtitle) {
                doc.setFont(bodyFont, 'normal');
                doc.setFontSize(16);
                y += 0.15;
                y = addPdfWrappedText(doc, novelMetadata.subtitle, marginX, y, maxWidth, 0.24);
            }

            if (author) {
                doc.setFont(bodyFont, 'normal');
                doc.setFontSize(14);
                y += 0.25;
                y = addPdfWrappedText(doc, `By ${author}`, marginX, y, maxWidth, 0.22);
            }

            if (novelMetadata.copyright) {
                doc.setFont(bodyFont, 'normal');
                doc.setFontSize(10);
                y += 0.35;
                y = addPdfWrappedText(doc, `© ${novelMetadata.copyright}`, marginX, y, maxWidth, 0.18);
            }

            compiled.forEach((chapter, index) => {
                doc.addPage();
                showExportProgress(`Rendering PDF section ${index + 1} of ${compiled.length}...`, 20 + (((index + 1) / Math.max(compiled.length, 1)) * 70));
                y = marginTop + 0.35;

                if (chapter.isDivider) {
                    doc.setFont(bodyFont, 'bold');
                    doc.setFontSize(24);
                    addPdfWrappedText(doc, chapter.title, marginX, pageHeight / 2, maxWidth, 0.32);
                    return;
                }

                doc.setFont(bodyFont, 'bold');
                doc.setFontSize(18);
                y = addPdfWrappedText(doc, chapter.title, marginX, y, maxWidth, 0.26);
                y += 0.2;

                doc.setFont(bodyFont, 'normal');
                doc.setFontSize(12);

                const paragraphs = htmlToPdfParagraphs(chapter.content);
                paragraphs.forEach((paragraph) => {
                    const lines = doc.splitTextToSize(paragraph, maxWidth);
                    const paragraphHeight = lines.length * 0.22;

                    if (y + paragraphHeight > pageHeight - marginBottom) {
                        doc.addPage();
                        y = marginTop;
                        doc.setFont(bodyFont, 'normal');
                        doc.setFontSize(12);
                    }

                    doc.text(lines, marginX, y);
                    y += paragraphHeight + 0.08;
                });
            });

            doc.save(`${title}.pdf`);
            showExportProgress('Saving PDF...', 100);
            setSaveStatus('', false);
            setTimeout(() => hideExportProgress(), 300);
        } catch (err) {
            console.error(err);
            setSaveStatus('', false);
            hideExportProgress();
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
            showExportProgress('Collecting chapters from GitLab...', 8);
            const compiledChapters = await compileManuscript();
            if(!compiledChapters) {
                hideExportProgress();
                return;
            }
            
            showExportProgress('Building EPUB package...', 18);
            setSaveStatus('Fetching Cover & Compiling EPUB...', true);
            
            const zip = new JSZip();
            zip.file("mimetype", "application/epub+zip", { compression: "STORE" });
            
            const metaInf = zip.folder("META-INF");
            metaInf.file("container.xml", `<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
    <rootfiles>
        <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
    </rootfiles>
</container>`);

            const oebps = zip.folder("OEBPS");
            const title = novelMetadata.title || "Untitled Novel";
            const author = novelMetadata.author || "Unknown Author";
            const bookId = createEpubUuid();
            const safeTitle = escapeXml(title);
            const safeAuthor = escapeXml(author);
            const safeSubtitle = escapeXml(novelMetadata.subtitle || '');
            const safeCopyright = escapeXml(novelMetadata.copyright || '');
            
            let manifestItems = '';
            let spineItems = '';
            let ncxNavPoints = '';
            let playOrder = 1;
            let guideItems = '<reference type="text" title="Start" href="title.xhtml"/>\n';
            let coverPageHref = '';

            if (hasCoverFile) {
                showExportProgress('Adding cover image...', 28);
                const coverRes = await reqGL(`/repository/files/_cover.jpg/raw?ref=${glBranch}`, true);
                const coverBlob = await coverRes.blob();
                const coverFormat = await detectImageFormat(coverBlob);
                const coverImageHref = `Images/cover.${coverFormat.extension}`;
                coverPageHref = 'cover.xhtml';

                oebps.file(coverImageHref, coverBlob);
                manifestItems += `<item id="cover-image" href="${coverImageHref}" media-type="${coverFormat.mediaType}"/>\n`;
                oebps.file(coverPageHref, buildEpubXhtml(
                    'Cover',
                    `<div style="margin:0;padding:0;text-align:center;"><img src="${coverImageHref}" alt="Cover" style="display:block;height:auto;max-width:100%;margin:0 auto;"/></div>`
                ));
                manifestItems += `<item id="cover-page" href="${coverPageHref}" media-type="application/xhtml+xml"/>\n`;
                spineItems += `<itemref idref="cover-page"/>\n`;
                guideItems = `<reference type="cover" title="Cover" href="${coverPageHref}"/>\n${guideItems}`;
            }
            
            showExportProgress('Creating title page...', 38);
            oebps.file("title.xhtml", buildEpubXhtml(
                title,
                `<div style="text-align:center;margin-top:30%;">
<h1>${safeTitle}</h1>${safeSubtitle ? `<h2>${safeSubtitle}</h2>` : ''}<h2>By ${safeAuthor}</h2>${safeCopyright ? `<p>${safeCopyright}</p>` : ''}
</div>`
            ));
            
            manifestItems += `<item id="title" href="title.xhtml" media-type="application/xhtml+xml"/>\n`;
            spineItems += `<itemref idref="title"/>\n`;
            ncxNavPoints += `<navPoint id="navPoint-title" playOrder="${playOrder}"><navLabel><text>Title Page</text></navLabel><content src="title.xhtml"/></navPoint>\n`;
            playOrder++;

            compiledChapters.forEach((ch, idx) => {
                const chapterProgress = 40 + (((idx + 1) / Math.max(compiledChapters.length, 1)) * 40);
                showExportProgress(`Formatting chapter ${idx + 1} of ${compiledChapters.length}...`, chapterProgress);
                const fileId = `chapter_${idx}`;
                const fileName = `${fileId}.xhtml`;
                const safeChapterTitle = escapeXml(ch.title);
                
                if (ch.isDivider) {
                    oebps.file(fileName, buildEpubXhtml(
                        ch.title,
                        `<div style="text-align:center;margin-top:30%;"><h1>${safeChapterTitle}</h1></div>`
                    ));
                } else {
                    const xmlContent = convertHtmlToXhtmlFragment(ch.content);
                    const epubFontFamily = fontSelect.value.split(',')[0].replace(/'/g, "");

                    oebps.file(fileName, buildEpubXhtml(
                        ch.title,
                        `<h2 style="text-align:center;margin-bottom:2em;">${safeChapterTitle}</h2>${xmlContent}`,
                        `<style>body { font-family: ${escapeXml(epubFontFamily)}, serif; }</style>`
                    ));
                }

                manifestItems += `<item id="${fileId}" href="${fileName}" media-type="application/xhtml+xml"/>\n`;
                spineItems += `<itemref idref="${fileId}"/>\n`;
                ncxNavPoints += `<navPoint id="nav-${fileId}" playOrder="${playOrder}"><navLabel><text>${safeChapterTitle}</text></navLabel><content src="${fileName}"/></navPoint>\n`;
                playOrder++;
            });

            showExportProgress('Writing EPUB metadata...', 84);
            oebps.file("content.opf", `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="BookID" version="2.0">
    <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
        <dc:title>${safeTitle}</dc:title>
        <dc:creator>${safeAuthor}</dc:creator>
        <dc:language>en</dc:language>
        <dc:identifier id="BookID">${bookId}</dc:identifier>
        ${hasCoverFile ? '<meta name="cover" content="cover-image"/>' : ''}
    </metadata>
    <manifest>
        <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
        ${manifestItems}
    </manifest>
    <spine toc="ncx">
        ${spineItems}
    </spine>
    <guide>
        ${guideItems}
    </guide>
</package>`);

            oebps.file("toc.ncx", `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE ncx PUBLIC "-//NISO//DTD ncx 2005-1//EN"
  "http://www.daisy.org/z3986/2005/ncx-2005-1.dtd">
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
    <head><meta name="dtb:uid" content="${bookId}"/><meta name="dtb:depth" content="1"/><meta name="dtb:totalPageCount" content="0"/><meta name="dtb:maxPageNumber" content="0"/></head>
    <docTitle><text>${safeTitle}</text></docTitle>
    <navMap>${ncxNavPoints}</navMap>
</ncx>`);

            showExportProgress('Compressing EPUB file...', 90);
            const content = await zip.generateAsync({
                type: "blob",
                mimeType: "application/epub+zip",
                compression: "DEFLATE",
                compressionOptions: { level: 9 }
            }, (metadata) => {
                const zipProgress = 90 + (metadata.percent * 0.1);
                showExportProgress('Compressing EPUB file...', zipProgress);
            });

            showExportProgress('Saving EPUB...', 100);
            saveAs(content, `${title}.epub`);
            setSaveStatus('', false);
            setTimeout(() => hideExportProgress(), 300);

        } catch (err) {
            console.error(err);
            setSaveStatus('', false);
            hideExportProgress();
            alert("Export Error: " + err.message);
        }
    });

    const initialSession = getSessionSnapshot();
    glAuthMode = initialSession.authMode;
    glRefreshToken = initialSession.refreshToken;
    glTokenExpiresAt = initialSession.tokenExpiresAt;

    if (initialSession.authMode === 'pat' && initialSession.token) {
        inputToken.value = initialSession.token;
        inputBranch.value = initialSession.branch;
    } else if (initialSession.authMode === 'oauth') {
        inputBranch.value = initialSession.branch;
        inputToken.placeholder = 'Signed in with GitLab';
        inputToken.disabled = true;
    }

    if (!isOAuthConfigured()) {
        oauthBtn.disabled = true;
        oauthHint.textContent = 'Add your GitLab OAuth client ID in static/js/main.js to enable one-click sign-in.';
    } else {
        oauthHint.textContent = 'Recommended: sign in with GitLab';
    }

    if (initialSession.token && !isSessionActive()) {
        clearSession();
        authErrorMsg.textContent = 'Your session expired after 30 minutes of inactivity.';
        inputToken.value = '';
        inputBranch.value = 'main';
        inputToken.disabled = false;
        glAuthMode = 'pat';
        glRefreshToken = '';
        glTokenExpiresAt = 0;
    }

    const savedTheme = localStorage.getItem('theme') || 'dark-theme';
    document.body.className = savedTheme;
    updateChapterWordCountDisplay(0);
    updateOverallWordCountDisplay(0);
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
        const callbackHandled = await handleOAuthCallback();
        if (callbackHandled) return;
        if (!isSessionActive()) return;

        const session = getSessionSnapshot();
        glAuthMode = session.authMode;
        glToken = session.token;
        glRefreshToken = session.refreshToken;
        glTokenExpiresAt = session.tokenExpiresAt;
        glProject = session.project;
        glBranch = session.branch;
        inputBranch.value = glBranch;

        if (!glProject) {
            try {
                await populateProjectSelection();
            } catch (err) {
                console.error(err);
                authErrorMsg.textContent = 'Session restore failed. Please sign in again.';
                clearSession();
                resetToWelcomeScreen();
            }
            return;
        }

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
