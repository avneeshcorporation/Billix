// All DOM references - declared as let so they can be re-assigned inside DOMContentLoaded
let productBody, addRowBtn, downloadBtn, downloadMenu, dropdownItems;
let sameAsBillTo, shipToFields, invoiceForm;
let darkModeIcon, darkModeBtn, copySelector;
let searchOldBtn, historyModal, closeModal, historyResults;
let loginOverlay, loginBtn, logoutBtn, loginUserId, loginPassword, loginError;
let formattingToolbar, toolbarBtns, spacingBtns;
let foreColorPicker, backColorPicker, currentColorIndicator, triggerForeColor, triggerBackColor;
let sidebarBtns, dashboardView, createInvoiceView, totalInvoicesCount, dashboardTableBody;
let dashboardFilters, resetFiltersBtn, sidebar, sidebarToggle, btnDownloadExcel;
let homeView, helpView, queryListContainer, queryModal, btnOpenQueryForm, closeQueryModal, btnCancelQuery, btnSubmitQuery, queryText, successToast, toastMessage;

let activeEditor = null;

// Utility to get value from either input or rich-editable div
function getFieldVal(id) {
    const el = document.getElementById(id);
    if (!el) return '';
    if (el.classList.contains('rich-editable')) {
        return el.innerText || '';
    }
    return el.value || '';
}

// Initialize
function initApp() {
    try {
        // Resolve all selectors after DOM is ready
        productBody = document.getElementById('productBody');
        addRowBtn = document.getElementById('addRow');
        downloadBtn = document.getElementById('downloadBtn');
        downloadMenu = document.getElementById('downloadMenu');
        dropdownItems = document.querySelectorAll('.dropdown-item');
        sameAsBillTo = document.getElementById('sameAsBillTo');
        shipToFields = document.getElementById('shipToFields');
        invoiceForm = document.getElementById('invoiceForm');
        darkModeIcon = document.getElementById('darkModeIcon');
        darkModeBtn = document.getElementById('darkModeBtn');
        copySelector = document.getElementById('copySelector');
        searchOldBtn = document.getElementById('searchOldBtn');
        historyModal = document.getElementById('historyModal');
        closeModal = document.getElementById('closeModal');
        historyResults = document.getElementById('historyResults');
        loginOverlay = document.getElementById('loginOverlay');
        loginBtn = document.getElementById('loginBtn');
        logoutBtn = document.getElementById('logoutBtn');
        loginUserId = document.getElementById('loginUserId');
        loginPassword = document.getElementById('loginPassword');
        loginError = document.getElementById('loginError');
        formattingToolbar = document.getElementById('formattingToolbar');
        toolbarBtns = document.querySelectorAll('.toolbar-btn:not(.dropdown-group > .toolbar-btn), .dropdown-cmd[data-cmd], .dropdown-icon-btn');
        spacingBtns = document.querySelectorAll('.spacing-btn');
        foreColorPicker = document.getElementById('foreColorPicker');
        backColorPicker = document.getElementById('backColorPicker');
        currentColorIndicator = document.getElementById('currentColorIndicator');
        triggerForeColor = document.getElementById('triggerForeColor');
        triggerBackColor = document.getElementById('triggerBackColor');

        // Sidebar & Views
        sidebarBtns = document.querySelectorAll('.sidebar-btn');
        dashboardView = document.getElementById('dashboardView');
        createInvoiceView = document.getElementById('createInvoiceView');
        totalInvoicesCount = document.getElementById('totalInvoicesCount');
        dashboardTableBody = document.getElementById('dashboardTableBody');
        dashboardFilters = document.querySelectorAll('.dashboard-filter');
        resetFiltersBtn = document.getElementById('resetFilters');
        sidebar = document.getElementById('sidebar');
        sidebarToggle = document.getElementById('sidebarToggle');
        btnDownloadExcel = document.getElementById('btnDownloadExcel');

        // Views
        homeView = document.getElementById('homeView');
        helpView = document.getElementById('helpView');
        queryListContainer = document.getElementById('queryListContainer');
        queryModal = document.getElementById('queryModal');
        btnOpenQueryForm = document.getElementById('btnOpenQueryForm');
        closeQueryModal = document.getElementById('closeQueryModal');
        btnCancelQuery = document.getElementById('btnCancelQuery');
        btnSubmitQuery = document.getElementById('btnSubmitQuery');
        queryText = document.getElementById('queryText');
        successToast = document.getElementById('successToast');
        toastMessage = document.getElementById('toastMessage');

        // Check Auth
        checkAuth();

        // Add 1 default row
        if (productBody && productBody.children.length === 0) addNewRow();

        // Set default date
        const dateInput = document.getElementById('invoiceDate');
        if (dateInput && !dateInput.value) {
            const today = new Date();
            dateInput.value = today.toISOString().split('T')[0];
            updatePreviewText('invoice-date', formatDate(today));
        }

        // Setup event listeners
        setupEventListeners();

        // Initial sync of all default values
        syncAllToPreview();

        // Check Dark Mode Preference
        if (typeof initDarkMode === 'function') initDarkMode();

        // Init Toolbar
        if (typeof initToolbar === 'function') initToolbar();

        console.log("App Initialized Successfully");
    } catch (err) {
        console.error("Initialization error:", err);
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}


function initToolbar() {
    // Basic commands
    toolbarBtns.forEach(btn => {
        btn.addEventListener('mousedown', (e) => {
            e.preventDefault();
            const cmd = btn.getAttribute('data-cmd');
            const val = btn.getAttribute('data-val') || null;
            document.execCommand(cmd, false, val);
            checkToolbarState();
            if (activeEditor) syncEditorToPreview(activeEditor);
        });
    });

    // Spacing commands (Custom Logic)
    spacingBtns.forEach(btn => {
        btn.addEventListener('mousedown', (e) => {
            e.preventDefault();
            const spacing = btn.getAttribute('data-spacing');
            applyLineSpacing(spacing);
        });
    });

    // Color Pickers
    if (triggerForeColor) triggerForeColor.addEventListener('click', () => foreColorPicker.click());
    if (triggerBackColor) triggerBackColor.addEventListener('click', () => backColorPicker.click());

    if (foreColorPicker) foreColorPicker.addEventListener('input', (e) => {
        document.execCommand('foreColor', false, e.target.value);
        if (currentColorIndicator) currentColorIndicator.style.backgroundColor = e.target.value;
        if (activeEditor) syncEditorToPreview(activeEditor);
    });

    if (backColorPicker) backColorPicker.addEventListener('input', (e) => {
        document.execCommand('hiliteColor', false, e.target.value);
        if (activeEditor) syncEditorToPreview(activeEditor);
    });

    // Mobile Support: Click to toggle dropdowns
    const dropdownGroups = document.querySelectorAll('.dropdown-group');
    dropdownGroups.forEach(group => {
        const toggleBtn = group.querySelector('.toolbar-btn');
        if (toggleBtn) {
            toggleBtn.addEventListener('mousedown', (e) => {
                e.preventDefault();
                // Close others
                dropdownGroups.forEach(g => { if (g !== group) g.classList.remove('mobile-open'); });
                group.classList.toggle('mobile-open');
                // Calculate right alignment if needed
                const dropdown = group.querySelector('.toolbar-dropdown');
                if (dropdown && dropdown.getBoundingClientRect().right > window.innerWidth) {
                    dropdown.classList.add('right-align');
                }
            });
        }
    });

    // Editors
    window.attachEditorListeners = function (editor) {
        editor.addEventListener('focus', showToolbar);
        editor.addEventListener('blur', hideToolbar);
        editor.addEventListener('mouseup', updateToolbarPosition);
        editor.addEventListener('keyup', checkToolbarState);
        editor.addEventListener('input', (e) => {
            const syncKey = e.target.getAttribute('data-sync');
            if (syncKey) syncEditorToPreview(e.target);
        });

        // Prevent enter key on single-line editors
        if (editor.classList.contains('single-line')) {
            editor.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                }
            });
        }
    };

    function initDarkMode() {
        try {
            const savedTheme = localStorage.getItem('billix-theme');
            if (savedTheme === 'dark' || (!savedTheme && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
                document.body.classList.add('dark-mode');
                if (darkModeIcon) darkModeIcon.textContent = '☀️';
            }
        } catch (e) {
            console.error("Dark mode init failed", e);
        }
    }

    const editors = document.querySelectorAll('.rich-editable');
    editors.forEach(editor => attachEditorListeners(editor));

    document.addEventListener('mousedown', (e) => {
        if (!formattingToolbar.contains(e.target) && !e.target.classList.contains('rich-editable')) {
            hideToolbar();
        }
    });
}

function attachEditorListeners(editor) {
    if (!editor) return;
    editor.addEventListener('focus', showToolbar);
    editor.addEventListener('blur', hideToolbar);
    editor.addEventListener('input', () => {
        syncEditorToPreview(editor);
        // Constraint for single-line editors
        if (editor.classList.contains('single-line')) {
            editor.innerHTML = editor.innerText.replace(/\r?\n|\r/g, "");
        }
    });
}

function showToolbar(e) {
    activeEditor = e.target;
    formattingToolbar.classList.add('visible');
    updateToolbarPosition(e);
    checkToolbarState();
}

function hideToolbar() {
    setTimeout(() => {
        if (!formattingToolbar.contains(document.activeElement) && document.activeElement !== activeEditor) {
            formattingToolbar.classList.remove('visible');
            activeEditor = null;
        }
    }, 150);
}

function updateToolbarPosition(e) {
    if (!activeEditor || !formattingToolbar.classList.contains('visible')) return;
    const rect = activeEditor.getBoundingClientRect();
    const toolbarRect = formattingToolbar.getBoundingClientRect();
    let top = rect.top + window.scrollY - toolbarRect.height - 8;
    if (top < 50) top = rect.bottom + window.scrollY + 8;
    formattingToolbar.style.top = `${top}px`;
    formattingToolbar.style.left = `${rect.left + window.scrollX}px`;
}

function checkToolbarState() {
    if (!activeEditor) return;
    toolbarBtns.forEach(btn => {
        const cmd = btn.getAttribute('data-cmd');
        const val = btn.getAttribute('data-val');
        if (!cmd) return;
        try {
            if (val && cmd === 'formatBlock') {
                const currentBlock = document.queryCommandValue(cmd);
                if (currentBlock && currentBlock.toLowerCase() === val.toLowerCase()) btn.classList.add('active');
                else btn.classList.remove('active');
            } else if (document.queryCommandState(cmd)) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        } catch (e) { }
    });

    // Evaluate spacing
    let spacingVal = '1';
    let node = getSelectionNode();
    if (node) {
        let parentBlock = findParentBlock(node);
        if (parentBlock && parentBlock.style.lineHeight) {
            spacingVal = parentBlock.style.lineHeight;
        }
    }
    spacingBtns.forEach(btn => {
        if (btn.getAttribute('data-spacing') === spacingVal) btn.classList.add('active');
        else btn.classList.remove('active');
    });
}

function getSelectionNode() {
    const selection = window.getSelection();
    if (selection.rangeCount > 0) return selection.getRangeAt(0).startContainer;
    return null;
}

function findParentBlock(node) {
    let curr = node;
    while (curr && curr !== activeEditor && curr.nodeType === 1) {
        const display = window.getComputedStyle(curr).display;
        if (display === 'block' || display === 'list-item') return curr;
        curr = curr.parentNode;
    }
    return activeEditor; // fallback
}

function applyLineSpacing(spacing) {
    if (!activeEditor) return;
    const node = getSelectionNode();
    if (!node) return;
    let block = findParentBlock(node);
    if (!block || block === activeEditor) {
        document.execCommand('formatBlock', false, 'p');
        const newNode = getSelectionNode();
        block = findParentBlock(newNode);
    }
    if (block) {
        block.style.lineHeight = spacing;
        syncEditorToPreview(activeEditor);
        checkToolbarState();
    }
}

function syncEditorToPreview(editor) {
    const syncKey = editor.getAttribute('data-sync');
    if (syncKey) {
        updatePreviewText(syncKey, editor.innerHTML, true);
        if (sameAsBillTo.checked && syncKey.startsWith('bill-to-')) {
            const shipKey = syncKey.replace('bill-to-', 'ship-to-');
            updatePreviewText(shipKey, editor.innerHTML, true);
        }
    }
}

function initDarkMode() {
    const isDark = localStorage.getItem('billix-theme') === 'dark';
    if (isDark) {
        document.body.classList.add('dark-mode');
        darkModeIcon.textContent = '☀️';
    }
}

function setupEventListeners() {
    // Helper to get value from either input or contenteditable
    const getVal = (id) => {
        const el = document.getElementById(id);
        if (!el) return '';
        return el.classList.contains('rich-editable') ? el.innerText : el.value;
    };

    // Checkbox for Ship To
    if (sameAsBillTo) {
        sameAsBillTo.addEventListener('change', (e) => {
            if (e.target.checked) {
                if (shipToFields) shipToFields.classList.add('hidden');
                // Sync Ship To with Bill To in preview
                const previewName = document.getElementById('preview-ship-to-name');
                const previewAddr = document.getElementById('preview-ship-to-address');
                const previewGST = document.getElementById('preview-ship-to-gst');
                const previewState = document.getElementById('preview-ship-to-state');
                if (previewName) previewName.textContent = getVal('buyerName');
                if (previewAddr) previewAddr.textContent = getVal('buyerAddress');
                if (previewGST) previewGST.textContent = getVal('buyerGST');
                if (previewState) previewState.textContent = getVal('buyerStateCode');
            } else {
                if (shipToFields) shipToFields.classList.remove('hidden');
            }
        });
    }

    // Dark Mode Toggle
    if (darkModeBtn) {
        darkModeBtn.addEventListener('click', () => {
            document.body.classList.toggle('dark-mode');
            const isDark = document.body.classList.contains('dark-mode');
            if (darkModeIcon) darkModeIcon.textContent = isDark ? '☀️' : '🌙';
            localStorage.setItem('billix-theme', isDark ? 'dark' : 'light');
        });
    }

    // Copy Type Selector
    if (copySelector) {
        copySelector.addEventListener('change', (e) => {
            const copyText = document.getElementById('copyText');
            if (copyText) copyText.textContent = `(${e.target.value})`;
        });
    }

    // Theme Color Selector
    const themeSelector = document.getElementById('themeSelector');
    if (themeSelector) {
        themeSelector.addEventListener('change', (e) => {
            if (e.target.value === 'blue') {
                document.body.classList.add('theme-blue');
            } else {
                document.body.classList.remove('theme-blue');
            }
        });
    }

    // Invoice Size Selector
    const invoiceSizeSelector = document.getElementById('invoiceSizeSelector');
    if (invoiceSizeSelector) {
        invoiceSizeSelector.addEventListener('change', (e) => {
            const preview = document.getElementById('invoicePreview');
            if (preview) {
                if (e.target.value === 'a5') {
                    preview.classList.add('size-a5');
                } else {
                    preview.classList.remove('size-a5');
                }
            }
        });
    }

    // Row management
    if (addRowBtn) {
        addRowBtn.addEventListener('click', () => {
            addNewRow();
            calculateTotals();
        });
    }

    // Live Sync & Calculations
    if (invoiceForm) {
        invoiceForm.addEventListener('input', (e) => {
            const syncKey = e.target.getAttribute('data-sync');

            if (syncKey) {
                let isRich = e.target.classList.contains('rich-editable');
                let val = isRich ? e.target.innerHTML : e.target.value;
                if (e.target.type === 'date') val = formatDate(new Date(val));
                updatePreviewText(syncKey, val, isRich);

                // If Ship To is synced with Bill To
                if (sameAsBillTo && sameAsBillTo.checked && syncKey.startsWith('bill-to-')) {
                    const shipKey = syncKey.replace('bill-to-', 'ship-to-');
                    updatePreviewText(shipKey, val, isRich);
                }
            }

            // Trigger calculations if numerical inputs change
            if (e.target.type === 'number' || e.target.classList.contains('calc-trigger')) {
                const row = e.target.closest('tr');
                if (row) calculateRow(row);
                calculateTotals();
            }
        });
    }

    // Dropdown Toggle
    if (downloadBtn) {
        downloadBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            downloadBtn.parentElement.classList.toggle('open');
        });

        // Dropdown item selection
        dropdownItems.forEach(item => {
            item.addEventListener('click', () => {
                const copyType = item.getAttribute('data-copy');
                generatePDF(copyType);
                downloadBtn.parentElement.classList.remove('open');
            });
        });

        // Close dropdown clicking outside
        window.addEventListener('click', () => {
            downloadBtn.parentElement.classList.remove('open');
        });
    }

    // State Selector Logic
    const stateSelectors = document.querySelectorAll('.state-selector');
    stateSelectors.forEach(select => {
        select.addEventListener('change', (e) => {
            const targetId = select.getAttribute('data-target');
            const targetInput = document.getElementById(targetId);
            if (targetInput) {
                if (targetInput.classList.contains('rich-editable')) {
                    targetInput.innerHTML = select.value;
                } else {
                    targetInput.value = select.value;
                }
                targetInput.dispatchEvent(new Event('input', { bubbles: true }));
            }
        });
    });

    // History Search
    if (searchOldBtn) {
        searchOldBtn.addEventListener('click', () => {
            const day = document.getElementById('searchDay').value;
            const month = document.getElementById('searchMonth').value;
            const year = document.getElementById('searchYear').value;
            searchHistory(day, month, year);
        });
    }

    if (closeModal) {
        closeModal.addEventListener('click', () => {
            if (historyModal) historyModal.classList.add('hidden');
        });
    }

    window.addEventListener('click', (e) => {
        if (historyModal && e.target === historyModal) historyModal.classList.add('hidden');
    });

    // Login Logic
    if (loginBtn) {
        loginBtn.addEventListener('click', handleLogin);
    }
    if (loginPassword) {
        loginPassword.addEventListener('keypress', (e) => { if (e.key === 'Enter') handleLogin(); });
    }

    // Logout Logic
    if (logoutBtn) {
        logoutBtn.addEventListener('click', () => {
            sessionStorage.removeItem('billix-auth');
            window.location.reload();
        });
    }

    // Sidebar Navigation
    sidebarBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const view = btn.getAttribute('data-view');
            switchView(view);
        });
    });

    // Dashboard Filters
    dashboardFilters.forEach(filter => {
        filter.addEventListener('change', () => {
            renderDashboardRecords();
        });
    });

    if (resetFiltersBtn) {
        resetFiltersBtn.addEventListener('click', () => {
            dashboardFilters.forEach(f => f.value = '');
            renderDashboardRecords();
        });
    }

    if (btnDownloadExcel) {
        btnDownloadExcel.addEventListener('click', downloadExcel);
    }

    if (sidebarToggle) {
        sidebarToggle.addEventListener('click', toggleSidebar);
    }

    // Help View Events
    if (btnOpenQueryForm) {
        btnOpenQueryForm.addEventListener('click', () => {
            if (queryModal) queryModal.classList.remove('hidden');
        });
    }

    if (closeQueryModal) {
        closeQueryModal.addEventListener('click', () => {
            if (queryModal) queryModal.classList.add('hidden');
        });
    }

    if (btnCancelQuery) {
        btnCancelQuery.addEventListener('click', () => {
            if (queryModal) queryModal.classList.add('hidden');
        });
    }

    if (btnSubmitQuery) {
        btnSubmitQuery.addEventListener('click', submitQuery);
    }
}

function toggleSidebar() {
    if (!sidebar) return;
    sidebar.classList.toggle('collapsed');

    // Update icon ☰ → ✖
    if (sidebarToggle) {
        sidebarToggle.textContent = sidebar.classList.contains('collapsed') ? '☰' : '✖';
        sidebarToggle.title = sidebar.classList.contains('collapsed') ? 'Expand Sidebar' : 'Collapse Sidebar';
    }
}

function switchView(view) {
    sidebarBtns.forEach(b => b.classList.toggle('active', b.getAttribute('data-view') === view));

    if (view === 'dashboard') {
        dashboardView.classList.remove('hidden');
        createInvoiceView.classList.add('hidden');
        if (homeView) homeView.classList.add('hidden');
        if (helpView) helpView.classList.add('hidden');
        initDashboard();
    } else if (view === 'home') {
        dashboardView.classList.add('hidden');
        createInvoiceView.classList.add('hidden');
        if (homeView) homeView.classList.remove('hidden');
        if (helpView) helpView.classList.add('hidden');
    } else if (view === 'help') {
        dashboardView.classList.add('hidden');
        createInvoiceView.classList.add('hidden');
        if (homeView) homeView.classList.add('hidden');
        if (helpView) {
            helpView.classList.remove('hidden');
            initHelp();
        }
    } else {
        dashboardView.classList.add('hidden');
        createInvoiceView.classList.remove('hidden');
        if (homeView) homeView.classList.add('hidden');
        if (helpView) helpView.classList.add('hidden');
    }
}

function initHelp() {
    renderHelpQueries();
}

function submitQuery() {
    const text = queryText.value.trim();
    if (!text) {
        showToast("Please enter a query first", "❌");
        return;
    }

    const queries = JSON.parse(localStorage.getItem('billix-queries') || '[]');
    const newQuery = {
        id: Date.now(),
        text: text,
        status: 'Pending',
        date: new Date().toISOString(),
        replies: []
    };

    queries.unshift(newQuery);
    localStorage.setItem('billix-queries', JSON.stringify(queries));

    queryText.value = '';
    queryModal.classList.add('hidden');

    showToast("Your query has been posted successfully", "✅");
    renderHelpQueries();
}

function renderHelpQueries() {
    if (!queryListContainer) return;

    const queries = JSON.parse(localStorage.getItem('billix-queries') || '[]');
    queryListContainer.innerHTML = '';

    if (queries.length === 0) {
        queryListContainer.innerHTML = '<div class="empty-state"><p>No queries posted yet.</p></div>';
        return;
    }

    queries.forEach(q => {
        const div = document.createElement('div');
        div.className = 'query-item';

        let repliesHtml = '';
        if (q.replies && q.replies.length > 0) {
            repliesHtml = `
                <div class="query-replies">
                    ${q.replies.map(r => `
                        <div class="reply-item reply-${r.sender.toLowerCase()}">
                            <div class="reply-header">
                                <span class="sender-name">${r.sender}</span>
                                <span class="reply-date">${formatDate(new Date(r.date))}</span>
                            </div>
                            <div class="reply-content">${r.text}</div>
                        </div>
                    `).join('')}
                </div>
            `;
        }

        const isCompleted = q.status === 'Completed';

        div.innerHTML = `
            <div class="query-header">
                <span class="query-date">${formatDate(new Date(q.date))}</span>
                <span class="query-status status-${q.status.toLowerCase()}">${q.status}</span>
            </div>
            <div class="query-content">${q.text}</div>
            
            ${repliesHtml}

            <div class="query-actions">
                <button class="btn btn-secondary btn-inline" onclick="toggleReplyInput(${q.id})">
                    💬 Reply
                </button>
                <button class="btn btn-secondary btn-inline" onclick="markQueryCompleted(${q.id})" ${isCompleted ? 'disabled' : ''}>
                    ✅ ${isCompleted ? 'Completed' : 'Mark as Completed'}
                </button>
            </div>

            <div id="reply-container-${q.id}" class="reply-input-container hidden">
                <textarea id="reply-text-${q.id}" class="reply-textarea" rows="3" placeholder="Type your reply..."></textarea>
                <div style="text-align: right;">
                    <button class="btn btn-secondary btn-sm" onclick="toggleReplyInput(${q.id})">Cancel</button>
                    <button class="btn btn-primary btn-sm" onclick="submitReply(${q.id})">Send Reply</button>
                </div>
            </div>
        `;
        queryListContainer.appendChild(div);
    });
}

window.toggleReplyInput = (id) => {
    const container = document.getElementById(`reply-container-${id}`);
    if (container) container.classList.toggle('hidden');
};

window.markQueryCompleted = (id) => {
    const queries = JSON.parse(localStorage.getItem('billix-queries') || '[]');
    const q = queries.find(item => item.id === id);
    if (q) {
        q.status = 'Completed';
        localStorage.setItem('billix-queries', JSON.stringify(queries));
        showToast("Query marked as completed", "✅");
        renderHelpQueries();
    }
};

window.submitReply = (id) => {
    const textEl = document.getElementById(`reply-text-${id}`);
    const text = textEl.value.trim();
    if (!text) return;

    const queries = JSON.parse(localStorage.getItem('billix-queries') || '[]');
    const q = queries.find(item => item.id === id);
    if (q) {
        if (!q.replies) q.replies = [];
        q.replies.push({
            id: Date.now(),
            text: text,
            sender: 'User',
            date: new Date().toISOString()
        });
        localStorage.setItem('billix-queries', JSON.stringify(queries));
        textEl.value = '';
        toggleReplyInput(id);
        showToast("Reply sent successfully", "✅");
        renderHelpQueries();
    }
};

function showToast(message, icon = '✅') {
    if (!successToast || !toastMessage) return;

    toastMessage.textContent = message;
    const iconSpan = successToast.querySelector('.toast-icon');
    if (iconSpan) iconSpan.textContent = icon;

    successToast.classList.remove('hidden');

    setTimeout(() => {
        successToast.classList.add('hidden');
    }, 4000);
}

function initDashboard() {
    const history = JSON.parse(localStorage.getItem('billix-history') || '[]');
    if (totalInvoicesCount) totalInvoicesCount.textContent = history.length;

    populateDashboardFilters(history);
    renderDashboardRecords();
}

function populateDashboardFilters(history) {
    const filters = {
        filterBillToName: new Set(),
        filterShipToName: new Set(),
        filterBillToGST: new Set(),
        filterShipToGST: new Set(),
        filterBillToState: new Set(),
        filterShipToState: new Set(),
        filterYear: new Set(),
        filterMonth: new Set()
    };

    history.forEach(inv => {
        if (inv.buyer) filters.filterBillToName.add(inv.buyer);

        // Ship To Name might be in form data
        const shipName = inv.form['shipToName'] || inv.buyer;
        if (shipName) filters.filterShipToName.add(shipName);

        if (inv.form['buyerGST']) filters.filterBillToGST.add(inv.form['buyerGST']);
        if (inv.form['shipToGST']) filters.filterShipToGST.add(inv.form['shipToGST']);

        if (inv.form['buyerStateCode']) filters.filterBillToState.add(inv.form['buyerStateCode']);
        if (inv.form['shipToStateCode']) filters.filterShipToState.add(inv.form['shipToStateCode']);

        if (inv.date) {
            const d = new Date(inv.date);
            if (!isNaN(d.getTime())) {
                filters.filterYear.add(d.getFullYear().toString());
                filters.filterMonth.add((d.getMonth() + 1).toString().padStart(2, '0'));
            }
        }
    });

    // Populate dropdowns
    for (const [id, values] of Object.entries(filters)) {
        const select = document.getElementById(id);
        if (!select) continue;

        const currentValue = select.value;
        const firstOption = select.options[0];
        select.innerHTML = '';
        select.appendChild(firstOption);

        Array.from(values).sort().forEach(val => {
            const opt = document.createElement('option');
            opt.value = val;
            opt.textContent = val;
            if (id === 'filterMonth') {
                const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                opt.textContent = months[parseInt(val) - 1];
            }
            select.appendChild(opt);
        });
        select.value = currentValue;
    }
}

function renderDashboardRecords() {
    if (!dashboardTableBody) return;

    const history = JSON.parse(localStorage.getItem('billix-history') || '[]');
    const filters = {};
    dashboardFilters.forEach(f => {
        if (f.value) filters[f.id] = f.value;
    });

    const filtered = history.filter(inv => {
        const d = new Date(inv.date);
        const invYear = d.getFullYear().toString();
        const invMonth = (d.getMonth() + 1).toString().padStart(2, '0');

        if (filters.filterBillToName && inv.buyer !== filters.filterBillToName) return false;
        if (filters.filterShipToName && (inv.form['shipToName'] || inv.buyer) !== filters.filterShipToName) return false;
        if (filters.filterBillToGST && inv.form['buyerGST'] !== filters.filterBillToGST) return false;
        if (filters.filterShipToGST && inv.form['shipToGST'] !== filters.filterShipToGST) return false;
        if (filters.filterBillToState && inv.form['buyerStateCode'] !== filters.filterBillToState) return false;
        if (filters.filterShipToState && inv.form['shipToStateCode'] !== filters.filterShipToState) return false;
        if (filters.filterYear && invYear !== filters.filterYear) return false;
        if (filters.filterMonth && invMonth !== filters.filterMonth) return false;

        return true;
    });

    dashboardTableBody.innerHTML = '';

    if (filtered.length === 0) {
        dashboardTableBody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding: 2rem; color: var(--text-muted);">No records found matching filters.</td></tr>';
        return;
    }

    filtered.forEach(inv => {
        // Calculate total amount for display
        let total = 0;
        if (inv.products) {
            inv.products.forEach(p => total += (parseFloat(p.qty) || 0) * (parseFloat(p.rate) || 0));
        }
        // Add taxes from form if available
        const cgst = parseFloat(inv.form['cgst']) || 0;
        const sgst = parseFloat(inv.form['sgst']) || 0;
        const igst = parseFloat(inv.form['igst']) || 0;
        total = total + (total * (cgst + sgst + igst) / 100);

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${formatDate(new Date(inv.date))}</td>
            <td><strong>${inv.invNo}</strong></td>
            <td>${inv.buyer}</td>
            <td><span class="badge-amount">₹ ${formatNumber(total)}</span></td>
        `;
        dashboardTableBody.appendChild(tr);
    });
}

function downloadExcel() {
    const history = JSON.parse(localStorage.getItem('billix-history') || '[]');
    const filters = {};
    dashboardFilters.forEach(f => {
        if (f.value) filters[f.id] = f.value;
    });

    const filtered = history.filter(inv => {
        const d = new Date(inv.date);
        const invYear = d.getFullYear().toString();
        const invMonth = (d.getMonth() + 1).toString().padStart(2, '0');

        if (filters.filterBillToName && inv.buyer !== filters.filterBillToName) return false;
        if (filters.filterShipToName && (inv.form['shipToName'] || inv.buyer) !== filters.filterShipToName) return false;
        if (filters.filterBillToGST && inv.form['buyerGST'] !== filters.filterBillToGST) return false;
        if (filters.filterShipToGST && inv.form['shipToGST'] !== filters.filterShipToGST) return false;
        if (filters.filterBillToState && inv.form['buyerStateCode'] !== filters.filterBillToState) return false;
        if (filters.filterShipToState && inv.form['shipToStateCode'] !== filters.filterShipToState) return false;
        if (filters.filterYear && invYear !== filters.filterYear) return false;
        if (filters.filterMonth && invMonth !== filters.filterMonth) return false;

        return true;
    });

    if (filtered.length === 0) {
        showToast("No data available for selected filters", "⚠️");
        return;
    }

    showToast("Generating Excel file...", "🔄");

    try {
        const excelData = filtered.map(inv => {
            // Calculate total amount
            let subTotal = 0;
            let productDetails = '';
            if (inv.products) {
                inv.products.forEach(p => {
                    const lineTotal = (parseFloat(p.qty) || 0) * (parseFloat(p.rate) || 0);
                    subTotal += lineTotal;
                    productDetails += `${p.desc} (${p.qty} ${p.qtyUnit} @ ${p.rate}), `;
                });
            }
            const cgst = parseFloat(inv.form['cgst']) || 0;
            const sgst = parseFloat(inv.form['sgst']) || 0;
            const igst = parseFloat(inv.form['igst']) || 0;
            const total = subTotal + (subTotal * (cgst + sgst + igst) / 100);

            return {
                'Invoice Number': inv.invNo,
                'Date': formatDate(new Date(inv.date)),
                'Bill To Name': inv.buyer,
                'Ship To Name': inv.form['shipToName'] || inv.buyer,
                'Bill To GSTIN': inv.form['buyerGST'] || '',
                'Ship To GSTIN': inv.form['shipToGST'] || '',
                'Bill To State': inv.form['buyerStateCode'] || '',
                'Ship To State': inv.form['shipToStateCode'] || '',
                'Bill To Address': inv.form['buyerAddress'] || '',
                'Ship To Address': inv.form['shipToAddress'] || '',
                'Transport': inv.form['transportName'] || '',
                'Vehicle No': inv.form['lorryNo'] || '',
                'Subtotal': subTotal,
                'CGST %': cgst,
                'SGST %': sgst,
                'IGST %': igst,
                'Total Amount': total,
                'Items': productDetails.slice(0, -2) // Remove trailing comma
            };
        });

        const worksheet = XLSX.utils.json_to_sheet(excelData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Invoices");

        // Format column widths
        const wscols = [
            { wch: 15 }, { wch: 12 }, { wch: 25 }, { wch: 25 }, { wch: 18 }, { wch: 18 },
            { wch: 15 }, { wch: 15 }, { wch: 35 }, { wch: 35 }, { wch: 15 }, { wch: 15 },
            { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 15 }, { wch: 50 }
        ];
        worksheet['!cols'] = wscols;

        const fileName = `Invoices_Report_${new Date().toISOString().split('T')[0]}.xlsx`;
        XLSX.writeFile(workbook, fileName);

        showToast("Excel downloaded successfully", "✅");
    } catch (error) {
        console.error("Excel generation failed:", error);
        showToast("Failed to generate Excel", "❌");
    }
}

function checkAuth() {
    const isAuth = sessionStorage.getItem('billix-auth') === 'true';
    if (isAuth) {
        loginOverlay.classList.add('hidden');
        switchView('dashboard');
    } else {
        loginOverlay.classList.remove('hidden');
    }
}

function handleLogin() {
    window.actualHandleLogin();
}

window.actualHandleLogin = function () {
    const userEl = loginUserId || document.getElementById('loginUserId');
    const passEl = loginPassword || document.getElementById('loginPassword');
    const errorEl = loginError || document.getElementById('loginError');
    const overlayEl = loginOverlay || document.getElementById('loginOverlay');

    if (!userEl || !passEl) return;

    const userId = userEl.value.trim();
    const password = passEl.value.trim();

    // Hardcoded Credentials - Case insensitive for User ID
    if (userId.toLowerCase() === 'avneesh.co' && password === 'Avneesh@2026') {
        try {
            sessionStorage.setItem('billix-auth', 'true');
        } catch (e) {
            console.error("Session storage blocked:", e);
        }

        if (overlayEl) {
            overlayEl.style.opacity = '0';
            setTimeout(() => {
                overlayEl.style.display = 'none';
                overlayEl.classList.add('hidden');
                // Redirect to Dashboard
                switchView('dashboard');
            }, 500);
        }
    } else {
        if (errorEl) {
            errorEl.classList.remove('hidden');
            setTimeout(() => {
                errorEl.classList.add('hidden');
            }, 3000);
        }
    }
}

function updatePreviewText(key, value, isRich = false) {
    const el = document.getElementById(`preview-${key}`) || document.getElementById(`p-${key}`);
    if (el) {
        if (isRich) {
            el.innerHTML = value || '';
        } else {
            el.textContent = value || '';
        }
    }
}

function syncAllToPreview() {
    const inputs = invoiceForm.querySelectorAll('[data-sync]');
    inputs.forEach(input => {
        let isRich = input.classList.contains('rich-editable');
        let val = isRich ? input.innerHTML : input.value;
        if (input.type === 'date') val = formatDate(new Date(val));
        updatePreviewText(input.getAttribute('data-sync'), val, isRich);
    });
}

function formatDate(date) {
    if (!date || isNaN(date.getTime())) return '';
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function addNewRow() {
    const rowCount = productBody.children.length + 1;
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td class="sn-cell">${rowCount}.</td>
        <td><div class="rich-editable single-line product-desc calc-trigger" contenteditable="true" data-placeholder="e.g. MS SKULL"></div></td>
        <td><div class="rich-editable single-line product-hsn calc-trigger center-text" contenteditable="true" data-placeholder="HSN/SAC"></div></td>
        <td>
            <input type="number" class="product-qty calc-trigger right-text" value="0" step="0.001">
            <select class="product-unit calc-trigger" style="font-size: 0.65rem; padding: 2px; margin-top: 2px; width: 100%;">
                <option value="KGS" selected>Kg</option>
                <option value="NOS">Nos</option>
                <option value="PCS">Pcs</option>
                <option value="GRAM">Gram</option>
                <option value="TON">Ton</option>
                <option value="META TON">Meta ton</option>
                <option value="LITRE">Litre</option>
                <option value="ML">Ml</option>
                <option value="METER">Meter</option>
                <option value="FEET">Feet</option>
                <option value="BOX">Box</option>
                <option value="PACK">Pack</option>
                <option value="DOZEN">Dozen</option>
                <option value="HOURS">Hours</option>
                <option value="DAYS">Days</option>
                <option value="UNITS">Units</option>
                <option value="OTHER">Other</option>
            </select>
        </td>
        <td>
            <input type="number" class="product-rate calc-trigger right-text" value="0" step="0.01">
            <select class="product-rate-unit calc-trigger" style="font-size: 0.65rem; padding: 2px; margin-top: 2px; width: 100%;">
                <option value="Per KGS" selected>Per Kg</option>
                <option value="Per NOS">Per Nos</option>
                <option value="Per PCS">Per Pcs</option>
                <option value="Per GRAM">Per Gram</option>
                <option value="Per TON">Per Ton</option>
                <option value="Per META TON">Per Meta ton</option>
                <option value="Per LITRE">Per Litre</option>
                <option value="Per ML">Per Ml</option>
                <option value="Per METER">Per Meter</option>
                <option value="Per FEET">Per Feet</option>
                <option value="Per BOX">Per Box</option>
                <option value="Per PACK">Per Pack</option>
                <option value="Per DOZEN">Per Dozen</option>
                <option value="Per HOURS">Per Hour</option>
                <option value="Per DAYS">Per Day</option>
                <option value="Per UNITS">Per Unit</option>
                <option value="OTHER">Other</option>
            </select>
        </td>
        <td><button type="button" class="btn btn-danger btn-sm remove-row" style="padding: 0.2rem 0.5rem; font-size: 0.9rem;">×</button></td>
    `;

    const qtyUnit = tr.querySelector('.product-unit');
    const rateUnit = tr.querySelector('.product-rate-unit');

    qtyUnit.addEventListener('change', () => {
        rateUnit.value = "Per " + qtyUnit.value;
        updatePreviewTable();
    });

    tr.querySelector('.remove-row').addEventListener('click', () => {
        if (productBody.children.length > 1) {
            tr.remove();
            updateRowNumbers();
            calculateTotals();
        }
    });

    productBody.appendChild(tr);

    // Attach listeners to new rich-editable fields
    tr.querySelectorAll('.rich-editable').forEach(editor => attachEditorListeners(editor));

    updatePreviewTable();
}

function updateRowNumbers() {
    Array.from(productBody.children).forEach((row, index) => {
        row.cells[0].textContent = (index + 1) + '.';
    });
}

function calculateRow(row) {
    // This function is kept for legacy but calculations are now centralized in updatePreviewTable
    updatePreviewTable();
}

function updatePreviewTable() {
    const previewBody = document.getElementById('preview-product-body');
    previewBody.innerHTML = '';

    let subTotal = 0;
    const rows = Array.from(productBody.children);

    rows.forEach((row, index) => {
        const descEl = row.querySelector('.product-desc');
        const hsnEl = row.querySelector('.product-hsn');
        const desc = descEl.classList.contains('rich-editable') ? descEl.innerHTML : descEl.value;
        const hsn = hsnEl.classList.contains('rich-editable') ? hsnEl.innerHTML : hsnEl.value;
        const qty = parseFloat(row.querySelector('.product-qty').value) || 0;
        const rate = parseFloat(row.querySelector('.product-rate').value) || 0;
        const qtyUnit = row.querySelector('.product-unit').value;
        const rateUnit = row.querySelector('.product-rate-unit').value;
        const amount = qty * rate;
        subTotal += amount;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="center">${index + 1}.</td>
            <td>${desc || '---'}</td>
            <td class="center">${hsn || '---'}</td>
            <td class="center">${formatNumber(qty)}<br>${qtyUnit}</td>
            <td class="center">${formatNumber(rate)} /-<br>${rateUnit}</td>
            <td class="right">${formatNumber(amount)}</td>
        `;
        previewBody.appendChild(tr);
    });

    // Add spacer row if items are few
    if (rows.length < 5) {
        const spacer = document.createElement('tr');
        spacer.className = 'spacer-row';
        spacer.innerHTML = `<td colspan="6"></td>`;
        previewBody.appendChild(spacer);
    }

    calculateTotals(subTotal);
}

function calculateTotals(subTotalValue) {
    let subTotal = subTotalValue;
    if (subTotal === undefined) {
        subTotal = 0;
        Array.from(productBody.children).forEach(row => {
            const qty = parseFloat(row.querySelector('.product-qty').value) || 0;
            const rate = parseFloat(row.querySelector('.product-rate').value) || 0;
            subTotal += (qty * rate);
        });
    }

    const cgstRate = parseFloat(document.getElementById('cgst').value) || 0;
    const sgstRate = parseFloat(document.getElementById('sgst').value) || 0;
    const igstRate = parseFloat(document.getElementById('igst').value) || 0;

    const cgstVal = (subTotal * cgstRate) / 100;
    const sgstVal = (subTotal * sgstRate) / 100;
    const igstVal = (subTotal * igstRate) / 100;
    const grossTotal = subTotal + cgstVal + sgstVal + igstVal;

    // Update Preview
    document.getElementById('preview-subtotal').textContent = formatNumber(subTotal);

    document.getElementById('preview-cgst-rate').textContent = cgstRate;
    document.getElementById('preview-cgst-val').textContent = cgstVal > 0 ? formatNumber(cgstVal) : '-';

    document.getElementById('preview-sgst-rate').textContent = sgstRate;
    document.getElementById('preview-sgst-val').textContent = sgstVal > 0 ? formatNumber(sgstVal) : '-';

    document.getElementById('preview-igst-rate').textContent = igstRate;
    document.getElementById('preview-igst-val').textContent = igstVal > 0 ? formatNumber(igstVal) : '-';

    document.getElementById('preview-gross-total').textContent = formatNumber(grossTotal);

    const words = numberToWords(Math.round(grossTotal * 100) / 100);
    document.getElementById('preview-amount-words').textContent = words ? (words + " ONLY") : "ZERO RUPEES ONLY";
}

function formatNumber(num) {
    return num.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Indian Number System Converter
function numberToWords(num) {
    if (num === 0) return '';
    const a = ['', 'ONE ', 'TWO ', 'THREE ', 'FOUR ', 'FIVE ', 'SIX ', 'SEVEN ', 'EIGHT ', 'NINE ', 'TEN ', 'ELEVEN ', 'TWELVE ', 'THIRTEEN ', 'FOURTEEN ', 'FIFTEEN ', 'SIXTEEN ', 'SEVENTEEN ', 'EIGHTEEN ', 'NINETEEN '];
    const b = ['', '', 'TWENTY ', 'THIRTY ', 'FORTY ', 'FIFTY ', 'SIXTY ', 'SEVENTY ', 'EIGHTY ', 'NINETY '];

    const convert_less_than_thousand = (n) => {
        if (n === 0) return '';
        if (n < 20) return a[n];
        const res = b[Math.floor(n / 10)] + a[n % 10];
        return res;
    };

    const convert_with_suffix = (n, suffix) => {
        if (n === 0) return '';
        if (n < 100) return convert_less_than_thousand(n) + suffix;
        return a[Math.floor(n / 100)] + 'HUNDRED ' + convert_less_than_thousand(n % 100) + suffix;
    };

    const integerPart = Math.floor(num);
    const decimalPart = Math.round((num - integerPart) * 100);

    let str = '';
    str += convert_with_suffix(Math.floor(integerPart / 10000000), 'CRORE ');
    str += convert_with_suffix(Math.floor((integerPart / 100000) % 100), 'LAKH ');
    str += convert_with_suffix(Math.floor((integerPart / 1000) % 100), 'THOUSAND ');
    str += convert_with_suffix(integerPart % 1000, '');

    let result = str.trim();

    if (decimalPart > 0) {
        const paiseStr = convert_with_suffix(decimalPart, '').trim();
        if (result !== '') {
            result += ' AND ' + paiseStr + ' PAISE';
        } else {
            result = paiseStr + ' PAISE';
        }
    }

    return result;
}

// PDF Generation Logic
async function generatePDF(copyType) {
    const { jsPDF } = window.jspdf;
    const pageSize = document.getElementById('invoiceSizeSelector').value || 'a4';
    const doc = new jsPDF('p', 'mm', pageSize);

    const allCopies = [
        "ORIGINAL FOR RECIPIENT",
        "DUPLICATE FOR TRANSPORTER",
        "SUPPLIER COPY"
    ];

    const copiesToRender = copyType === 'ALL' ? allCopies : [copyType];

    copiesToRender.forEach((title, index) => {
        if (index > 0) doc.addPage(pageSize, 'p');
        renderPDFPage(doc, title);
    });

    const invNo = getFieldVal('invoiceNo') || 'Draft';
    const fileName = copyType === 'ALL' ?
        `Invoice_${invNo}_All_Copies.pdf` :
        `Invoice_${invNo}_${copyType.split(' ')[0]}.pdf`;

    doc.save(fileName);

    // Save state as latest
    saveCurrentStateAsLatest(copyType);

    // Save to history
    saveToHistory(copyType);
}

function saveToHistory(copyType) {
    const history = JSON.parse(localStorage.getItem('billix-history') || '[]');
    const invNo = getFieldVal('invoiceNo') || 'Draft';
    const date = document.getElementById('invoiceDate').value;
    const buyer = getFieldVal('buyerName');

    const state = {
        id: Date.now(),
        invNo,
        date,
        buyer,
        copyType,
        form: {},
        products: Array.from(productBody.children).map(row => {
            const descEl = row.querySelector('.product-desc');
            const hsnEl = row.querySelector('.product-hsn');
            return {
                desc: descEl.classList.contains('rich-editable') ? descEl.innerHTML : descEl.value,
                hsn: hsnEl.classList.contains('rich-editable') ? hsnEl.innerHTML : hsnEl.value,
                qty: row.querySelector('.product-qty').value,
                rate: row.querySelector('.product-rate').value,
                qtyUnit: row.querySelector('.product-unit').value,
                rateUnit: row.querySelector('.product-rate-unit').value
            };
        })
    };

    const inputs = invoiceForm.querySelectorAll('input, select, .rich-editable');
    inputs.forEach(el => {
        if (el.id) {
            state.form[el.id] = el.classList.contains('rich-editable') ? el.innerHTML : el.value;
        }
    });

    history.unshift(state);
    // Keep only last 100 invoices to save space
    if (history.length > 100) history.pop();
    localStorage.setItem('billix-history', JSON.stringify(history));

    // Update dashboard if it's the current view
    if (typeof initDashboard === 'function') initDashboard();
}

function searchHistory(day, month, year) {
    const history = JSON.parse(localStorage.getItem('billix-history') || '[]');
    const filtered = history.filter(inv => {
        const d = new Date(inv.date);
        const invDay = String(d.getDate()).padStart(2, '0');
        const invMonth = String(d.getMonth() + 1).padStart(2, '0');
        const invYear = String(d.getFullYear());

        return (!day || day === invDay) &&
            (!month || month === invMonth) &&
            (!year || year === invYear);
    });

    displayHistoryResults(filtered);
}

function displayHistoryResults(results) {
    historyResults.innerHTML = '';
    if (results.length === 0) {
        historyResults.innerHTML = '<p style="text-align: center; padding: 20px;">No invoices found for the selected criteria.</p>';
    } else {
        results.forEach(inv => {
            const div = document.createElement('div');
            div.className = 'history-item';
            div.innerHTML = `
                <div class="history-info">
                    <p><strong>Inv: ${inv.invNo}</strong> | ${formatDate(new Date(inv.date))}</p>
                    <p style="font-size: 0.8rem; color: var(--text-muted);">${inv.buyer}</p>
                </div>
                <div class="history-actions">
                    <button class="btn-icon" onclick="downloadOldInvoiceById(${inv.id})">📥 PDF</button>
                </div>
            `;
            historyResults.appendChild(div);
        });
    }
    historyModal.classList.remove('hidden');
}

window.downloadOldInvoiceById = (id) => {
    const history = JSON.parse(localStorage.getItem('billix-history') || '[]');
    const inv = history.find(i => i.id === id);
    if (inv) {
        downloadLatestSavedInvoice(inv);
        historyModal.classList.add('hidden');
    }
};

function saveCurrentStateAsLatest(copyType) {
    const state = {
        copyType: copyType,
        form: {}
    };

    // Capture all inputs with data-sync
    const inputs = invoiceForm.querySelectorAll('input, select, .rich-editable');
    inputs.forEach(el => {
        if (el.id) {
            state.form[el.id] = el.classList.contains('rich-editable') ? el.innerHTML : el.value;
        }
    });

    // Capture products
    state.products = Array.from(productBody.children).map(row => {
        const descEl = row.querySelector('.product-desc');
        const hsnEl = row.querySelector('.product-hsn');
        return {
            desc: descEl.classList.contains('rich-editable') ? descEl.innerHTML : descEl.value,
            hsn: hsnEl.classList.contains('rich-editable') ? hsnEl.innerHTML : hsnEl.value,
            qty: row.querySelector('.product-qty').value,
            rate: row.querySelector('.product-rate').value,
            qtyUnit: row.querySelector('.product-unit').value,
            rateUnit: row.querySelector('.product-rate-unit').value
        };
    });

    localStorage.setItem('billix-latest-invoice', JSON.stringify(state));
}

async function downloadLatestSavedInvoice(state) {
    // To generate the PDF, we need the DOM to have the data.
    // We'll backup current state, load saved state, generate, then restore.
    const backup = {
        products: Array.from(productBody.children).map(row => {
            const descEl = row.querySelector('.product-desc');
            const hsnEl = row.querySelector('.product-hsn');
            return {
                desc: descEl.classList.contains('rich-editable') ? descEl.innerHTML : descEl.value,
                hsn: hsnEl.classList.contains('rich-editable') ? hsnEl.innerHTML : hsnEl.value,
                qty: row.querySelector('.product-qty').value,
                rate: row.querySelector('.product-rate').value,
                qtyUnit: row.querySelector('.product-unit').value,
                rateUnit: row.querySelector('.product-rate-unit').value
            };
        }),
        form: {}
    };

    const inputs = invoiceForm.querySelectorAll('input, select, .rich-editable');
    inputs.forEach(el => {
        if (el.id) {
            backup.form[el.id] = el.classList.contains('rich-editable') ? el.innerHTML : el.value;
        }
    });

    // Load saved state
    for (const [id, val] of Object.entries(state.form)) {
        const el = document.getElementById(id);
        if (el) {
            if (el.classList.contains('rich-editable')) el.innerHTML = val;
            else el.value = val;
        }
    }

    // Load products
    productBody.innerHTML = '';
    state.products.forEach(p => {
        addNewRow();
        const row = productBody.lastElementChild;
        row.querySelector('.product-desc').value = p.desc;
        row.querySelector('.product-hsn').value = p.hsn;
        row.querySelector('.product-qty').value = p.qty;
        row.querySelector('.product-rate').value = p.rate;
        row.querySelector('.product-unit').value = p.qtyUnit;
        row.querySelector('.product-rate-unit').value = p.rateUnit;
    });

    // Trigger sync and generation
    syncAllToPreview();
    calculateTotals();

    await generatePDF(state.copyType);

    // Restore backup
    for (const [id, val] of Object.entries(backup.form)) {
        const el = document.getElementById(id);
        if (el) {
            if (el.classList.contains('rich-editable')) el.innerHTML = val;
            else el.value = val;
        }
    }

    productBody.innerHTML = '';
    backup.products.forEach(p => {
        addNewRow();
        const row = productBody.lastElementChild;
        const descEl = row.querySelector('.product-desc');
        const hsnEl = row.querySelector('.product-hsn');

        if (descEl.classList.contains('rich-editable')) descEl.innerHTML = p.desc;
        else descEl.value = p.desc;

        if (hsnEl.classList.contains('rich-editable')) hsnEl.innerHTML = p.hsn;
        else hsnEl.value = p.hsn;

        row.querySelector('.product-qty').value = p.qty;
        row.querySelector('.product-rate').value = p.rate;
        row.querySelector('.product-unit').value = p.qtyUnit;
        row.querySelector('.product-rate-unit').value = p.rateUnit;
    });

    syncAllToPreview();
    calculateTotals();
}

function renderPDFPage(doc, copyTitle) {
    const isBlueTheme = document.getElementById('themeSelector') && document.getElementById('themeSelector').value === 'blue';
    const themeColor = isBlueTheme ? [30, 64, 175] : [0, 0, 0];

    const pageSize = document.getElementById('invoiceSizeSelector').value || 'a4';
    const scale = pageSize === 'a5' ? 0.707 : 1.0;

    const margin = 10;
    const width = 190;
    const startY = 15;

    // Scaling helpers
    const sX = (x) => x * scale;
    const sY = (y) => y * scale;
    const sV = (v) => v * scale;

    // Drawing helpers
    const line = (x1, y1, x2, y2) => {
        doc.setDrawColor(...themeColor);
        doc.line(sX(x1), sY(y1), sX(x2), sY(y2));
    };
    const rect = (x, y, w, h) => {
        doc.setDrawColor(...themeColor);
        doc.setLineWidth(0.3 * scale);
        doc.rect(sX(x), sY(y), sV(w), sV(h));
    };
    const drawText = (txt, x, y, options = {}) => {
        const currentSize = doc.getFontSize();
        doc.setFontSize(currentSize * scale);
        doc.text(txt, sX(x), sY(y), options);
        doc.setFontSize(currentSize);
    };

    // Header
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(...themeColor);
    drawText("INVOICE", 105, startY, { align: 'center' });

    doc.setFontSize(8);
    doc.setTextColor(120);
    drawText(`(${copyTitle})`, 195, startY, { align: 'right' });

    // Main Box
    rect(margin, startY + 5, width, 265);

    // Row 1: Seller and Invoice
    line(margin, startY + 55, margin + width, startY + 55);
    line(115, startY + 5, 115, startY + 55);

    doc.setTextColor(...themeColor);
    doc.setFontSize(9);
    drawText("SELLER", margin + 2, startY + 10);

    doc.setTextColor(15, 23, 42);
    doc.setFontSize(12);
    drawText(getFieldVal('sellerName'), margin + 2, startY + 16);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(0);

    const sellerAddr = doc.splitTextToSize("Address : " + getFieldVal('sellerAddress'), 100 * scale);
    drawText(sellerAddr, margin + 2, startY + 22);

    let currentY = startY + 22 + (sellerAddr.length * 4);
    drawText("GSTIN NO : " + getFieldVal('sellerGST'), margin + 2, currentY);
    drawText("State Code : " + getFieldVal('sellerStateCode'), margin + 2, currentY + 4);
    doc.setFont('helvetica', 'bold');
    drawText("Dispatch From : " + getFieldVal('dispatchFrom'), margin + 2, currentY + 9);

    // Invoice Info Area
    doc.setFontSize(9);
    doc.setTextColor(...themeColor);
    drawText("INVOICE DETAILS", 117, startY + 10);
    doc.setTextColor(0);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    drawText("Invoice No : " + getFieldVal('invoiceNo'), 117, startY + 16);
    drawText("Date       : " + formatDate(new Date(document.getElementById('invoiceDate').value)), 117, startY + 21);

    line(115, startY + 26, margin + width, startY + 26);
    doc.setTextColor(...themeColor);
    doc.setFont('helvetica', 'bold');
    drawText("TRANSPORT DETAILS", 117, startY + 31);
    doc.setTextColor(0);
    doc.setFont('helvetica', 'normal');
    drawText("Transport  : " + getFieldVal('transportName'), 117, startY + 36);
    drawText("Lorry No   : " + getFieldVal('lorryNo'), 117, startY + 41);
    drawText("Bilty No   : " + getFieldVal('biltyNo'), 117, startY + 46);

    // Row 2: Bill To and Ship To
    line(margin, startY + 95, margin + width, startY + 95);
    line(115, startY + 55, 115, startY + 95);

    doc.setTextColor(...themeColor);
    doc.setFont('helvetica', 'bold');
    drawText("BILL TO (RECIPIENT)", margin + 2, startY + 60);
    doc.setTextColor(0);
    doc.setFont('helvetica', 'normal');
    drawText("Name : " + getFieldVal('buyerName'), margin + 2, startY + 66);
    const buyerAddr = doc.splitTextToSize("Address : " + getFieldVal('buyerAddress'), 95 * scale);
    drawText(buyerAddr, margin + 2, startY + 71);
    drawText("GSTIN: " + getFieldVal('buyerGST'), margin + 2, startY + 86);
    drawText("State: " + getFieldVal('buyerStateCode'), margin + 2, startY + 91);

    const isSame = sameAsBillTo.checked;
    const sName = isSame ? getFieldVal('buyerName') : getFieldVal('shipToName');
    const sAddr = isSame ? getFieldVal('buyerAddress') : getFieldVal('shipToAddress');
    const sGST = isSame ? getFieldVal('buyerGST') : getFieldVal('shipToGST');
    const sState = isSame ? getFieldVal('buyerStateCode') : getFieldVal('shipToStateCode');

    doc.setTextColor(...themeColor);
    doc.setFont('helvetica', 'bold');
    drawText("SHIP TO (CONSIGNEE)", 117, startY + 60);
    doc.setTextColor(0);
    doc.setFont('helvetica', 'normal');
    drawText("Name : " + sName, 117, startY + 66);
    const shipToAddrArr = doc.splitTextToSize("Address : " + sAddr, 80 * scale);
    drawText(shipToAddrArr, 117, startY + 71);
    drawText("GSTIN: " + sGST, 117, startY + 86);
    drawText("State: " + sState, 117, startY + 91);

    // Row 3: Product Table
    const tableData = Array.from(productBody.children).map((row, index) => {
        const descEl = row.querySelector('.product-desc');
        const hsnEl = row.querySelector('.product-hsn');
        const desc = descEl.classList.contains('rich-editable') ? descEl.innerText : descEl.value;
        const hsn = hsnEl.classList.contains('rich-editable') ? hsnEl.innerText : hsnEl.value;
        const qty = parseFloat(row.querySelector('.product-qty').value) || 0;
        const rate = parseFloat(row.querySelector('.product-rate').value) || 0;
        const qtyUnit = row.querySelector('.product-unit').value;
        const rateUnit = row.querySelector('.product-rate-unit').value;
        return [
            index + 1,
            desc,
            hsn,
            formatNumber(qty) + "\n" + qtyUnit,
            formatNumber(rate) + "\n" + rateUnit,
            formatNumber(qty * rate)
        ];
    });

    doc.autoTable({
        head: [["S.N", "Description of Goods", "HSN/SAC", "Quantity", "Rate", "Amount"]],
        body: tableData,
        startY: sY(startY + 95),
        margin: { left: sX(margin), right: sX(margin) },
        theme: 'grid',
        styles: {
            fontSize: 8 * scale,
            font: 'helvetica',
            cellPadding: 2 * scale,
            lineColor: themeColor,
            lineWidth: 0.1 * scale
        },
        headStyles: { fillColor: themeColor, textColor: 255 },
        columnStyles: {
            0: { cellWidth: 10 * scale, halign: 'center' },
            2: { cellWidth: 25 * scale, halign: 'center' },
            3: { cellWidth: 25 * scale, halign: 'center' },
            4: { cellWidth: 30 * scale, halign: 'center' },
            5: { cellWidth: 30 * scale, halign: 'right' }
        }
    });

    // Summary Section - AFTER table
    let currentPos = doc.lastAutoTable.finalY;
    const summaryHeight = 8;
    const totalY = Math.max(currentPos, startY + 185);

    // Grid lines for bottom section
    line(margin, totalY, margin + width, totalY);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);

    const drawRow = (label, value, y) => {
        line(150, y, 200, y);
        drawText(label, 153, y + 5);
        drawText(value, 198, y + 5, { align: 'right' });
    };

    drawRow("SUB TOTAL", document.getElementById('preview-subtotal').innerText, totalY);
    drawRow(`CGST ${getFieldVal('cgst')}%`, document.getElementById('preview-cgst-val').innerText, totalY + 7);
    drawRow(`SGST ${getFieldVal('sgst')}%`, document.getElementById('preview-sgst-val').innerText, totalY + 14);
    drawRow(`IGST ${getFieldVal('igst')}%`, document.getElementById('preview-igst-val').innerText, totalY + 21);

    // Grand Total Row
    const grandY = totalY + 28;
    doc.setFillColor(...themeColor);
    doc.rect(sX(150), sY(grandY), sV(50), sV(9), 'F');
    doc.setTextColor(255);
    drawText("TOTAL AMOUNT", 153, grandY + 6);
    drawText(document.getElementById('preview-gross-total').innerText, 198, grandY + 6, { align: 'right' });

    // Amount in Words
    doc.setTextColor(0);
    doc.setFontSize(8);
    line(margin, totalY + 38, margin + width, totalY + 38);
    drawText("Amount Chargeable (in words):", margin + 2, totalY + 43);
    doc.setFont('helvetica', 'bold');
    drawText(document.getElementById('preview-amount-words').innerText, margin + 45, totalY + 43);

    // Bank & Signature
    const bottomBoxY = totalY + 50;
    line(margin, bottomBoxY, margin + width, bottomBoxY);
    line(115, bottomBoxY, 115, startY + 270);

    doc.setFontSize(9);
    doc.setTextColor(...themeColor);
    drawText("BANK DETAILS", margin + 2, bottomBoxY + 6);

    doc.setTextColor(0);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    drawText("A/C Holder’s Name : " + getFieldVal('bankHolder'), margin + 2, bottomBoxY + 12);
    drawText("Bank Name : " + getFieldVal('bankName'), margin + 2, bottomBoxY + 17);
    drawText("A/C No.    : " + getFieldVal('bankAccount'), margin + 2, bottomBoxY + 22);
    drawText("IFSC Code : " + getFieldVal('bankIFSC'), margin + 2, bottomBoxY + 27);

    doc.setTextColor(...themeColor);
    doc.setFont('helvetica', 'bold');
    drawText("FOR AVNEESH CORPORATION", 157, bottomBoxY + 6, { align: 'center' });
    doc.setTextColor(0);
    doc.setFontSize(7);
    drawText("Authorized Signatory", 157, startY + 265, { align: 'center' });

    // Footer
    doc.setFontSize(7);
    doc.setTextColor(150);
    drawText("NOTE : ALL SUBJECT TO INDORE JURISDICTION", 105, startY + 273, { align: 'center' });
}
