/* ======================================
   EXCEL SHORTCUT MASTER — APP JS
   With animated visual illustrations
   ====================================== */

// ===== SHORTCUT DATA =====
const SHORTCUTS = [
    { title: "Paste", key: "Ctrl+V", category: "Clipboard", mustHave: true },
    { title: "Copy special", key: "Ctrl+C", category: "Clipboard", mustHave: true },
    { title: "Cut", key: "Ctrl+X", category: "Clipboard", mustHave: true },
    { title: "Copy the format right in a range", key: "Ctrl+R", category: "Clipboard", mustHave: true },
    { title: "Copy the format down in a range", key: "Ctrl+D", category: "Clipboard", mustHave: true },
    { title: "Copy formulas", key: "Ctrl+Enter", category: "Clipboard", mustHave: true },
    { title: "Cycle absolute and relative references", key: "F4", category: "Formulas", mustHave: true },
    { title: "Sum rows", key: "Alt+=", category: "Formulas", mustHave: true },
    { title: "Open data validation list", key: "Alt+Down", category: "Data", mustHave: true },
    { title: "Insert or Modify comment", key: "Shift+F2", category: "Insert" },
    { title: "Clear range values", key: "Del", category: "Editing", mustHave: true },
    { title: "Undo last action", key: "Ctrl+Z", category: "Editing", mustHave: true },
    { title: "Redo last action", key: "Ctrl+Y / F4", category: "Editing" },
    { title: "Open 'Name Range' dialog", key: "Ctrl+F3", category: "Formulas" },
    { title: "Open file", key: "Ctrl+O", category: "File" },
    { title: "Print", key: "Ctrl+P", category: "File" },
    { title: "Save", key: "Ctrl+S", category: "File" },
    { title: "Group rows or columns", key: "Alt+D+G+G", category: "Data", mustHave: true },
    { title: "Ungroup rows or columns", key: "Alt+D+G+U", category: "Data", mustHave: true },
    { title: "Turn on/off autofilter", key: "Alt+A T", category: "Data", mustHave: true },
    { title: "Show all (remove existing filters)", key: "Alt+A C", category: "Data" },
    { title: "Find", key: "Ctrl+F", category: "Editing", mustHave: true },
    { title: "Select and show A1", key: "Ctrl+Home", category: "Navigation", mustHave: true },
    { title: "Select last used cell", key: "Ctrl+End", category: "Navigation", mustHave: true },
    { title: "Select range", key: "Shift+Arrow", category: "Selection", mustHave: true },
    { title: "Select entire column", key: "Ctrl+Space", category: "Selection", mustHave: true },
    { title: "Select rows", key: "Ctrl+Shift", category: "Selection" },
    { title: "Select visible cells only", key: "Alt+;", category: "Selection", mustHave: true },
    { title: "Select all surrounding cells", key: "Ctrl+A", category: "Selection", mustHave: true },
    { title: "Replace", key: "Ctrl+H", category: "Editing" },
    { title: "Scientific number format", key: "Ctrl+^", category: "Formatting" },
    { title: "Cell format", key: "Ctrl+Shift+1", category: "Formatting", mustHave: true },
    { title: "Number format", key: "Ctrl+!", category: "Formatting", mustHave: true },
    { title: "Cycle heading", key: "Ctrl+Alt+H", category: "Formatting", mustHave: true },
    { title: "Cycle style", key: "Ctrl+Alt+S", category: "Formatting", mustHave: true },
    { title: "Clear formats", key: "Alt+E+A+S", category: "Formatting", mustHave: true },
    { title: "Currency format", key: "Ctrl+$", category: "Formatting" },
    { title: "Italic", key: "Ctrl+I", category: "Formatting" },
    { title: "Underline", key: "Ctrl+U", category: "Formatting" },
    { title: "Bold", key: "Ctrl+B", category: "Formatting" },
    { title: "Percentage format", key: "Ctrl+Shift+%", category: "Formatting", mustHave: true },
    { title: "Insert time", key: "Ctrl+:", category: "Insert" },
    { title: "Delete selected cells", key: "Ctrl+-", category: "Insert", mustHave: true },
    { title: "Insert table", key: "Ctrl+T", category: "Insert", mustHave: true },
    { title: "Insert sheet", key: "Shift+F11", category: "Insert" },
    { title: "Insert selected cells", key: "Ctrl++", category: "Insert", mustHave: true },
    { title: "Insert date", key: "Ctrl+;", category: "Insert" },
    { title: "Insert range name", key: "F3", category: "Formulas" },
    { title: "Move to right sheet", key: "Ctrl+Page Down", category: "Navigation", mustHave: true },
    { title: "Move to left sheet", key: "Ctrl+Page Up", category: "Navigation", mustHave: true },
    { title: "Open contextual menu (right-click)", key: "Shift+F10", category: "Navigation" },
    { title: "Move to edge of data region", key: "Ctrl+Arrow", category: "Navigation", mustHave: true },
    { title: "Open 'Go To' dialog", key: "Ctrl+G", category: "Navigation" },
    { title: "Spellcheck", key: "F7", category: "Editing" },
    { title: "Delete sheet", key: "Alt+E+L", category: "Sheet" },
    { title: "Move or copy sheet", key: "Alt+E+M", category: "Sheet" },
    { title: "Change sheet name", key: "Alt+H O R", category: "Sheet" },
    { title: "Show cell content", key: "F2", category: "Editing", mustHave: true },
    { title: "Hide or show ribbon", key: "Ctrl+F1", category: "View" },
    { title: "Open zoom dialog", key: "Alt+V+Z", category: "View" },
    { title: "Close application", key: "Alt+F4", category: "File" },
    { title: "Lock computer", key: "Window+L", category: "System" },
    { title: "Open Windows Explorer", key: "Window+E", category: "System" },
    { title: "Close window", key: "Ctrl+W", category: "File" },
    { title: "Show desktop", key: "Window+D", category: "System" },
    { title: "Open this list of keyboard shortcuts", key: "Ctrl+Alt+Q", category: "System" },
    { title: "Switch between open items", key: "Alt+Tab", category: "System" },
    { title: "New File", key: "Ctrl+N", category: "File", mustHave: true },
    { title: "Save As", key: "F12", category: "File" },
    { title: "Move one screen up/down", key: "PageUp / PageDown", category: "Navigation" },
    { title: "Move one screen left/right", key: "Alt+PageUp / PageDown", category: "Navigation" },
    { title: "Select between current cell and navigation result", key: "Shift+Navigation key", category: "Selection" },
    { title: "Select entire row", key: "Shift+Spacebar", category: "Selection", mustHave: true },
    { title: "Delete", key: "Alt+H, D, D", category: "Editing" },
    { title: "Insert Row", key: "Alt+H, I, R", category: "Insert", mustHave: true },
    { title: "Insert Column", key: "Alt+H, I, C", category: "Insert", mustHave: true },
    { title: "Hide column/row", key: "Ctrl+0 / 9", category: "View", mustHave: true },
    { title: "Unhide column/row", key: "Ctrl+Shift+0 / 9", category: "View" },
    { title: "Calculate entire file", key: "F9", category: "Formulas" },
    { title: "Calculate active sheet only", key: "Shift+F9", category: "Formulas" },
    { title: "Toggle formula bar", key: "Ctrl+Shift+U", category: "View" },
    { title: "Paste Special", key: "Ctrl+Alt+V", category: "Clipboard", mustHave: true },
    { title: "Copy exact formula from above", key: "Ctrl+'", category: "Formulas" },
    { title: "Insert PivotTable", key: "Alt+N V T", category: "Data" },
    { title: "Format Cells", key: "Ctrl+1", category: "Formatting", mustHave: true },
    { title: "General number format", key: "Ctrl+Shift+~", category: "Formatting" },
    { title: "Format Font Dialog Box", key: "Ctrl+Shift+F", category: "Formatting" },
    { title: "Auto fit column", key: "Alt+H, O, I", category: "Formatting" },
    { title: "Auto fit row", key: "Alt+H, O, A", category: "Formatting" },
    { title: "Decrease decimals", key: "Alt+H, 9", category: "Formatting" },
    { title: "Increase decimals", key: "Alt+H, 0", category: "Formatting" }
];

const CATEGORIES = [...new Set(SHORTCUTS.map(s => s.category))];

// ===== VISUAL GENERATORS =====
function visTitleBar(label) {
    return `<div class="vis-title-bar"><span class="dot"></span><span class="dot"></span><span class="dot"></span><span class="bar-text">${label || 'Excel'}</span></div>`;
}

// Helper: draw a mini spreadsheet grid
function miniGrid(opts = {}) {
    const { cols = 4, rows = 3, startX = 8, startY = 4, cellW = 38, cellH = 16, headers = true } = opts;
    let html = '';
    if (headers) {
        const colLabels = ['A','B','C','D','E','F'];
        for (let c = 0; c < cols; c++) {
            html += `<div class="vis-cell-header" style="left:${startX + 14 + c * cellW}px;top:${startY}px;width:${cellW}px;height:${cellH - 2}px">${colLabels[c]}</div>`;
        }
        for (let r = 0; r < rows; r++) {
            html += `<div class="vis-cell-header" style="left:${startX}px;top:${startY + (cellH - 2) + r * cellH}px;width:14px;height:${cellH}px">${r + 1}</div>`;
        }
    }
    for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
            html += `<div class="vis-cell" style="left:${startX + 14 + c * cellW}px;top:${startY + (cellH - 2) + r * cellH}px;width:${cellW}px;height:${cellH}px" data-r="${r}" data-c="${c}"></div>`;
        }
    }
    return html;
}

const VISUALS = {
    // === CLIPBOARD ===
    "Paste": () => `<div class="vis">${visTitleBar('Clipboard')}
        <div class="vis-canvas">
            ${miniGrid()}
            <div class="vis-selection" style="left:60px;top:20px;width:38px;height:16px;animation:pasteAppear 4s ease-in-out infinite;background:rgba(34,81,255,.12)"></div>
            <div style="position:absolute;left:65px;top:22px;font-size:6.5px;color:var(--accent);font-weight:600;animation:pasteAppear 4s ease-in-out infinite">Data</div>
            <div style="position:absolute;right:12px;top:70px;font-size:14px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#128203;</div>
            <span class="vis-label">PASTE</span>
        </div></div>`,

    "Copy special": () => `<div class="vis">${visTitleBar('Clipboard')}
        <div class="vis-canvas">
            ${miniGrid()}
            <div class="vis-selection" style="left:60px;top:20px;width:38px;height:16px;animation:cellCopy 4s ease-in-out infinite"></div>
            <div style="position:absolute;left:65px;top:22px;font-size:6.5px;color:#333">Sales</div>
            <div style="position:absolute;right:15px;top:65px;text-align:center">
                <div style="font-size:14px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#128464;</div>
                <div style="font-size:5px;color:#888">Copied</div>
            </div>
            <span class="vis-label">COPY</span>
        </div></div>`,

    "Cut": () => `<div class="vis">${visTitleBar('Clipboard')}
        <div class="vis-canvas">
            ${miniGrid()}
            <div style="position:absolute;left:65px;top:22px;font-size:6.5px;color:#333;animation:cutFade 4s ease-in-out infinite">Data</div>
            <div class="vis-selection" style="left:60px;top:20px;width:38px;height:16px;animation:cellCopy 4s ease-in-out infinite"></div>
            <div style="position:absolute;right:15px;top:65px;font-size:14px;color:#E07A20;animation:visPulse 4s ease-in-out infinite">&#9986;</div>
            <span class="vis-label">CUT</span>
        </div></div>`,

    "Copy the format right in a range": () => `<div class="vis">${visTitleBar('Clipboard')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:20px;width:50px;height:16px;background:#E8EEFF;border:1px solid var(--accent)"></div>
            <div style="position:absolute;left:77px;top:22px;font-size:6px;color:var(--accent);font-weight:600">100</div>
            <div style="position:absolute;left:122px;top:20px;width:50px;height:16px;animation:pasteAppear 4s ease-in-out infinite;background:#E8EEFF;border:1px solid var(--accent)"></div>
            <div style="position:absolute;left:127px;top:22px;font-size:6px;color:var(--accent);font-weight:600;animation:pasteAppear 4s ease-in-out infinite">100</div>
            <div style="position:absolute;left:100px;top:48px;font-size:12px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#8594;</div>
            <span class="vis-label">FILL RIGHT</span>
        </div></div>`,

    "Copy the format down in a range": () => `<div class="vis">${visTitleBar('Clipboard')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:20px;width:50px;height:16px;background:#E8EEFF;border:1px solid var(--accent)"></div>
            <div style="position:absolute;left:77px;top:22px;font-size:6px;color:var(--accent);font-weight:600">100</div>
            <div style="position:absolute;left:72px;top:36px;width:50px;height:16px;animation:pasteAppear 4s ease-in-out infinite;background:#E8EEFF;border:1px solid var(--accent)"></div>
            <div style="position:absolute;left:77px;top:38px;font-size:6px;color:var(--accent);font-weight:600;animation:pasteAppear 4s ease-in-out infinite">100</div>
            <div style="position:absolute;left:140px;top:32px;font-size:12px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#8595;</div>
            <span class="vis-label">FILL DOWN</span>
        </div></div>`,

    "Copy formulas": () => `<div class="vis">${visTitleBar('Formulas')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 4, cellW: 48})}
            <div style="position:absolute;left:27px;top:22px;font-size:5.5px;color:#333">10</div>
            <div style="position:absolute;left:27px;top:38px;font-size:5.5px;color:#333">20</div>
            <div style="position:absolute;left:27px;top:54px;font-size:5.5px;color:#333">30</div>
            <div style="position:absolute;left:75px;top:22px;font-size:5px;color:var(--accent);font-weight:600;animation:pasteAppear 4s ease-in-out infinite">=A1*2</div>
            <div style="position:absolute;left:75px;top:38px;font-size:5px;color:var(--accent);font-weight:600;animation:pasteAppear 4s ease-in-out infinite .2s">=A2*2</div>
            <div style="position:absolute;left:75px;top:54px;font-size:5px;color:var(--accent);font-weight:600;animation:pasteAppear 4s ease-in-out infinite .4s">=A3*2</div>
            <span class="vis-label">CTRL+ENTER</span>
        </div></div>`,

    // === FORMULAS ===
    "Cycle absolute and relative references": () => `<div class="vis">${visTitleBar('Formulas')}
        <div class="vis-canvas">
            <div style="position:absolute;left:10px;top:4px;right:10px;height:16px;background:white;border:1px solid #D0D5DD;border-radius:2px;display:flex;align-items:center;padding:0 6px">
                <span style="font-size:5px;color:#888;margin-right:4px">fx</span>
                <span style="font-size:6px;font-family:monospace;color:var(--accent);font-weight:600">=SUM(</span>
            </div>
            <div style="position:absolute;left:50%;top:38px;transform:translateX(-50%);text-align:center">
                <div style="display:flex;gap:6px;justify-content:center;margin-bottom:8px">
                    <div style="padding:4px 8px;background:#E8EEFF;border:1.5px solid var(--accent);border-radius:4px;font-size:8px;font-family:monospace;font-weight:700;color:var(--accent);animation:visPulse 4s ease-in-out infinite">A1</div>
                    <div style="font-size:10px;color:#888;display:flex;align-items:center">&#8594;</div>
                    <div style="padding:4px 8px;background:#E8EEFF;border:1.5px solid var(--accent);border-radius:4px;font-size:8px;font-family:monospace;font-weight:700;color:var(--accent);animation:pasteAppear 4s ease-in-out infinite">$A$1</div>
                </div>
                <div style="display:flex;gap:6px;justify-content:center;margin-bottom:4px">
                    <div style="padding:3px 6px;background:#f5f5f5;border:1px solid #ddd;border-radius:3px;font-size:6px;font-family:monospace;color:#666;animation:pasteAppear 4s ease-in-out infinite .3s">A$1</div>
                    <div style="padding:3px 6px;background:#f5f5f5;border:1px solid #ddd;border-radius:3px;font-size:6px;font-family:monospace;color:#666;animation:pasteAppear 4s ease-in-out infinite .5s">$A1</div>
                </div>
                <div style="font-size:5.5px;color:#888">Press F4 to cycle</div>
            </div>
            <span class="vis-label">ABSOLUTE REF</span>
        </div></div>`,

    "Sum rows": () => `<div class="vis">${visTitleBar('Formulas')}
        <div class="vis-canvas">
            ${miniGrid({cols: 4, rows: 3, cellW: 36})}
            <div style="position:absolute;left:27px;top:22px;font-size:5.5px;color:#333">10</div>
            <div style="position:absolute;left:63px;top:22px;font-size:5.5px;color:#333">20</div>
            <div style="position:absolute;left:99px;top:22px;font-size:5.5px;color:#333">30</div>
            <div style="position:absolute;left:135px;top:20px;width:36px;height:16px;background:#E8EEFF;border:2px solid var(--accent);border-radius:1px;animation:pasteAppear 4s ease-in-out infinite"></div>
            <div style="position:absolute;left:137px;top:22px;font-size:5.5px;color:var(--accent);font-weight:700;animation:sumFlash 4s ease-in-out infinite">60</div>
            <div style="position:absolute;right:12px;top:65px;font-size:8px;color:var(--accent);font-weight:700">&#931; SUM</div>
            <span class="vis-label">AUTOSUM</span>
        </div></div>`,

    "Open data validation list": () => `<div class="vis">${visTitleBar('Data')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:20px;width:50px;height:16px;border:2px solid var(--accent)">
                <div style="position:absolute;right:0;top:0;width:12px;height:100%;background:var(--accent);display:flex;align-items:center;justify-content:center;font-size:7px;color:white">&#9660;</div>
            </div>
            <div style="position:absolute;left:72px;top:36px;width:50px;background:white;border:1px solid var(--accent);border-radius:0 0 3px 3px;box-shadow:0 2px 8px rgba(0,0,0,.1);z-index:2;animation:filterDrop 4s ease-in-out infinite;overflow:hidden">
                <div style="padding:3px 6px;font-size:5.5px;color:#333;background:#E8EEFF">Option A</div>
                <div style="padding:3px 6px;font-size:5.5px;color:#333">Option B</div>
                <div style="padding:3px 6px;font-size:5.5px;color:#333">Option C</div>
            </div>
            <span class="vis-label">VALIDATION LIST</span>
        </div></div>`,

    "Open 'Name Range' dialog": () => `<div class="vis">${visTitleBar('Formulas')}
        <div class="vis-canvas">
            <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:140px;background:white;border:1.5px solid var(--accent);border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.1);animation:dialogPop 4s ease-in-out infinite">
                <div style="background:var(--accent);padding:4px 8px;border-radius:4px 4px 0 0;font-size:6px;color:white;font-weight:600">Name Manager</div>
                <div style="padding:6px 8px">
                    <div style="display:flex;justify-content:space-between;font-size:5px;color:#666;margin-bottom:3px"><span>SalesData</span><span>A1:D10</span></div>
                    <div style="display:flex;justify-content:space-between;font-size:5px;color:#666;margin-bottom:3px"><span>Revenue</span><span>B2:B50</span></div>
                    <div style="display:flex;justify-content:space-between;font-size:5px;color:#666"><span>Costs</span><span>C2:C50</span></div>
                </div>
            </div>
            <span class="vis-label">NAME MANAGER</span>
        </div></div>`,

    "Insert range name": () => `<div class="vis">${visTitleBar('Formulas')}
        <div class="vis-canvas">
            <div style="position:absolute;left:10px;top:8px;right:10px;height:16px;background:white;border:1px solid #D0D5DD;border-radius:2px;display:flex;align-items:center;padding:0 6px">
                <span style="font-size:5px;color:#888;margin-right:4px">fx</span>
                <span style="font-size:6px;font-family:monospace;color:var(--accent)">=SUM(</span>
                <span style="font-size:6px;font-family:monospace;color:var(--accent);font-weight:700;animation:pasteAppear 4s ease-in-out infinite">SalesData</span>
                <span style="font-size:6px;font-family:monospace;color:var(--accent)">)</span>
            </div>
            <div style="position:absolute;left:30px;top:35px;width:90px;background:white;border:1px solid var(--accent);border-radius:3px;box-shadow:0 2px 8px rgba(0,0,0,.1);animation:insertAppear 4s ease-in-out infinite">
                <div style="padding:3px 6px;font-size:5.5px;color:var(--accent);font-weight:600;background:#E8EEFF">SalesData</div>
                <div style="padding:3px 6px;font-size:5.5px;color:#333">Revenue</div>
                <div style="padding:3px 6px;font-size:5.5px;color:#333">Costs</div>
            </div>
            <span class="vis-label">PASTE NAME</span>
        </div></div>`,

    // === EDITING ===
    "Insert or Modify comment": () => `<div class="vis">${visTitleBar('Insert')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:18px;width:0;height:0;border-left:6px solid transparent;border-bottom:6px solid transparent;border-top:6px solid #C62828;position:absolute;right:73px;top:18px"></div>
            <div style="position:absolute;left:100px;top:6px;width:80px;background:#FFFDE7;border:1px solid #FFD54F;border-radius:3px;padding:4px;box-shadow:2px 2px 6px rgba(0,0,0,.08);animation:insertAppear 4s ease-in-out infinite">
                <div style="font-size:5px;color:#F57F17;font-weight:600;margin-bottom:2px">Comment:</div>
                <div style="font-size:5px;color:#666">Check this value against Q3 data</div>
            </div>
            <span class="vis-label">COMMENT</span>
        </div></div>`,

    "Clear range values": () => `<div class="vis">${visTitleBar('Editing')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 3, cellW: 50})}
            <div style="position:absolute;left:27px;top:22px;font-size:5.5px;color:#333;animation:deleteContent 4s ease-in-out infinite">42</div>
            <div style="position:absolute;left:77px;top:22px;font-size:5.5px;color:#333;animation:deleteContent 4s ease-in-out infinite .1s">Sales</div>
            <div style="position:absolute;left:27px;top:38px;font-size:5.5px;color:#333;animation:deleteContent 4s ease-in-out infinite .2s">78</div>
            <div style="position:absolute;left:77px;top:38px;font-size:5.5px;color:#333;animation:deleteContent 4s ease-in-out infinite .3s">Revenue</div>
            <div style="position:absolute;left:22px;top:20px;width:105px;height:32px;border:2px solid var(--accent);background:rgba(34,81,255,.06)"></div>
            <span class="vis-label">DELETE VALUES</span>
        </div></div>`,

    "Undo last action": () => `<div class="vis">${visTitleBar('Editing')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center;gap:12px">
            <div style="text-align:center">
                <div style="font-size:28px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#8630;</div>
                <div style="font-size:7px;color:#888;margin-top:4px">Undo</div>
            </div>
            <div style="width:1px;height:40px;background:#ddd"></div>
            <div style="text-align:center">
                <div style="font-size:6.5px;color:#888;margin-bottom:3px">Restores previous</div>
                <div style="font-size:6.5px;color:#888">state of cell</div>
            </div>
            <span class="vis-label">UNDO</span>
        </div></div>`,

    "Redo last action": () => `<div class="vis">${visTitleBar('Editing')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center;gap:12px">
            <div style="text-align:center">
                <div style="font-size:28px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#8631;</div>
                <div style="font-size:7px;color:#888;margin-top:4px">Redo</div>
            </div>
            <div style="width:1px;height:40px;background:#ddd"></div>
            <div style="text-align:center">
                <div style="font-size:6.5px;color:#888;margin-bottom:3px">Re-applies last</div>
                <div style="font-size:6.5px;color:#888">undone action</div>
            </div>
            <span class="vis-label">REDO</span>
        </div></div>`,

    "Find": () => `<div class="vis">${visTitleBar('Editing')}
        <div class="vis-canvas">
            <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:140px;background:white;border:1.5px solid var(--accent);border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.1);animation:dialogPop 4s ease-in-out infinite">
                <div style="background:var(--accent);padding:4px 8px;border-radius:4px 4px 0 0;font-size:6px;color:white;font-weight:600">Find and Replace</div>
                <div style="padding:6px 8px">
                    <div style="font-size:5px;color:#666;margin-bottom:2px">Find what:</div>
                    <div style="width:100%;height:12px;background:#f5f5f5;border:1px solid #ddd;border-radius:2px;padding:0 4px;display:flex;align-items:center">
                        <span style="font-size:5.5px;color:var(--accent);font-weight:600">Revenue</span>
                        <div style="width:1px;height:7px;background:var(--accent);margin-left:1px;animation:visPulse 1s ease infinite"></div>
                    </div>
                </div>
            </div>
            <span class="vis-label">FIND</span>
        </div></div>`,

    "Replace": () => `<div class="vis">${visTitleBar('Editing')}
        <div class="vis-canvas">
            <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:140px;background:white;border:1.5px solid var(--accent);border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.1);animation:dialogPop 4s ease-in-out infinite">
                <div style="background:var(--accent);padding:4px 8px;border-radius:4px 4px 0 0;font-size:6px;color:white;font-weight:600">Find and Replace</div>
                <div style="padding:5px 8px">
                    <div style="font-size:5px;color:#666;margin-bottom:1px">Find:</div>
                    <div style="width:100%;height:11px;background:#f5f5f5;border:1px solid #ddd;border-radius:2px;padding:0 4px;display:flex;align-items:center;margin-bottom:3px">
                        <span style="font-size:5px;color:#C62828">Old Text</span>
                    </div>
                    <div style="font-size:5px;color:#666;margin-bottom:1px">Replace:</div>
                    <div style="width:100%;height:11px;background:#f5f5f5;border:1px solid #ddd;border-radius:2px;padding:0 4px;display:flex;align-items:center">
                        <span style="font-size:5px;color:var(--accent);font-weight:600">New Text</span>
                    </div>
                </div>
            </div>
            <span class="vis-label">REPLACE</span>
        </div></div>`,

    "Spellcheck": () => `<div class="vis">${visTitleBar('Editing')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center;animation:insertAppear 4s ease-in-out infinite">
                <div style="font-size:9px;color:#C62828;text-decoration:wavy underline #C62828;margin-bottom:8px">Recieve</div>
                <div style="font-size:10px;color:var(--accent)">&#8595;</div>
                <div style="font-size:9px;color:var(--accent);font-weight:600;margin-top:4px">Receive &#10003;</div>
            </div>
            <span class="vis-label">SPELLCHECK</span>
        </div></div>`,

    "Show cell content": () => `<div class="vis">${visTitleBar('Editing')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:20px;width:50px;height:16px;border:2px solid var(--accent);background:rgba(34,81,255,.08)"></div>
            <div style="position:absolute;left:77px;top:22px;font-size:6px;color:#333">42</div>
            <div style="position:absolute;left:10px;top:62px;right:10px;height:16px;background:white;border:1px solid var(--accent);border-radius:2px;display:flex;align-items:center;padding:0 6px;animation:insertAppear 4s ease-in-out infinite">
                <span style="font-size:5px;color:#888;margin-right:4px">fx</span>
                <span style="font-size:6px;font-family:monospace;color:var(--accent)">=A1+B1*2</span>
                <div style="width:1px;height:8px;background:var(--accent);margin-left:1px;animation:visPulse 1s ease infinite"></div>
            </div>
            <div style="position:absolute;left:10px;top:82px;font-size:5px;color:#888">F2 enters edit mode</div>
            <span class="vis-label">EDIT CELL</span>
        </div></div>`,

    // === DATA ===
    "Group rows or columns": () => `<div class="vis">${visTitleBar('Data')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 4, cellW: 42, startX: 20})}
            <div style="position:absolute;left:8px;top:18px;width:10px;height:64px;border:1.5px solid var(--accent);border-right:none;border-radius:3px 0 0 3px;animation:insertAppear 4s ease-in-out infinite"></div>
            <div style="position:absolute;left:3px;top:46px;font-size:8px;color:var(--accent);font-weight:700;animation:visPulse 4s ease-in-out infinite">&#8722;</div>
            <div style="position:absolute;right:12px;top:70px;font-size:6px;color:#888">Rows grouped</div>
            <span class="vis-label">GROUP</span>
        </div></div>`,

    "Ungroup rows or columns": () => `<div class="vis">${visTitleBar('Data')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 4, cellW: 42, startX: 20})}
            <div style="position:absolute;left:8px;top:18px;width:10px;height:64px;border:1.5px solid #aaa;border-right:none;border-radius:3px 0 0 3px;animation:removeItem 4s ease-in-out infinite"></div>
            <div style="position:absolute;left:3px;top:46px;font-size:8px;color:#aaa;animation:removeItem 4s ease-in-out infinite">&#43;</div>
            <div style="position:absolute;right:12px;top:70px;font-size:6px;color:#888">Ungrouped</div>
            <span class="vis-label">UNGROUP</span>
        </div></div>`,

    "Insert or remove auto filter": () => `<div class="vis">${visTitleBar('Data')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 3, cellW: 50})}
            <div style="position:absolute;left:27px;top:22px;font-size:5.5px;color:var(--accent);font-weight:600">Name &#9660;</div>
            <div style="position:absolute;left:77px;top:22px;font-size:5.5px;color:var(--accent);font-weight:600">Sales &#9660;</div>
            <div style="position:absolute;left:127px;top:22px;font-size:5.5px;color:var(--accent);font-weight:600">Date &#9660;</div>
            <div style="position:absolute;left:22px;top:20px;width:155px;height:16px;background:rgba(34,81,255,.08);animation:cellHighlight 4s ease-in-out infinite"></div>
            <span class="vis-label">AUTO FILTER</span>
        </div></div>`,

    "Show all rows in autofiltered table": () => `<div class="vis">${visTitleBar('Data')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 4, cellW: 50})}
            <div style="position:absolute;left:27px;top:22px;font-size:5.5px;color:var(--accent);font-weight:600">Name &#9660;</div>
            <div style="position:absolute;left:27px;top:38px;font-size:5.5px;color:#333">Alice</div>
            <div style="position:absolute;left:27px;top:54px;font-size:5.5px;color:#888;animation:pasteAppear 4s ease-in-out infinite">Bob</div>
            <div style="position:absolute;left:27px;top:70px;font-size:5.5px;color:#888;animation:pasteAppear 4s ease-in-out infinite .2s">Carol</div>
            <div style="position:absolute;right:12px;top:75px;font-size:6px;color:var(--accent);font-weight:600">All shown</div>
            <span class="vis-label">SHOW ALL</span>
        </div></div>`,

    // === FILE ===
    "Open file": () => `<div class="vis">${visTitleBar('File')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center;animation:insertAppear 4s ease-in-out infinite">
                <div style="font-size:28px">&#128194;</div>
                <div style="font-size:7px;color:#888;margin-top:4px">Open file dialog</div>
            </div>
            <span class="vis-label">OPEN</span>
        </div></div>`,

    "Print": () => `<div class="vis">${visTitleBar('File')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center;gap:10px">
            <div style="width:60px;height:70px;background:white;border:1px solid #ddd;border-radius:3px;padding:4px;animation:insertAppear 4s ease-in-out infinite">
                <div style="width:100%;height:3px;background:#E2E8DB;margin-bottom:2px;border-radius:1px"></div>
                <div style="width:80%;height:3px;background:#E2E8DB;margin-bottom:2px;border-radius:1px"></div>
                <div style="width:90%;height:3px;background:#E2E8DB;margin-bottom:2px;border-radius:1px"></div>
                <div style="width:60%;height:3px;background:#E2E8DB;border-radius:1px"></div>
            </div>
            <div style="font-size:22px;color:var(--accent)">&#128424;</div>
            <span class="vis-label">PRINT</span>
        </div></div>`,

    "Save": () => `<div class="vis">${visTitleBar('File')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center;animation:insertAppear 4s ease-in-out infinite">
                <div style="font-size:28px;color:var(--accent)">&#128190;</div>
                <div style="font-size:7px;color:#888;margin-top:4px">File saved</div>
            </div>
            <span class="vis-label">SAVE</span>
        </div></div>`,

    "Close application": () => `<div class="vis">${visTitleBar('File')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="width:120px;height:70px;background:white;border:1.5px solid #ddd;border-radius:6px;position:relative;animation:removeItem 4s ease-in-out infinite">
                <div style="height:14px;background:linear-gradient(90deg,#061F79,#2251FF);border-radius:4px 4px 0 0;display:flex;align-items:center;padding:0 4px">
                    <span style="font-size:5px;color:white">Excel</span>
                    <div style="margin-left:auto;width:10px;height:8px;background:#C62828;border-radius:1px;display:flex;align-items:center;justify-content:center;font-size:6px;color:white;font-weight:700">&#215;</div>
                </div>
            </div>
            <span class="vis-label">CLOSE APP</span>
        </div></div>`,

    "Close window": () => `<div class="vis">${visTitleBar('File')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center;animation:removeItem 4s ease-in-out infinite">
                <div style="width:80px;height:50px;background:white;border:1.5px solid #ddd;border-radius:4px;margin-bottom:4px">
                    <div style="height:10px;background:#E2E8DB;border-radius:3px 3px 0 0;display:flex;align-items:center;justify-content:flex-end;padding-right:3px">
                        <div style="font-size:5px;color:#C62828;font-weight:700">&#215;</div>
                    </div>
                </div>
                <div style="font-size:6px;color:#888">Window closes</div>
            </div>
            <span class="vis-label">CLOSE WINDOW</span>
        </div></div>`,

    // === NAVIGATION ===
    "Select and show A1": () => `<div class="vis">${visTitleBar('Navigation')}
        <div class="vis-canvas">
            ${miniGrid({cols: 4, rows: 4, cellW: 34})}
            <div class="vis-selection" style="width:34px;height:16px;animation:jumpToA1 4s ease-in-out infinite;background:rgba(34,81,255,.12)"></div>
            <div style="position:absolute;right:10px;top:75px;font-size:7px;color:var(--accent);font-weight:700;animation:visPulse 4s ease-in-out infinite">&#8598; A1</div>
            <span class="vis-label">GO TO A1</span>
        </div></div>`,

    "Select last used cell": () => `<div class="vis">${visTitleBar('Navigation')}
        <div class="vis-canvas">
            ${miniGrid({cols: 4, rows: 4, cellW: 34})}
            <div class="vis-selection" style="width:34px;height:16px;animation:jumpToEnd 4s ease-in-out infinite;background:rgba(34,81,255,.12)"></div>
            <div style="position:absolute;right:10px;top:75px;font-size:7px;color:var(--accent);font-weight:700;animation:visPulse 4s ease-in-out infinite">&#8600; End</div>
            <span class="vis-label">GO TO END</span>
        </div></div>`,

    "Move to right sheet": () => makeSheetNavVisual('right'),
    "Move to left sheet": () => makeSheetNavVisual('left'),

    "Open contextual menu (right-click)": () => `<div class="vis">${visTitleBar('Navigation')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, cellW: 50})}
            <div style="position:absolute;left:80px;top:32px;width:70px;background:white;border:1px solid #ccc;border-radius:4px;box-shadow:2px 2px 8px rgba(0,0,0,.12);animation:insertAppear 4s ease-in-out infinite;z-index:2">
                <div style="padding:3px 8px;font-size:5px;color:#333;border-bottom:1px solid #eee">Cut</div>
                <div style="padding:3px 8px;font-size:5px;color:#333;border-bottom:1px solid #eee">Copy</div>
                <div style="padding:3px 8px;font-size:5px;color:#333;border-bottom:1px solid #eee">Paste Special</div>
                <div style="padding:3px 8px;font-size:5px;color:#333">Insert...</div>
            </div>
            <span class="vis-label">CONTEXT MENU</span>
        </div></div>`,

    "Move to edge of data region": () => `<div class="vis">${visTitleBar('Navigation')}
        <div class="vis-canvas">
            ${miniGrid({cols: 4, rows: 4, cellW: 34})}
            <div class="vis-selection" style="left:22px;top:18px;width:34px;height:16px;animation:moveSelection 4s ease-in-out infinite;background:rgba(34,81,255,.12)"></div>
            <div style="position:absolute;left:80px;top:60px;font-size:16px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#8680;</div>
            <span class="vis-label">JUMP TO EDGE</span>
        </div></div>`,

    "Open 'Go To' dialog": () => `<div class="vis">${visTitleBar('Navigation')}
        <div class="vis-canvas">
            <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:120px;background:white;border:1.5px solid var(--accent);border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.1);animation:dialogPop 4s ease-in-out infinite">
                <div style="background:var(--accent);padding:4px 8px;border-radius:4px 4px 0 0;font-size:6px;color:white;font-weight:600">Go To</div>
                <div style="padding:6px 8px">
                    <div style="font-size:5px;color:#666;margin-bottom:2px">Reference:</div>
                    <div style="width:100%;height:12px;background:#f5f5f5;border:1px solid #ddd;border-radius:2px;padding:0 4px;display:flex;align-items:center">
                        <span style="font-size:5.5px;color:var(--accent)">A1:D100</span>
                    </div>
                </div>
            </div>
            <span class="vis-label">GO TO</span>
        </div></div>`,

    // === SELECTION ===
    "Select range": () => `<div class="vis">${visTitleBar('Selection')}
        <div class="vis-canvas">
            ${miniGrid({cols: 4, rows: 4, cellW: 34})}
            <div style="position:absolute;left:22px;top:18px;background:rgba(34,81,255,.12);border:2px solid var(--accent);animation:expandSelection 4s ease-in-out infinite"></div>
            <div style="position:absolute;right:10px;top:75px;font-size:6px;color:var(--accent);font-weight:600">Shift+Arrow</div>
            <span class="vis-label">SELECT RANGE</span>
        </div></div>`,

    "Select rows or columns": () => `<div class="vis">${visTitleBar('Selection')}
        <div class="vis-canvas">
            ${miniGrid({cols: 4, rows: 4, cellW: 34})}
            <div style="position:absolute;left:56px;top:18px;width:34px;height:64px;background:rgba(34,81,255,.15);border:2px solid var(--accent);animation:cellHighlight 4s ease-in-out infinite"></div>
            <div style="position:absolute;right:8px;top:75px;font-size:6px;color:var(--accent);font-weight:600">Column B</div>
            <span class="vis-label">SELECT COL</span>
        </div></div>`,

    "Select rows": () => `<div class="vis">${visTitleBar('Selection')}
        <div class="vis-canvas">
            ${miniGrid({cols: 4, rows: 4, cellW: 34})}
            <div style="position:absolute;left:22px;top:34px;width:136px;height:16px;background:rgba(34,81,255,.15);border:2px solid var(--accent);animation:cellHighlight 4s ease-in-out infinite"></div>
            <div style="position:absolute;right:8px;top:75px;font-size:6px;color:var(--accent);font-weight:600">Row 2</div>
            <span class="vis-label">SELECT ROW</span>
        </div></div>`,

    "Select visible cells only": () => `<div class="vis">${visTitleBar('Selection')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 4, cellW: 42, startX: 14})}
            <div style="position:absolute;left:28px;top:20px;width:126px;height:16px;background:rgba(34,81,255,.12);border:1px solid var(--accent)"></div>
            <div style="position:absolute;left:28px;top:52px;width:126px;height:16px;background:rgba(34,81,255,.12);border:1px solid var(--accent)"></div>
            <div style="position:absolute;left:28px;top:36px;width:126px;height:16px;background:#f0f0f0;border:1px dashed #aaa;display:flex;align-items:center;justify-content:center;font-size:5px;color:#aaa">hidden</div>
            <div style="position:absolute;right:8px;top:78px;font-size:5.5px;color:var(--accent);font-weight:600">Visible only</div>
            <span class="vis-label">VISIBLE CELLS</span>
        </div></div>`,

    "Select all surrounding cells": () => `<div class="vis">${visTitleBar('Selection')}
        <div class="vis-canvas">
            ${miniGrid({cols: 4, rows: 4, cellW: 34})}
            <div style="position:absolute;left:22px;top:18px;background:rgba(34,81,255,.12);border:2px solid var(--accent);animation:selectAll 4s ease-in-out infinite"></div>
            <div style="position:absolute;right:8px;top:78px;font-size:6px;color:var(--accent);font-weight:600">Select All</div>
            <span class="vis-label">CTRL+A</span>
        </div></div>`,

    // === FORMATTING ===
    "Bold": () => makeFontFormatVisual('bold'),
    "Italic": () => makeFontFormatVisual('italic'),
    "Underline": () => makeFontFormatVisual('underline'),

    "Scientific number format": () => `<div class="vis">${visTitleBar('Formatting')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center">
                <div style="font-size:10px;color:#888;margin-bottom:6px">1500000</div>
                <div style="font-size:10px;color:var(--accent)">&#8595;</div>
                <div style="font-size:10px;color:var(--accent);font-weight:700;font-family:monospace;animation:pasteAppear 4s ease-in-out infinite">1.50E+06</div>
            </div>
            <span class="vis-label">SCIENTIFIC</span>
        </div></div>`,

    "Cell format": () => `<div class="vis">${visTitleBar('Formatting')}
        <div class="vis-canvas">
            <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:150px;background:white;border:1.5px solid var(--accent);border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.1);animation:dialogPop 4s ease-in-out infinite">
                <div style="background:var(--accent);padding:4px 8px;border-radius:4px 4px 0 0;font-size:6px;color:white;font-weight:600">Format Cells</div>
                <div style="display:flex;gap:1px;padding:3px 6px;border-bottom:1px solid #eee">
                    <span style="font-size:4.5px;padding:2px 4px;background:#E8EEFF;border-radius:2px;color:var(--accent);font-weight:600">Number</span>
                    <span style="font-size:4.5px;padding:2px 4px;color:#888">Alignment</span>
                    <span style="font-size:4.5px;padding:2px 4px;color:#888">Font</span>
                    <span style="font-size:4.5px;padding:2px 4px;color:#888">Border</span>
                </div>
                <div style="padding:4px 8px;font-size:5px;color:#666">Number, Currency, Date, Text...</div>
            </div>
            <span class="vis-label">FORMAT CELLS</span>
        </div></div>`,

    "Number format": () => `<div class="vis">${visTitleBar('Formatting')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center">
                <div style="font-size:10px;color:#888;margin-bottom:6px">1234.5</div>
                <div style="font-size:10px;color:var(--accent)">&#8595;</div>
                <div style="font-size:10px;color:var(--accent);font-weight:700;font-family:monospace;animation:pasteAppear 4s ease-in-out infinite">1,234.50</div>
            </div>
            <span class="vis-label">NUMBER FMT</span>
        </div></div>`,

    "Cycle heading": () => `<div class="vis">${visTitleBar('Formatting')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center;gap:8px">
            <div style="display:flex;flex-direction:column;gap:4px;animation:visPulse 4s ease-in-out infinite">
                <div style="padding:3px 8px;background:var(--accent);border-radius:3px;font-size:6px;color:white;font-weight:700">Heading 1</div>
                <div style="padding:3px 8px;background:#4A73FF;border-radius:3px;font-size:5.5px;color:white;font-weight:600">Heading 2</div>
                <div style="padding:3px 8px;background:#6B8FFF;border-radius:3px;font-size:5px;color:white">Heading 3</div>
            </div>
            <div style="font-size:20px;color:var(--accent);animation:visPulse 2s ease infinite">&#8635;</div>
            <span class="vis-label">CYCLE HEADING</span>
        </div></div>`,

    "Cycle style": () => `<div class="vis">${visTitleBar('Formatting')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center;gap:8px">
            <div style="display:flex;gap:4px">
                <div style="width:30px;height:30px;background:var(--accent);border-radius:4px;animation:insertAppear 4s ease-in-out infinite"></div>
                <div style="width:30px;height:30px;background:#E07A20;border-radius:4px;animation:insertAppear 4s ease-in-out infinite .2s"></div>
                <div style="width:30px;height:30px;background:#1565C0;border-radius:4px;animation:insertAppear 4s ease-in-out infinite .4s"></div>
            </div>
            <div style="font-size:20px;color:var(--accent);animation:visPulse 2s ease infinite">&#8635;</div>
            <span class="vis-label">CYCLE STYLE</span>
        </div></div>`,

    "Clear formats": () => `<div class="vis">${visTitleBar('Formatting')}
        <div class="vis-canvas">
            <div style="position:absolute;left:20px;top:20px;text-align:center">
                <div style="font-size:6px;color:#888;margin-bottom:3px">BEFORE:</div>
                <div style="padding:4px 10px;background:#E8EEFF;border:2px solid var(--accent);border-radius:4px">
                    <span style="font-size:8px;font-weight:700;color:var(--accent);font-style:italic">$1,234</span>
                </div>
            </div>
            <div style="position:absolute;left:50%;top:42px;transform:translateX(-50%);font-size:12px;color:var(--accent)">&#10140;</div>
            <div style="position:absolute;right:20px;top:20px;text-align:center">
                <div style="font-size:6px;color:#888;margin-bottom:3px">AFTER:</div>
                <div style="padding:4px 10px;background:white;border:1px solid #ddd;border-radius:4px;animation:insertAppear 4s ease-in-out infinite">
                    <span style="font-size:8px;color:#333">1234</span>
                </div>
            </div>
            <span class="vis-label">CLEAR FMT</span>
        </div></div>`,

    "Currency format": () => `<div class="vis">${visTitleBar('Formatting')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center">
                <div style="font-size:10px;color:#888;margin-bottom:6px">1234.5</div>
                <div style="font-size:10px;color:var(--accent)">&#8595;</div>
                <div style="font-size:11px;color:var(--accent);font-weight:700;font-family:monospace;animation:pasteAppear 4s ease-in-out infinite">$1,234.50</div>
            </div>
            <span class="vis-label">CURRENCY</span>
        </div></div>`,

    "Percentage format": () => `<div class="vis">${visTitleBar('Formatting')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center">
                <div style="font-size:10px;color:#888;margin-bottom:6px">0.75</div>
                <div style="font-size:10px;color:var(--accent)">&#8595;</div>
                <div style="font-size:11px;color:var(--accent);font-weight:700;font-family:monospace;animation:pasteAppear 4s ease-in-out infinite">75%</div>
            </div>
            <span class="vis-label">PERCENTAGE</span>
        </div></div>`,

    // === INSERT ===
    "Insert time": () => `<div class="vis">${visTitleBar('Insert')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:20px;width:50px;height:16px;border:2px solid var(--accent);background:rgba(34,81,255,.08);animation:cellHighlight 4s ease-in-out infinite"></div>
            <div style="position:absolute;left:75px;top:22px;font-size:6px;color:var(--accent);font-weight:600;animation:pasteAppear 4s ease-in-out infinite">14:30:00</div>
            <div style="position:absolute;right:15px;top:60px;font-size:16px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#128339;</div>
            <span class="vis-label">INSERT TIME</span>
        </div></div>`,

    "Insert date": () => `<div class="vis">${visTitleBar('Insert')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:20px;width:50px;height:16px;border:2px solid var(--accent);background:rgba(34,81,255,.08);animation:cellHighlight 4s ease-in-out infinite"></div>
            <div style="position:absolute;left:74px;top:22px;font-size:5.5px;color:var(--accent);font-weight:600;animation:pasteAppear 4s ease-in-out infinite">2/9/2026</div>
            <div style="position:absolute;right:15px;top:60px;font-size:16px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#128197;</div>
            <span class="vis-label">INSERT DATE</span>
        </div></div>`,

    "Delete cell/range/row/column": () => `<div class="vis">${visTitleBar('Insert')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:36px;width:50px;height:16px;background:rgba(198,40,40,.1);border:2px solid #C62828;animation:removeItem 4s ease-in-out infinite"></div>
            <div style="position:absolute;left:77px;top:38px;font-size:5.5px;color:#C62828;animation:removeItem 4s ease-in-out infinite">Data</div>
            <div style="position:absolute;right:12px;top:70px;font-size:12px;color:#C62828;animation:visPulse 4s ease-in-out infinite">&#8722;</div>
            <span class="vis-label">DELETE</span>
        </div></div>`,

    "Insert table": () => `<div class="vis">${visTitleBar('Insert')}
        <div class="vis-canvas">
            <div style="position:absolute;left:30px;top:10px;animation:insertAppear 4s ease-in-out infinite">
                <div style="display:flex;border-radius:4px 4px 0 0;overflow:hidden">
                    <div style="width:40px;height:14px;background:var(--accent);border:1px solid var(--accent-dark);font-size:5px;color:white;display:flex;align-items:center;justify-content:center;font-weight:600">Name</div>
                    <div style="width:40px;height:14px;background:var(--accent);border:1px solid var(--accent-dark);font-size:5px;color:white;display:flex;align-items:center;justify-content:center;font-weight:600">Sales</div>
                    <div style="width:40px;height:14px;background:var(--accent);border:1px solid var(--accent-dark);font-size:5px;color:white;display:flex;align-items:center;justify-content:center;font-weight:600">Q1</div>
                </div>
                ${[0,1,2].map(r => `<div style="display:flex">
                    ${[0,1,2].map(c => `<div style="width:40px;height:12px;background:${r%2===0?'#E8EEFF':'white'};border:1px solid #D0D5DD;font-size:4.5px;display:flex;align-items:center;justify-content:center;color:#333">${['Alice','100','Q1','Bob','200','Q2','Carol','300','Q3'][r*3+c]}</div>`).join('')}
                </div>`).join('')}
            </div>
            <span class="vis-label">TABLE</span>
        </div></div>`,

    "Insert sheet": () => `<div class="vis">${visTitleBar('Insert')}
        <div class="vis-canvas">
            <div style="position:absolute;bottom:8px;left:10px;display:flex;gap:2px;align-items:flex-end">
                <div class="vis-tab active-tab">Sheet1</div>
                <div class="vis-tab">Sheet2</div>
                <div class="vis-tab" style="border-color:var(--accent);color:var(--accent);font-weight:700;animation:tabSlideIn 4s ease-in-out infinite">+ Sheet3</div>
            </div>
            <div style="position:absolute;left:50%;top:30px;transform:translateX(-50%);text-align:center">
                <div style="font-size:22px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#10133;</div>
                <div style="font-size:6px;color:#888;margin-top:2px">New sheet tab</div>
            </div>
            <span class="vis-label">NEW SHEET</span>
        </div></div>`,

    "Insert cell/range/row/column": () => `<div class="vis">${visTitleBar('Insert')}
        <div class="vis-canvas">
            ${miniGrid({cols: 3, rows: 3, cellW: 50})}
            <div style="position:absolute;left:72px;top:36px;width:50px;height:16px;background:rgba(34,81,255,.1);border:2px solid var(--accent);animation:insertAppear 4s ease-in-out infinite"></div>
            <div style="position:absolute;left:72px;top:52px;width:50px;height:16px;animation:slideDown 4s ease-in-out infinite">
                <div style="font-size:5.5px;color:#888;padding-left:5px;padding-top:2px">old data</div>
            </div>
            <div style="position:absolute;right:12px;top:70px;font-size:12px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#43;</div>
            <span class="vis-label">INSERT CELL</span>
        </div></div>`,

    // === SHEET ===
    "Delete sheet": () => `<div class="vis">${visTitleBar('Sheet')}
        <div class="vis-canvas">
            <div style="position:absolute;bottom:8px;left:10px;display:flex;gap:2px;align-items:flex-end">
                <div class="vis-tab active-tab">Sheet1</div>
                <div class="vis-tab" style="border-color:#C62828;color:#C62828;text-decoration:line-through;animation:removeItem 4s ease-in-out infinite">Sheet2</div>
                <div class="vis-tab">Sheet3</div>
            </div>
            <div style="position:absolute;left:50%;top:28px;transform:translateX(-50%);text-align:center">
                <div style="font-size:22px;color:#C62828;animation:visPulse 4s ease-in-out infinite">&#128465;</div>
                <div style="font-size:6px;color:#888;margin-top:2px">Sheet deleted</div>
            </div>
            <span class="vis-label">DELETE SHEET</span>
        </div></div>`,

    "Move or copy sheet": () => `<div class="vis">${visTitleBar('Sheet')}
        <div class="vis-canvas">
            <div style="position:absolute;bottom:8px;left:10px;display:flex;gap:2px;align-items:flex-end">
                <div class="vis-tab active-tab">Sheet1</div>
                <div class="vis-tab">Sheet2</div>
            </div>
            <div style="position:absolute;left:50%;top:30px;transform:translate(-50%,0);text-align:center">
                <div style="display:flex;gap:12px;align-items:center;margin-bottom:6px">
                    <div style="padding:3px 8px;background:#E8EEFF;border:1px solid var(--accent);border-radius:3px;font-size:6px;color:var(--accent)">Sheet1</div>
                    <div style="font-size:12px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">&#8594;</div>
                    <div style="padding:3px 8px;background:#E8EEFF;border:1px solid var(--accent);border-radius:3px;font-size:6px;color:var(--accent);animation:insertAppear 4s ease-in-out infinite">Sheet1 (2)</div>
                </div>
                <div style="font-size:5.5px;color:#888">Move or create a copy</div>
            </div>
            <span class="vis-label">MOVE/COPY</span>
        </div></div>`,

    "Rename sheet": () => `<div class="vis">${visTitleBar('Sheet')}
        <div class="vis-canvas">
            <div style="position:absolute;bottom:8px;left:10px;display:flex;gap:2px;align-items:flex-end">
                <div class="vis-tab" style="border:2px solid var(--accent);color:var(--accent);font-weight:600;background:#E8EEFF">
                    <span style="animation:visPulse 4s ease-in-out infinite">Sales Q1</span>
                </div>
                <div class="vis-tab">Sheet2</div>
            </div>
            <div style="position:absolute;left:50%;top:28px;transform:translateX(-50%);text-align:center">
                <div style="font-size:7px;color:#888;margin-bottom:6px">Sheet1 &#8594;</div>
                <div style="padding:4px 12px;background:#E8EEFF;border:2px solid var(--accent);border-radius:4px;font-size:8px;color:var(--accent);font-weight:700;animation:pasteAppear 4s ease-in-out infinite">Sales Q1</div>
            </div>
            <span class="vis-label">RENAME</span>
        </div></div>`,

    // === VIEW ===
    "Hide or show ribbon": () => `<div class="vis">${visTitleBar('View')}
        <div class="vis-canvas">
            <div style="position:absolute;left:0;right:0;top:0;animation:ribbonHide 4s ease-in-out infinite;overflow:hidden">
                <div class="vis-ribbon">
                    <div class="vis-ribbon-btn">&#9998;</div>
                    <div class="vis-ribbon-btn" style="font-weight:700">B</div>
                    <div class="vis-ribbon-btn" style="font-style:italic">I</div>
                    <div class="vis-ribbon-btn" style="text-decoration:underline">U</div>
                    <div style="width:1px;height:8px;background:#D0D5DD;margin:0 2px"></div>
                    <div class="vis-ribbon-btn">&#9783;</div>
                    <div class="vis-ribbon-btn">&#9776;</div>
                </div>
            </div>
            <div style="position:absolute;left:0;right:0;top:26px;bottom:0">
                ${miniGrid({cols: 4, rows: 3, cellW: 40, startY: 2})}
            </div>
            <div style="position:absolute;right:10px;bottom:6px;font-size:6px;color:var(--accent);font-weight:600;animation:visPulse 4s ease-in-out infinite">Toggle Ribbon</div>
            <span class="vis-label">RIBBON</span>
        </div></div>`,

    "Open zoom dialog": () => `<div class="vis">${visTitleBar('View')}
        <div class="vis-canvas">
            <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:110px;background:white;border:1.5px solid var(--accent);border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.1);animation:dialogPop 4s ease-in-out infinite">
                <div style="background:var(--accent);padding:3px 8px;border-radius:4px 4px 0 0;font-size:6px;color:white;font-weight:600">Zoom</div>
                <div style="padding:6px 8px;text-align:center">
                    <div style="font-size:18px;font-weight:800;color:var(--accent)">100%</div>
                    <div style="display:flex;justify-content:center;gap:8px;margin-top:4px">
                        <span style="font-size:5px;padding:2px 6px;background:#E8EEFF;border-radius:2px;color:var(--accent)">75%</span>
                        <span style="font-size:5px;padding:2px 6px;background:#E8EEFF;border-radius:2px;color:var(--accent);font-weight:700;border:1px solid var(--accent)">100%</span>
                        <span style="font-size:5px;padding:2px 6px;background:#E8EEFF;border-radius:2px;color:var(--accent)">200%</span>
                    </div>
                </div>
            </div>
            <span class="vis-label">ZOOM</span>
        </div></div>`,

    // === SYSTEM ===
    "Lock computer": () => `<div class="vis">${visTitleBar('System')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center;animation:insertAppear 4s ease-in-out infinite">
                <div style="font-size:28px">&#128274;</div>
                <div style="font-size:7px;color:#888;margin-top:4px">Computer locked</div>
            </div>
            <span class="vis-label">LOCK</span>
        </div></div>`,

    "Open Windows Explorer": () => `<div class="vis">${visTitleBar('System')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center;animation:insertAppear 4s ease-in-out infinite">
                <div style="font-size:28px">&#128193;</div>
                <div style="font-size:7px;color:#888;margin-top:4px">File Explorer opens</div>
            </div>
            <span class="vis-label">EXPLORER</span>
        </div></div>`,

    "Show desktop": () => `<div class="vis">${visTitleBar('System')}
        <div class="vis-canvas">
            <div style="position:absolute;left:20px;top:10px;width:60px;height:45px;background:white;border:1.5px solid #ddd;border-radius:4px;animation:removeItem 4s ease-in-out infinite">
                <div style="height:8px;background:#E2E8DB;border-radius:3px 3px 0 0"></div>
            </div>
            <div style="position:absolute;left:50px;top:20px;width:60px;height:45px;background:white;border:1.5px solid #ddd;border-radius:4px;animation:removeItem 4s ease-in-out infinite .1s">
                <div style="height:8px;background:#BBDEFB;border-radius:3px 3px 0 0"></div>
            </div>
            <div style="position:absolute;left:50%;top:70px;transform:translateX(-50%);font-size:6px;color:#888">All windows minimized</div>
            <span class="vis-label">DESKTOP</span>
        </div></div>`,

    "Open this list of keyboard shortcuts": () => `<div class="vis">${visTitleBar('System')}
        <div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
            <div style="text-align:center">
                <div style="display:flex;gap:4px;margin-bottom:6px;justify-content:center">
                    <div style="width:22px;height:18px;background:#E8EEFF;border:1.5px solid #B8CAFF;border-radius:3px;font-size:6px;display:flex;align-items:center;justify-content:center;font-weight:700;color:var(--accent);animation:insertAppear 4s ease-in-out infinite">&#9000;</div>
                    <div style="width:22px;height:18px;background:#E8EEFF;border:1.5px solid #B8CAFF;border-radius:3px;font-size:6px;display:flex;align-items:center;justify-content:center;font-weight:700;color:var(--accent);animation:insertAppear 4s ease-in-out infinite .2s">&#8997;</div>
                    <div style="width:22px;height:18px;background:#E8EEFF;border:1.5px solid #B8CAFF;border-radius:3px;font-size:6px;display:flex;align-items:center;justify-content:center;font-weight:700;color:var(--accent);animation:insertAppear 4s ease-in-out infinite .4s">Q</div>
                </div>
                <div style="font-size:6px;color:#888">Show all shortcuts</div>
            </div>
            <span class="vis-label">SHORTCUTS</span>
        </div></div>`
};

// ===== VISUAL HELPER GENERATORS =====

function makeSheetNavVisual(direction) {
    const isRight = direction === 'right';
    const tb = visTitleBar('Navigation');
    return `<div class="vis">${tb}
        <div class="vis-canvas">
            <div style="position:absolute;bottom:8px;left:10px;display:flex;gap:2px;align-items:flex-end">
                <div class="vis-tab ${!isRight ? '' : 'active-tab'}" style="${!isRight ? 'animation:tabHighlight 4s ease-in-out infinite' : ''}">Sheet1</div>
                <div class="vis-tab active-tab">Sheet2</div>
                <div class="vis-tab ${isRight ? '' : 'active-tab'}" style="${isRight ? 'animation:tabHighlight 4s ease-in-out infinite' : ''}">Sheet3</div>
            </div>
            <div style="position:absolute;left:50%;top:30px;transform:translateX(-50%);font-size:28px;color:var(--accent);animation:visPulse 4s ease-in-out infinite">${isRight ? '&#8594;' : '&#8592;'}</div>
            <div style="position:absolute;left:50%;top:58px;transform:translateX(-50%);font-size:6px;color:#888">Move to ${direction} sheet</div>
            <span class="vis-label">${direction.toUpperCase()} SHEET</span>
        </div></div>`;
}

function makeFontFormatVisual(type) {
    const tb = visTitleBar('Formatting');
    let before, after, label;
    if (type === 'bold') {
        before = `<span style="font-size:12px;color:#888">Normal Text</span>`;
        after = `<span style="font-size:12px;color:var(--accent);font-weight:800;animation:boldToggle 4s ease-in-out infinite">Bold Text</span>`;
        label = 'BOLD';
    } else if (type === 'italic') {
        before = `<span style="font-size:12px;color:#888">Normal Text</span>`;
        after = `<span style="font-size:12px;color:var(--accent);font-style:italic;animation:italicToggle 4s ease-in-out infinite">Italic Text</span>`;
        label = 'ITALIC';
    } else {
        before = `<span style="font-size:12px;color:#888">Normal Text</span>`;
        after = `<span style="font-size:12px;color:var(--accent);text-decoration:underline;animation:underlineToggle 4s ease-in-out infinite">Underline Text</span>`;
        label = 'UNDERLINE';
    }
    return `<div class="vis">${tb}<div class="vis-canvas" style="display:flex;align-items:center;justify-content:center;gap:10px">
        <div style="text-align:center">
            <div style="margin-bottom:6px">${before}</div>
            <div style="font-size:10px;color:var(--accent)">&#8595;</div>
            <div style="margin-top:4px">${after}</div>
        </div>
        <span class="vis-label">${label}</span>
    </div></div>`;
}

// Render a visual for a given shortcut title
function renderVisual(title) {
    const gen = VISUALS[title];
    if (gen) return gen();
    const shortcut = SHORTCUTS.find(s => s.title === title);
    const cat = shortcut ? shortcut.category : 'Excel';
    return `<div class="vis">${visTitleBar(cat)}<div class="vis-canvas" style="display:flex;align-items:center;justify-content:center">
        <div style="text-align:center;color:#888"><div style="font-size:24px;margin-bottom:4px">&#9881;</div><div style="font-size:7px">${title}</div></div>
        <span class="vis-label">${cat.toUpperCase()}</span>
    </div></div>`;
}

// ===== STATE =====
let state = {
    progress: {},
    streak: 0,
    bestSpeed: 0,
    currentCard: 0,
    cardDeck: [],
    practiceIndex: 0,
    practiceQueue: [],
    practiceCorrect: 0,
    practiceWrong: 0,
    practiceAnswered: false,
    quizQuestions: [],
    quizIndex: 0,
    quizScore: 0,
    quizAnswered: false,
    speedScore: 0,
    speedTimer: null,
    speedTimeLeft: 60,
    activeCategory: null,
    mustHaveOnly: false
};

// ===== INIT =====
function init() {
    loadProgress();
    addSVGDefs();
    setupNav();
    setupDashboard();
    setupLearn();
    setupPractice();
    setupQuiz();
    setupSpeed();
    updateHeaderStats();
    document.getElementById('totalShortcuts').textContent = SHORTCUTS.length;
}

function addSVGDefs() {
    const svg = document.querySelector('.progress-ring');
    if (svg) {
        const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
        const gradient = document.createElementNS('http://www.w3.org/2000/svg', 'linearGradient');
        gradient.id = 'progressGradient';
        gradient.setAttribute('x1', '0%'); gradient.setAttribute('y1', '0%');
        gradient.setAttribute('x2', '100%'); gradient.setAttribute('y2', '100%');
        const stop1 = document.createElementNS('http://www.w3.org/2000/svg', 'stop');
        stop1.setAttribute('offset', '0%'); stop1.setAttribute('stop-color', '#2251FF');
        const stop2 = document.createElementNS('http://www.w3.org/2000/svg', 'stop');
        stop2.setAttribute('offset', '100%'); stop2.setAttribute('stop-color', '#4A73FF');
        gradient.appendChild(stop1); gradient.appendChild(stop2);
        defs.appendChild(gradient); svg.insertBefore(defs, svg.firstChild);
    }
}

// ===== PERSISTENCE =====
function loadProgress() {
    try {
        const saved = localStorage.getItem('xlsx-shortcut-master');
        if (saved) {
            const data = JSON.parse(saved);
            state.progress = data.progress || {};
            state.streak = data.streak || 0;
            state.bestSpeed = data.bestSpeed || 0;
        }
    } catch (e) { /* ignore */ }
    SHORTCUTS.forEach(s => {
        if (!state.progress[s.title]) {
            state.progress[s.title] = { status: 'new', correctStreak: 0, lastSeen: 0 };
        }
    });
    const validTitles = new Set(SHORTCUTS.map(s => s.title));
    Object.keys(state.progress).forEach(key => {
        if (!validTitles.has(key)) delete state.progress[key];
    });
}

function saveProgress() {
    try {
        localStorage.setItem('xlsx-shortcut-master', JSON.stringify({
            progress: state.progress, streak: state.streak, bestSpeed: state.bestSpeed
        }));
    } catch (e) { /* ignore */ }
}

function resetAllProgress() {
    if (confirm('Reset all progress? This cannot be undone.')) {
        state.progress = {};
        state.streak = 0;
        state.bestSpeed = 0;
        SHORTCUTS.forEach(s => {
            state.progress[s.title] = { status: 'new', correctStreak: 0, lastSeen: 0 };
        });
        saveProgress();
        setupDashboard();
        updateHeaderStats();
    }
}

// ===== HELPERS =====
function shuffle(arr) {
    const a = [...arr];
    for (let i = a.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
}

function getFilteredShortcuts() {
    let shortcuts = [...SHORTCUTS];
    if (state.activeCategory) shortcuts = shortcuts.filter(s => s.category === state.activeCategory);
    if (state.mustHaveOnly) shortcuts = shortcuts.filter(s => s.mustHave);
    return shortcuts;
}

function getMasteredCount() { return Object.values(state.progress).filter(p => p.status === 'mastered').length; }
function getLearningCount() { return Object.values(state.progress).filter(p => p.status === 'learning').length; }
function getNewCount() { return Object.values(state.progress).filter(p => p.status === 'new').length; }

function updateHeaderStats() {
    document.getElementById('masteredCount').textContent = getMasteredCount();
    document.getElementById('streakCount').textContent = state.streak;
}

function markCorrect(title) {
    const p = state.progress[title];
    p.correctStreak++;
    p.lastSeen = Date.now();
    if (p.correctStreak >= 3) p.status = 'mastered';
    else if (p.status === 'new') p.status = 'learning';
    state.streak++;
    saveProgress(); updateHeaderStats();
}

function markWrong(title) {
    const p = state.progress[title];
    p.correctStreak = 0;
    p.lastSeen = Date.now();
    if (p.status === 'new') p.status = 'learning';
    state.streak = 0;
    saveProgress(); updateHeaderStats();
}

// ===== NAVIGATION =====
function setupNav() {
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => switchView(btn.dataset.mode));
    });
    document.querySelectorAll('.mode-card[data-launch]').forEach(btn => {
        btn.addEventListener('click', () => switchView(btn.dataset.launch));
    });
}

function switchView(mode) {
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    const activeNav = document.querySelector(`.nav-btn[data-mode="${mode}"]`);
    if (activeNav) activeNav.classList.add('active');
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    const view = document.getElementById(mode);
    if (view) view.classList.add('active');
    if (mode === 'learn') initLearnDeck();
    if (mode === 'practice') initPractice();
    if (mode === 'quiz') initQuiz();
    if (mode === 'speed') initSpeed();
    if (mode === 'dashboard') setupDashboard();
}

// ===== DASHBOARD =====
function setupDashboard() {
    updateProgressRing();
    renderMustHaveToggle();
    renderCategoryChips();
    renderShortcutTable();
    document.getElementById('resetProgress').onclick = resetAllProgress;
}

function renderMustHaveToggle() {
    const container = document.getElementById('mustHaveToggle');
    if (!container) return;
    container.innerHTML = `
        <label class="toggle-switch-label">
            <span class="toggle-text">Must-Haves Only</span>
            <div class="toggle-switch">
                <input type="checkbox" id="mustHaveCheck" ${state.mustHaveOnly ? 'checked' : ''}>
                <span class="toggle-slider"></span>
            </div>
            <span class="toggle-hint">Applies across all modes</span>
        </label>
    `;
    document.getElementById('mustHaveCheck').addEventListener('change', (e) => {
        state.mustHaveOnly = e.target.checked;
        renderCategoryChips();
        renderShortcutTable();
        // Re-init active mode
        const activeView = document.querySelector('.view.active');
        if (activeView) {
            const mode = activeView.id;
            if (mode === 'learn') initLearnDeck();
            if (mode === 'practice') initPractice();
            if (mode === 'quiz') initQuiz();
            if (mode === 'speed') initSpeed();
        }
    });
}

function updateProgressRing() {
    const mastered = getMasteredCount();
    const total = SHORTCUTS.length;
    const percent = Math.round((mastered / total) * 100);
    document.getElementById('progressPercent').textContent = percent + '%';
    document.getElementById('newCount').textContent = getNewCount();
    document.getElementById('learningCount').textContent = getLearningCount();
    document.getElementById('masteredCountDash').textContent = mastered;
    const ring = document.getElementById('progressRingFill');
    const circumference = 2 * Math.PI * 70;
    ring.style.strokeDashoffset = circumference - (mastered / total) * circumference;
}

function renderCategoryChips() {
    const container = document.getElementById('categoryChips');
    container.innerHTML = '';
    const pool = state.mustHaveOnly ? SHORTCUTS.filter(s => s.mustHave) : SHORTCUTS;
    const allChip = document.createElement('button');
    allChip.className = 'category-chip' + (!state.activeCategory ? ' active' : '');
    allChip.textContent = `All (${pool.length})`;
    allChip.addEventListener('click', () => { state.activeCategory = null; renderCategoryChips(); renderShortcutTable(); });
    container.appendChild(allChip);
    CATEGORIES.forEach(cat => {
        const count = pool.filter(s => s.category === cat).length;
        if (state.mustHaveOnly && count === 0) return;
        const chip = document.createElement('button');
        chip.className = 'category-chip' + (state.activeCategory === cat ? ' active' : '');
        chip.textContent = `${cat} (${count})`;
        chip.addEventListener('click', () => { state.activeCategory = cat; renderCategoryChips(); renderShortcutTable(); });
        container.appendChild(chip);
    });
}

function renderShortcutTable(filter = '') {
    const tbody = document.getElementById('shortcutTableBody');
    tbody.innerHTML = '';
    let shortcuts = getFilteredShortcuts();
    if (filter) {
        const lf = filter.toLowerCase();
        shortcuts = shortcuts.filter(s => s.title.toLowerCase().includes(lf) || s.key.toLowerCase().includes(lf) || s.category.toLowerCase().includes(lf));
    }
    shortcuts.forEach(s => {
        const p = state.progress[s.title];
        const tr = document.createElement('tr');
        const star = s.mustHave ? '<span class="must-have-badge" title="Must-have shortcut">&#9733;</span>' : '';
        tr.innerHTML = `<td>${star}${s.title}</td><td><span class="shortcut-key">${s.key}</span></td><td><span class="category-badge">${s.category}</span></td><td><span class="status-badge ${p.status}">${p.status}</span></td>`;
        tbody.appendChild(tr);
    });
    const searchInput = document.getElementById('searchInput');
    searchInput.removeEventListener('input', handleSearch);
    searchInput.addEventListener('input', handleSearch);
}

function handleSearch(e) { renderShortcutTable(e.target.value); }

// ===== LEARN (FLASHCARDS) =====
function setupLearn() {
    document.getElementById('flashcard').addEventListener('click', flipCard);
    document.querySelectorAll('.recall-btn').forEach(btn => {
        btn.addEventListener('click', () => handleRecall(btn.dataset.recall));
    });
    document.getElementById('prevCard').addEventListener('click', prevCard);
    document.getElementById('skipCard').addEventListener('click', skipCard);
}

function initLearnDeck() {
    const filtered = getFilteredShortcuts();
    const newAndLearning = filtered.filter(s => state.progress[s.title].status !== 'mastered');
    const mastered = filtered.filter(s => state.progress[s.title].status === 'mastered');
    state.cardDeck = [...shuffle(newAndLearning), ...shuffle(mastered)];
    if (state.cardDeck.length === 0) state.cardDeck = shuffle(filtered);
    state.currentCard = 0;
    showCard();
}

function showCard() {
    const deck = state.cardDeck;
    if (deck.length === 0) return;
    const idx = state.currentCard % deck.length;
    const card = deck[idx];
    document.getElementById('cardFrontText').textContent = card.title;
    document.getElementById('cardBackText').textContent = card.key;
    document.getElementById('cardProgress').textContent = `${idx + 1} / ${deck.length}`;
    document.getElementById('cardProgressFill').style.width = ((idx + 1) / deck.length * 100) + '%';
    const visContainer = document.getElementById('cardVisual');
    if (visContainer) visContainer.innerHTML = renderVisual(card.title);
    document.getElementById('flashcardInner').classList.remove('flipped');
    document.getElementById('recallButtons').style.display = 'none';
}

function flipCard() {
    const inner = document.getElementById('flashcardInner');
    inner.classList.toggle('flipped');
    if (inner.classList.contains('flipped')) {
        document.getElementById('recallButtons').style.display = 'flex';
    }
}

function handleRecall(level) {
    const deck = state.cardDeck;
    const idx = state.currentCard % deck.length;
    const card = deck[idx];
    if (level === 'easy') {
        markCorrect(card.title);
    } else if (level === 'hard') {
        markWrong(card.title);
        deck.push(card);
    } else {
        const p = state.progress[card.title];
        if (p.status === 'new') p.status = 'learning';
        p.lastSeen = Date.now();
        saveProgress();
    }
    state.currentCard++;
    showCard();
}

function prevCard() { if (state.currentCard > 0) state.currentCard--; showCard(); }
function skipCard() { state.currentCard++; showCard(); }

// ===== PRACTICE (TYPE SHORTCUT) =====
function setupPractice() {
    document.getElementById('showAnswer').addEventListener('click', showPracticeAnswer);
    document.getElementById('nextPractice').addEventListener('click', nextPracticeQuestion);
    document.getElementById('submitTyped').addEventListener('click', handleTypedAnswer);
    document.getElementById('typeFallback').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); handleTypedAnswer(); }
        e.stopPropagation();
    });
    document.getElementById('typeFallback').addEventListener('keyup', (e) => e.stopPropagation());
}

function initPractice() {
    state.practiceQueue = shuffle(getFilteredShortcuts());
    state.practiceIndex = 0;
    state.practiceCorrect = 0;
    state.practiceWrong = 0;
    state.practiceAnswered = false;
    document.getElementById('practiceCorrect').textContent = '0';
    document.getElementById('practiceWrong').textContent = '0';
    showPracticeQuestion();
    startKeyListener();
}

function showPracticeQuestion() {
    const q = state.practiceQueue;
    if (q.length === 0) return;
    const idx = state.practiceIndex % q.length;
    const shortcut = q[idx];
    document.getElementById('practiceAction').textContent = shortcut.title;
    const visContainer = document.getElementById('practiceVisual');
    if (visContainer) visContainer.innerHTML = renderVisual(shortcut.title);
    document.getElementById('keyDisplay').innerHTML = '<span class="key-placeholder">Press the shortcut keys...</span>';
    document.getElementById('keyDisplay').className = 'key-display listening';
    document.getElementById('practiceFeedback').textContent = '';
    document.getElementById('practiceFeedback').className = 'practice-feedback';
    const fallbackInput = document.getElementById('typeFallback');
    if (fallbackInput) fallbackInput.value = '';
    state.practiceAnswered = false;
}

function startKeyListener() {
    document.removeEventListener('keydown', handlePracticeKey);
    document.addEventListener('keydown', handlePracticeKey);
}

function handlePracticeKey(e) {
    if (e.target.id === 'typeFallback') return;
    const practiceView = document.getElementById('practice');
    if (!practiceView.classList.contains('active')) return;
    if (state.practiceAnswered) return;
    e.preventDefault();
    const parts = [];
    if (e.ctrlKey || e.metaKey) parts.push('Ctrl');
    if (e.altKey) parts.push('Alt');
    if (e.shiftKey) parts.push('Shift');
    const keyMap = {
        'ArrowLeft':'Left Arrow', 'ArrowRight':'Right Arrow', 'ArrowUp':'Up Arrow', 'ArrowDown':'Down Arrow',
        'BracketRight':']', 'BracketLeft':'[', ' ':'Space', 'Space':'Space',
        'PageDown':'Page Down', 'PageUp':'Page Up', 'Home':'Home', 'End':'End',
        'Delete':'Del', 'Backspace':'Backspace', 'Enter':'Enter', 'Escape':'Escape',
        'F1':'F1','F2':'F2','F3':'F3','F4':'F4','F5':'F5','F6':'F6','F7':'F7',
        'F8':'F8','F9':'F9','F10':'F10','F11':'F11','F12':'F12',
        'Semicolon':';', 'Comma':',', 'Period':'.', 'Slash':'/', 'Backquote':'`',
        'Equal':'=', 'Minus':'-'
    };
    let key = e.key;
    if (keyMap[key]) key = keyMap[key];
    else if (keyMap[e.code]) key = keyMap[e.code];
    else if (key.length === 1) key = key.toUpperCase();
    if (['Control', 'Alt', 'Shift', 'Meta'].includes(key)) {
        document.getElementById('keyDisplay').innerHTML = `<span style="color:var(--accent)">${parts.join('+') + '+...'}</span>`;
        return;
    }
    parts.push(key);
    const pressed = parts.join('+');
    const display = document.getElementById('keyDisplay');
    display.innerHTML = `<span>${pressed}</span>`;
    const q = state.practiceQueue;
    const idx = state.practiceIndex % q.length;
    const correct = q[idx].key;
    if (normalizeKey(pressed) === normalizeKey(correct)) {
        display.className = 'key-display correct';
        document.getElementById('practiceFeedback').textContent = 'Correct!';
        document.getElementById('practiceFeedback').className = 'practice-feedback correct';
        state.practiceCorrect++;
        document.getElementById('practiceCorrect').textContent = state.practiceCorrect;
        markCorrect(q[idx].title);
        state.practiceAnswered = true;
    } else {
        display.className = 'key-display wrong';
        document.getElementById('practiceFeedback').innerHTML = `Wrong — the answer is <strong>${correct}</strong>`;
        document.getElementById('practiceFeedback').className = 'practice-feedback wrong';
        state.practiceWrong++;
        document.getElementById('practiceWrong').textContent = state.practiceWrong;
        markWrong(q[idx].title);
        state.practiceAnswered = true;
    }
}

function normalizeKey(k) {
    return k.toLowerCase().replace(/\s+/g, '')
        .replace('arrowleft', 'leftarrow').replace('arrowright', 'rightarrow')
        .replace('arrowup', 'uparrow').replace('arrowdown', 'downarrow')
        .replace('pagedown', 'pagedown').replace('pageup', 'pageup');
}

function showPracticeAnswer() {
    const q = state.practiceQueue;
    const idx = state.practiceIndex % q.length;
    const correct = q[idx].key;
    document.getElementById('keyDisplay').innerHTML = `<span style="color:var(--accent)">${correct}</span>`;
    document.getElementById('keyDisplay').className = 'key-display';
    document.getElementById('practiceFeedback').textContent = 'Answer revealed';
    document.getElementById('practiceFeedback').className = 'practice-feedback';
    state.practiceAnswered = true;
}

function handleTypedAnswer() {
    if (state.practiceAnswered) return;
    const input = document.getElementById('typeFallback');
    const typed = input.value.trim();
    if (!typed) return;
    const q = state.practiceQueue;
    const idx = state.practiceIndex % q.length;
    const shortcut = q[idx];
    const correct = shortcut.key;
    const display = document.getElementById('keyDisplay');
    display.innerHTML = `<span>${typed}</span>`;
    if (normalizeKey(typed) === normalizeKey(correct)) {
        display.className = 'key-display correct';
        document.getElementById('practiceFeedback').textContent = 'Correct!';
        document.getElementById('practiceFeedback').className = 'practice-feedback correct';
        state.practiceCorrect++;
        document.getElementById('practiceCorrect').textContent = state.practiceCorrect;
        markCorrect(shortcut.title);
    } else {
        display.className = 'key-display wrong';
        document.getElementById('practiceFeedback').innerHTML = `Wrong — the answer is <strong>${correct}</strong>`;
        document.getElementById('practiceFeedback').className = 'practice-feedback wrong';
        state.practiceWrong++;
        document.getElementById('practiceWrong').textContent = state.practiceWrong;
        markWrong(shortcut.title);
    }
    state.practiceAnswered = true;
}

function nextPracticeQuestion() { state.practiceIndex++; showPracticeQuestion(); }

// ===== QUIZ =====
function setupQuiz() {
    document.getElementById('quizNext').addEventListener('click', nextQuizQuestion);
    document.getElementById('retakeQuiz').addEventListener('click', initQuiz);
}

function initQuiz() {
    const pool = shuffle(getFilteredShortcuts());
    state.quizQuestions = pool.slice(0, Math.min(10, pool.length));
    state.quizIndex = 0;
    state.quizScore = 0;
    state.quizAnswered = false;
    document.getElementById('quizResults').style.display = 'none';
    document.querySelector('.quiz-area').style.display = 'block';
    showQuizQuestion();
}

function showQuizQuestion() {
    if (state.quizIndex >= state.quizQuestions.length) { showQuizResults(); return; }
    const current = state.quizQuestions[state.quizIndex];
    document.getElementById('quizAction').textContent = current.title;
    const visContainer = document.getElementById('quizVisual');
    if (visContainer) visContainer.innerHTML = renderVisual(current.title);
    document.getElementById('quizProgressText').textContent = `Question ${state.quizIndex + 1} / ${state.quizQuestions.length}`;
    document.getElementById('quizScoreText').textContent = `Score: ${state.quizScore}`;
    document.getElementById('quizProgressFill').style.width = ((state.quizIndex + 1) / state.quizQuestions.length * 100) + '%';
    document.getElementById('quizNext').style.display = 'none';
    state.quizAnswered = false;
    const correctKey = current.key;
    const wrongPool = SHORTCUTS.filter(s => s.key !== correctKey);
    const wrongOptions = shuffle(wrongPool).slice(0, 3).map(s => s.key);
    const options = shuffle([correctKey, ...wrongOptions]);
    const container = document.getElementById('quizOptions');
    container.innerHTML = '';
    options.forEach(opt => {
        const btn = document.createElement('button');
        btn.className = 'quiz-option';
        btn.textContent = opt;
        btn.addEventListener('click', () => handleQuizAnswer(btn, opt, correctKey, current.title));
        container.appendChild(btn);
    });
}

function handleQuizAnswer(btn, selected, correct, title) {
    if (state.quizAnswered) return;
    state.quizAnswered = true;
    document.querySelectorAll('.quiz-option').forEach(b => b.classList.add('disabled'));
    if (selected === correct) {
        btn.classList.add('selected-correct');
        state.quizScore++;
        markCorrect(title);
    } else {
        btn.classList.add('selected-wrong');
        markWrong(title);
        document.querySelectorAll('.quiz-option').forEach(b => { if (b.textContent === correct) b.classList.add('reveal-correct'); });
    }
    document.getElementById('quizScoreText').textContent = `Score: ${state.quizScore}`;
    document.getElementById('quizNext').style.display = 'block';
}

function nextQuizQuestion() { state.quizIndex++; showQuizQuestion(); }

function showQuizResults() {
    document.querySelector('.quiz-area').style.display = 'none';
    document.getElementById('quizResults').style.display = 'flex';
    const total = state.quizQuestions.length;
    const correct = state.quizScore;
    const wrong = total - correct;
    const accuracy = Math.round((correct / total) * 100);
    document.getElementById('resultCorrect').textContent = correct;
    document.getElementById('resultWrong').textContent = wrong;
    document.getElementById('resultAccuracy').textContent = accuracy + '%';
    if (accuracy >= 80) { document.getElementById('resultsTitle').textContent = 'Excellent!'; document.getElementById('resultsSubtitle').textContent = `You scored ${correct} out of ${total}!`; }
    else if (accuracy >= 60) { document.getElementById('resultsTitle').textContent = 'Good Effort!'; document.getElementById('resultsSubtitle').textContent = `You scored ${correct} out of ${total}. Keep practicing!`; }
    else { document.getElementById('resultsTitle').textContent = 'Keep Learning!'; document.getElementById('resultsSubtitle').textContent = `You scored ${correct} out of ${total}. Review the flashcards and try again.`; }
    const missedContainer = document.getElementById('resultsMissed');
    missedContainer.innerHTML = '';
    if (wrong > 0) {
        const missed = state.quizQuestions.filter(q => state.progress[q.title].status !== 'mastered');
        if (missed.length > 0) {
            missedContainer.innerHTML = '<h4>Review These:</h4>';
            missed.slice(0, wrong).forEach(m => {
                const div = document.createElement('div');
                div.className = 'missed-item';
                div.innerHTML = `<span>${m.title}</span><span class="shortcut-key">${m.key}</span>`;
                missedContainer.appendChild(div);
            });
        }
    }
}

// ===== SPEED RUN =====
function setupSpeed() {
    document.getElementById('startSpeed').addEventListener('click', startSpeedRun);
    document.getElementById('retrySpeed').addEventListener('click', initSpeed);
}

function initSpeed() {
    document.getElementById('speedPre').style.display = 'block';
    document.getElementById('speedGame').style.display = 'none';
    document.getElementById('speedResults').style.display = 'none';
    document.getElementById('bestScore').textContent = state.bestSpeed;
    if (state.speedTimer) clearInterval(state.speedTimer);
}

function startSpeedRun() {
    state.speedScore = 0;
    state.speedTimeLeft = 60;
    document.getElementById('speedPre').style.display = 'none';
    document.getElementById('speedGame').style.display = 'block';
    document.getElementById('speedResults').style.display = 'none';
    document.getElementById('speedScore').textContent = '0';
    document.getElementById('speedTimer').textContent = '60';
    document.getElementById('speedTimer').className = 'speed-timer';
    showSpeedQuestion();
    startSpeedTimer();
}

function startSpeedTimer() {
    state.speedTimer = setInterval(() => {
        state.speedTimeLeft--;
        const timerEl = document.getElementById('speedTimer');
        timerEl.textContent = state.speedTimeLeft;
        if (state.speedTimeLeft <= 10) timerEl.className = 'speed-timer danger';
        else if (state.speedTimeLeft <= 20) timerEl.className = 'speed-timer warning';
        if (state.speedTimeLeft <= 0) { clearInterval(state.speedTimer); endSpeedRun(); }
    }, 1000);
}

function showSpeedQuestion() {
    const pool = getFilteredShortcuts();
    const current = pool[Math.floor(Math.random() * pool.length)];
    document.getElementById('speedAction').textContent = current.title;
    const visContainer = document.getElementById('speedVisual');
    if (visContainer) visContainer.innerHTML = renderVisual(current.title);
    const correctKey = current.key;
    const wrongPool = SHORTCUTS.filter(s => s.key !== correctKey);
    const wrongOptions = shuffle(wrongPool).slice(0, 3).map(s => s.key);
    const options = shuffle([correctKey, ...wrongOptions]);
    const container = document.getElementById('speedOptions');
    container.innerHTML = '';
    options.forEach(opt => {
        const btn = document.createElement('button');
        btn.className = 'speed-option';
        btn.textContent = opt;
        btn.addEventListener('click', () => handleSpeedAnswer(btn, opt, correctKey, current.title));
        container.appendChild(btn);
    });
}

function handleSpeedAnswer(btn, selected, correct, title) {
    if (selected === correct) {
        btn.classList.add('flash-correct');
        state.speedScore++;
        document.getElementById('speedScore').textContent = state.speedScore;
        markCorrect(title);
        setTimeout(showSpeedQuestion, 200);
    } else {
        btn.classList.add('flash-wrong');
        markWrong(title);
        document.querySelectorAll('.speed-option').forEach(b => { if (b.textContent === correct) b.classList.add('flash-correct'); });
        setTimeout(showSpeedQuestion, 600);
    }
}

function endSpeedRun() {
    document.getElementById('speedGame').style.display = 'none';
    document.getElementById('speedResults').style.display = 'block';
    document.getElementById('speedFinalScore').textContent = state.speedScore;
    const newBestEl = document.getElementById('speedNewBest');
    if (state.speedScore > state.bestSpeed) {
        state.bestSpeed = state.speedScore;
        saveProgress();
        newBestEl.style.display = 'block';
    } else { newBestEl.style.display = 'none'; }
}

// ===== BOOTSTRAP =====
document.addEventListener('DOMContentLoaded', init);
