const vscode = require('vscode');
const { spawn } = require('child_process');
const path = require('path');

function activate(context) {
    const disposable = vscode.commands.registerCommand('excelViewer.openFile', async () => {
        const fileUri = await vscode.window.showOpenDialog({
            canSelectMany: false,
            filters: { 'Excel Files': ['xlsx', 'xls'] }
        });
        if (!fileUri) return;

        const filePath = fileUri[0].fsPath;

        // Ask Python for sheet names first
        const sheets = await runPython(context, ['read_excel.py', '--list-sheets', filePath]);
        const sheetList = JSON.parse(sheets || '[]');
        const pick = await vscode.window.showQuickPick(sheetList.map(s => s.name), {
            placeHolder: 'Select a sheet to view'
        });
        if (!pick) return;

        const sheet = sheetList.find(s => s.name === pick);
        const panel = vscode.window.createWebviewPanel(
            'excelViewer',
            `Excel Viewer: ${path.basename(filePath)} — ${sheet.name}`,
            vscode.ViewColumn.One,
            {
                enableScripts: true,
                retainContextWhenHidden: true
            }
        );

        const nonce = getNonce();
        const mediaRoot = vscode.Uri.joinPath(context.extensionUri, 'media');
        const cssUri = panel.webview.asWebviewUri(vscode.Uri.joinPath(mediaRoot, 'datatables.min.css'));
        const jsUri = panel.webview.asWebviewUri(vscode.Uri.joinPath(mediaRoot, 'datatables.min.js'));

        panel.webview.html = getWebviewBase({ cssUri, jsUri, nonce, columns: sheet.columns });

        // Initial page load
        let page = 1;
        const pageSize = 50;

        const loadPage = async (pageNumber) => {
            const out = await runPython(context, [
                'read_excel.py',
                '--read',
                filePath,
                '--sheet', sheet.name,
                '--page', String(pageNumber),
                '--size', String(pageSize)
            ]);
            panel.webview.postMessage({ type: 'page', page: pageNumber, rows: JSON.parse(out || '[]') });
        };

        // Handle messages from webview (paging, export, etc.)
        panel.webview.onDidReceiveMessage(async (msg) => {
            if (msg.type === 'next') {
                page += 1;
                await loadPage(page);
            } else if (msg.type === 'prev' && page > 1) {
                page -= 1;
                await loadPage(page);
            } else if (msg.type === 'reload') {
                await loadPage(page);
            }
        });

        await loadPage(page);
    });

    context.subscriptions.push(disposable);
}

function getWebviewBase({ cssUri, jsUri, nonce, columns }) {
    const headers = columns.map(c => `<th>${escapeHtml(c)}</th>`).join('');
    return `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy"
      content="default-src 'none'; img-src data:; style-src ${cssUri} 'unsafe-inline'; script-src 'nonce-${nonce}';">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="${cssUri}">
<title>Excel Viewer</title>
<style>
  body { font-family: system-ui, sans-serif; margin: 12px; }
  #controls { margin-bottom: 10px; display: flex; gap: 8px; align-items: center; }
  table { width: 100%; }
</style>
</head>
<body>
  <div id="controls">
    <button id="prev">Prev</button>
    <button id="next">Next</button>
    <button id="reload">Reload</button>
    <span id="pageInfo"></span>
  </div>
  <table id="excelTable" class="display compact">
    <thead><tr>${headers}</tr></thead>
    <tbody></tbody>
  </table>

<script nonce="${nonce}">
  const vscode = acquireVsCodeApi();
  const tableEl = document.getElementById('excelTable');
  const pageInfo = document.getElementById('pageInfo');
  let dataTable;

  window.addEventListener('message', event => {
    const { type, rows, page } = event.data;
    if (type === 'page') {
      pageInfo.textContent = 'Page ' + page;
      // Rebuild tbody
      const tbody = tableEl.querySelector('tbody');
      tbody.innerHTML = '';
      for (const row of rows) {
        const tr = document.createElement('tr');
        for (const cell of row) {
          const td = document.createElement('td');
          td.textContent = String(cell ?? '');
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
      }
      if (!dataTable) {
        // DataTables loaded next
      } else {
        dataTable.clear();
        dataTable.rows.add(rows);
        dataTable.draw();
      }
    }
  });

  document.getElementById('prev').addEventListener('click', () => vscode.postMessage({ type: 'prev' }));
  document.getElementById('next').addEventListener('click', () => vscode.postMessage({ type: 'next' }));
  document.getElementById('reload').addEventListener('click', () => vscode.postMessage({ type: 'reload' }));
</script>
<script nonce="${nonce}" src="${jsUri}"></script>
<script nonce="${nonce}">
  // Initialize DataTables after script is loaded
  document.addEventListener('DOMContentLoaded', () => {
    dataTable = new DataTable('#excelTable', {
      paging: false, // We handle paging server-side
      searching: true,
      ordering: true,
      info: false
    });
  });
</script>
</body>
</html>
`;
}

function runPython(context, args) {
    return new Promise((resolve, reject) => {
        const py = spawn(getPython(), [context.asAbsolutePath(args[0]), ...args.slice(1)], {
            env: { ...process.env, PYTHONUNBUFFERED: '1' }
        });
        let out = '';
        let err = '';
        py.stdout.on('data', (d) => (out += d.toString()));
        py.stderr.on('data', (d) => (err += d.toString()));
        py.on('close', (code) => {
            if (code === 0) resolve(out.trim());
            else reject(new Error(err || `Python exited with ${code}`));
        });
    });
}

function getPython() {
    // Prefer 'python3' if available, fall back to 'python'
    return process.platform === 'win32' ? 'python' : 'python3';
}

function getNonce() {
    const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let text = '';
    for (let i = 0; i < 32; i++) text += possible.charAt(Math.floor(Math.random() * possible.length));
    return text;
}

function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, (c) => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

function deactivate() {}

module.exports = { activate, deactivate };
