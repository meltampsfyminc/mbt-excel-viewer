const vscode = require('vscode');
const cp = require('child_process');

/**
 * Called when the extension is activated
 */
function activate(context) {
  // Register the command
  let disposable = vscode.commands.registerCommand('excelViewer.openFile', async (uri) => {
    let filePath;

    // If the command was triggered from right-click → "Open Excel File"
    if (uri && uri.fsPath) {
      filePath = uri.fsPath;
    } else {
      // Otherwise prompt the user to pick a file
      const files = await vscode.window.showOpenDialog({
        filters: { 'Excel Files': ['xls', 'xlsx'] },
        canSelectMany: false
      });
      if (files && files.length > 0) {
        filePath = files[0].fsPath;
      }
    }

    if (!filePath) {
      vscode.window.showErrorMessage("No Excel file selected.");
      return;
    }

    runPython(filePath);
  });

  context.subscriptions.push(disposable);
}

/**
 * Run the Python backend to parse the Excel file
 */
function runPython(filePath) {
  // Call your Python script
  cp.exec(`python read_excel.py "${filePath}"`, (err, stdout, stderr) => {
    if (err) {
      vscode.window.showErrorMessage(`Python error: ${stderr}`);
      return;
    }

    // Show the parsed output in VS Code
    const panel = vscode.window.createWebviewPanel(
      'excelViewer',
      `Excel Viewer: ${filePath}`,
      vscode.ViewColumn.One,
      {}
    );

    panel.webview.html = `
      <html>
        <head>
          <link rel="stylesheet" type="text/css" href="${panel.webview.asWebviewUri(vscode.Uri.file(__dirname + '/media/dataTables.min.css'))}">
        </head>
        <body>
          <div id="table">${stdout}</div>
          <script src="${panel.webview.asWebviewUri(vscode.Uri.file(__dirname + '/media/dataTables.min.js'))}"></script>
        </body>
      </html>
    `;
  });
}

exports.activate = activate;
