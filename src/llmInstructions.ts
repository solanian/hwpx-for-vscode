import * as vscode from 'vscode';

export type LlmTarget = 'claude' | 'codex' | 'antigravity' | 'cursor' | 'kiro';

export function getTargetFileName(target: LlmTarget): string {
    switch (target) {
        case 'claude': return 'CLAUDE.md';
        case 'codex': return 'AGENTS.md';
        case 'antigravity': return 'ANTIGRAVITY.md';
        case 'cursor': return '.cursorrules';
        case 'kiro': return '.kiro/rules/hwpx-api.md';
    }
}

export function generateInstructions(target: LlmTarget, port?: number, token?: string): string {
    const portStr = port ? String(port) : '<PORT>';
    const tokenStr = token || '<TOKEN>';
    const toolName = target === 'claude' ? 'Claude Code'
        : target === 'codex' ? 'Codex'
        : target === 'cursor' ? 'Cursor'
        : target === 'kiro' ? 'Kiro'
        : 'Antigravity';

    const shellTip = target === 'claude'
        ? 'Use the Bash tool to execute `curl` commands against the API.'
        : target === 'codex'
        ? 'Use shell commands to execute `curl` against the API.'
        : target === 'cursor'
        ? 'Use the terminal to execute `curl` commands against the API.'
        : target === 'kiro'
        ? 'Use the terminal to execute `curl` commands against the API.'
        : 'Use shell/terminal to execute `curl` commands against the API.';

    return `# HWPX Viewer for VS Code — ${toolName} Instructions

This workspace contains HWPX (한/글) documents. The **HWPX Viewer for VS Code** extension is installed and provides an HTTP API server for programmatic document access.

${shellTip}

## Connection Info

- **Base URL**: \`http://127.0.0.1:${portStr}\`
- **Token**: \`${tokenStr}\`

> **Note**: The port and token change each time the extension restarts. If you get connection errors or 401 responses, ask the user to re-run the "HWPX: Generate ${toolName} Instructions" command from the VS Code Command Palette to get updated connection info.

## Quick Start

1. Verify connectivity:
   \`\`\`bash
   curl http://127.0.0.1:${portStr}/api/help
   \`\`\`
2. Start reading/modifying the document via the API.

## Authentication

All endpoints except \`/api/help\` require a token:

\`\`\`bash
# Header
curl -H "Authorization: Bearer ${tokenStr}" http://127.0.0.1:${portStr}/api/documents

# Query parameter
curl "http://127.0.0.1:${portStr}/api/documents?token=${tokenStr}"
\`\`\`

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | \`/api/help\` | API reference (no auth required) |
| GET | \`/api/documents\` | List open HWPX documents |
| GET | \`/api/files?doc=<path>\` | List XML files inside the HWPX archive |
| GET | \`/api/xml?file=<xmlPath>\` | Raw XML content of an internal file |
| GET | \`/api/element?file=<xmlPath>&xpath=<xpath>\` | Query element by XPath (XML + JSON) |
| PUT | \`/api/element?file=<xmlPath>&xpath=<xpath>\` | Modify element |
| POST | \`/api/save?doc=<path>\` | Save to disk + auto-refresh viewer |
| POST | \`/api/reload?doc=<path>\` | Refresh viewer (no save) |

When only one document is open, the \`doc\` parameter can be omitted.

## XPath Format

HWPX uses OWPML XML (KS X 6101:2024) with namespace prefixes \`hp:\`, \`hh:\`, \`hc:\`, \`hs:\`.

- Absolute: \`/hs:sec/hp:p[2]/hp:run[0]/hp:t\` (0-based index)
- Descendant: \`/hs:sec//hp:t[0]\`
- Relative: \`hp:p[2]/hp:run[0]/hp:t\` (searches from root)
- Wrapper skip: \`hp:subList\` and similar wrappers are auto-skipped

## Modifying Elements

\`\`\`bash
# Text only
curl -X PUT -H "Authorization: Bearer ${tokenStr}" -H "Content-Type: application/json" \\
  "http://127.0.0.1:${portStr}/api/element?file=Contents/section0.xml&xpath=/hs:sec/hp:p[0]/hp:run[0]/hp:t" \\
  -d '{"text": "new text"}'

# Full XML replacement
curl -X PUT ... -d '{"xml": "<hp:t>new content</hp:t>"}'

# JSON replacement
curl -X PUT ... -d '{"json": [{"#text": "new content"}]}'
\`\`\`

## Typical Workflow

\`\`\`bash
# 1. List open documents
curl -H "Authorization: Bearer ${tokenStr}" http://127.0.0.1:${portStr}/api/documents

# 2. List internal XML files
curl -H "Authorization: Bearer ${tokenStr}" "http://127.0.0.1:${portStr}/api/files"

# 3. Read full XML to understand structure
curl -H "Authorization: Bearer ${tokenStr}" "http://127.0.0.1:${portStr}/api/xml?file=Contents/section0.xml"

# 4. Query a specific element
curl -H "Authorization: Bearer ${tokenStr}" \\
  "http://127.0.0.1:${portStr}/api/element?file=Contents/section0.xml&xpath=/hs:sec/hp:p[0]/hp:run[0]/hp:t"

# 5. Modify the element
curl -X PUT -H "Authorization: Bearer ${tokenStr}" -H "Content-Type: application/json" \\
  "http://127.0.0.1:${portStr}/api/element?file=Contents/section0.xml&xpath=/hs:sec/hp:p[0]/hp:run[0]/hp:t" \\
  -d '{"text": "modified text"}'

# 6. Save (viewer auto-refreshes)
curl -X POST -H "Authorization: Bearer ${tokenStr}" http://127.0.0.1:${portStr}/api/save
\`\`\`

## Select Mode for Precise Targeting

The user can switch to **Select** mode in the HWPX Viewer (top-right toggle), click any element, and the XPath is copied to clipboard. The format is:

\`<docPath>#<xmlFile>#<xpath>\`

Parse this to get the exact \`file\` and \`xpath\` parameters for API calls.

## HWPX Document Structure

HWPX is a ZIP archive containing XML files:
- \`mimetype\` — MIME type declaration
- \`META-INF/container.xml\` — manifest
- \`Contents/header.xml\` — styles, fonts, paragraph/character properties
- \`Contents/section0.xml\` (section1.xml, ...) — document body sections
- \`BinData/\` — embedded images and binary data

Key XML elements:
- \`hp:p\` — paragraph
- \`hp:run\` — text run (with character properties)
- \`hp:t\` — text content
- \`hp:tbl\` — table
- \`hp:tr\` / \`hp:tc\` — table row / cell
- \`hp:img\` — image
- \`hp:ctrl\` — control objects (page break, section break, etc.)
`;
}

export async function writeInstructionFile(target: LlmTarget, port?: number, token?: string): Promise<void> {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders || workspaceFolders.length === 0) {
        vscode.window.showErrorMessage('No workspace folder is open.');
        return;
    }

    const rootUri = workspaceFolders[0].uri;
    const fileName = getTargetFileName(target);
    const fileUri = vscode.Uri.joinPath(rootUri, fileName);

    // Check if file already exists
    try {
        await vscode.workspace.fs.stat(fileUri);
        const choice = await vscode.window.showWarningMessage(
            `${fileName} already exists. Overwrite?`,
            'Overwrite', 'Cancel'
        );
        if (choice !== 'Overwrite') {
            return;
        }
    } catch {
        // File doesn't exist — proceed
    }

    // Kiro: ensure .kiro/rules/ directory exists
    if (target === 'kiro') {
        const kiroDir = vscode.Uri.joinPath(rootUri, '.kiro', 'rules');
        try { await vscode.workspace.fs.createDirectory(kiroDir); } catch { /* already exists */ }
    }

    const content = generateInstructions(target, port, token);
    await vscode.workspace.fs.writeFile(fileUri, Buffer.from(content, 'utf-8'));

    const toolName = target === 'claude' ? 'Claude Code'
        : target === 'codex' ? 'Codex'
        : target === 'cursor' ? 'Cursor'
        : target === 'kiro' ? 'Kiro'
        : 'Antigravity';
    vscode.window.showInformationMessage(`${fileName} created for ${toolName}.`);
}
