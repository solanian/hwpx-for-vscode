import * as vscode from 'vscode';
import { HwpxEditorProvider } from './hwpxEditorProvider';
import { HwpxApiServer } from './hwpxApiServer';

let apiServer: HwpxApiServer | null = null;

export function getApiServer(): HwpxApiServer | null {
    return apiServer;
}

export async function activate(context: vscode.ExtensionContext) {
    console.log('HWPX Viewer extension is now active!');

    // Start API server
    apiServer = new HwpxApiServer();
    try {
        const port = await apiServer.start();
        const token = apiServer.getToken();
        vscode.window.showInformationMessage(`HWPX API: http://127.0.0.1:${port} (token copied to clipboard)`);

        // Register commands
        context.subscriptions.push(
            vscode.commands.registerCommand('hwpx.showApiPort', async () => {
                const info = `http://127.0.0.1:${apiServer?.getPort()}`;
                const tokenStr = apiServer?.getToken() || '';
                const selected = await vscode.window.showInformationMessage(
                    `HWPX API: ${info}`,
                    'Copy Token', 'Copy URL'
                );
                if (selected === 'Copy Token') {
                    await vscode.env.clipboard.writeText(tokenStr);
                } else if (selected === 'Copy URL') {
                    await vscode.env.clipboard.writeText(`${info}?token=${tokenStr}`);
                }
            }),
            vscode.commands.registerCommand('hwpx.copyApiHelp', async () => {
                const helpUrl = `http://127.0.0.1:${apiServer?.getPort()}/api/help`;
                await vscode.env.clipboard.writeText(helpUrl);
                vscode.window.showInformationMessage(`Copied: ${helpUrl}`);
            })
        );
    } catch (err: any) {
        console.error('Failed to start HWPX API server:', err);
        vscode.window.showWarningMessage(`HWPX API server failed to start: ${err.message}`);
    }

    // Register custom editor provider
    context.subscriptions.push(HwpxEditorProvider.register(context));

    // Cleanup on deactivation
    context.subscriptions.push({
        dispose: () => {
            if (apiServer) {
                apiServer.stop();
                apiServer = null;
            }
        }
    });
}

export function deactivate() {
    if (apiServer) {
        apiServer.stop();
        apiServer = null;
    }
}
