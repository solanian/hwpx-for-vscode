import * as vscode from 'vscode';
import { HwpxEditorProvider } from './hwpxEditorProvider';

export function activate(context: vscode.ExtensionContext) {
    console.log('HWPX Viewer extension is now active!');

    // Register our custom editor provider
    context.subscriptions.push(HwpxEditorProvider.register(context));
}

export function deactivate() {}
