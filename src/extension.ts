import * as vscode from 'vscode';
import { ExcelBridge } from './excelBridge';

/**
 * Definition of an LM tool that proxies to a Python bridge method.
 *
 * `gated` tools are only invocable when the user has flipped a settings flag,
 * because they let Copilot run arbitrary code (VBA, Python).
 */
interface ToolDef {
    /** Name as declared in package.json `contributes.languageModelTools[].name`. */
    toolName: string;
    /** Method name on the Python bridge. */
    method: string;
    /** Optional settings flag that must be enabled. */
    gatedBy?: 'axcelerator.allowVba' | 'axcelerator.allowPython';
    /** Short label used for the prepared invocation message. */
    label: string;
    /** Build a one-line invocation summary for the user. */
    summary?: (input: any) => string;
}

const TOOLS: ToolDef[] = [
    { toolName: 'excel_check_environment', method: 'check_environment', label: 'Check Excel environment' },
    { toolName: 'excel_list_workbooks', method: 'list_workbooks', label: 'List open workbooks' },
    {
        toolName: 'excel_open_workbook', method: 'open_workbook', label: 'Open workbook',
        summary: (i) => i?.path ? `Open ${i.path}` : 'Create new workbook',
    },
    { toolName: 'excel_save_workbook', method: 'save_workbook', label: 'Save workbook', summary: (i) => `Save ${i?.workbook ?? ''}` },
    { toolName: 'excel_close_workbook', method: 'close_workbook', label: 'Close workbook', summary: (i) => `Close ${i?.workbook ?? ''}` },
    { toolName: 'excel_list_sheets', method: 'list_sheets', label: 'List sheets', summary: (i) => `List sheets in ${i?.workbook ?? ''}` },
    { toolName: 'excel_add_sheet', method: 'add_sheet', label: 'Add sheet', summary: (i) => `Add sheet "${i?.name}" to ${i?.workbook ?? ''}` },
    { toolName: 'excel_delete_sheet', method: 'delete_sheet', label: 'Delete sheet', summary: (i) => `Delete sheet "${i?.sheet}"` },
    { toolName: 'excel_read_range', method: 'read_range', label: 'Read range', summary: (i) => `Read ${i?.sheet}!${i?.range}` },
    { toolName: 'excel_write_range', method: 'write_range', label: 'Write range', summary: (i) => `Write ${i?.sheet}!${i?.range}` },
    { toolName: 'excel_set_formula', method: 'set_formula', label: 'Set formula', summary: (i) => `Set formula on ${i?.sheet}!${i?.range}` },
    { toolName: 'excel_format_range', method: 'format_range', label: 'Format range', summary: (i) => `Format ${i?.sheet}!${i?.range}` },
    { toolName: 'excel_create_table', method: 'create_table', label: 'Create Excel table', summary: (i) => `Create table "${i?.name}" on ${i?.sheet}!${i?.range}` },
    { toolName: 'excel_create_chart', method: 'create_chart', label: 'Create chart', summary: (i) => `Create ${i?.chartType ?? 'chart'} on ${i?.sheet}` },
    { toolName: 'excel_create_pivot_table', method: 'create_pivot_table', label: 'Create PivotTable', summary: (i) => `PivotTable "${i?.name}" from ${i?.sourceTable}` },
    { toolName: 'excel_add_power_query', method: 'add_power_query', label: 'Add Power Query', summary: (i) => `Add Power Query "${i?.queryName}"` },
    { toolName: 'excel_refresh', method: 'refresh', label: 'Refresh queries', summary: (i) => i?.queryName ? `Refresh ${i.queryName}` : 'Refresh all' },
    { toolName: 'excel_run_vba', method: 'run_vba', label: 'Run VBA macro', gatedBy: 'axcelerator.allowVba', summary: (i) => `Run VBA: ${i?.macro}` },
    { toolName: 'excel_run_python', method: 'run_python', label: 'Run xlwings Python', gatedBy: 'axcelerator.allowPython' },
];

class ExcelTool implements vscode.LanguageModelTool<any> {
    constructor(private readonly def: ToolDef, private readonly bridge: ExcelBridge) {}

    async prepareInvocation(
        options: vscode.LanguageModelToolInvocationPrepareOptions<any>,
        _token: vscode.CancellationToken,
    ): Promise<vscode.PreparedToolInvocation> {
        const summary = this.def.summary ? this.def.summary(options.input) : this.def.label;
        const prepared: vscode.PreparedToolInvocation = {
            invocationMessage: summary,
        };
        // Anything that mutates the workbook or runs code asks for confirmation
        // unless it is a pure read.
        const mutating = !['check_environment', 'list_workbooks', 'list_sheets', 'read_range'].includes(this.def.method);
        if (mutating) {
            prepared.confirmationMessages = {
                title: this.def.label,
                message: new vscode.MarkdownString(
                    `Allow Copilot to perform this action in Excel?\n\n**${summary}**`,
                ),
            };
        }
        return prepared;
    }

    async invoke(
        options: vscode.LanguageModelToolInvocationOptions<any>,
        _token: vscode.CancellationToken,
    ): Promise<vscode.LanguageModelToolResult> {
        if (this.def.gatedBy) {
            const [section, key] = this.def.gatedBy.split('.');
            const enabled = vscode.workspace.getConfiguration(section).get<boolean>(key);
            if (!enabled) {
                return new vscode.LanguageModelToolResult([
                    new vscode.LanguageModelTextPart(
                        `Refused: tool "${this.def.toolName}" is disabled. ` +
                        `Set "${this.def.gatedBy}" to true in VS Code settings to enable it.`,
                    ),
                ]);
            }
        }
        try {
            const result = await this.bridge.call(this.def.method, options.input ?? {});
            const text = JSON.stringify(result, null, 2);
            return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(text)]);
        } catch (err: any) {
            const msg = err?.message ?? String(err);
            return new vscode.LanguageModelToolResult([
                new vscode.LanguageModelTextPart(`Error: ${msg}`),
            ]);
        }
    }
}

export function activate(context: vscode.ExtensionContext): void {
    const output = vscode.window.createOutputChannel('Axcelerator');
    context.subscriptions.push(output);

    const bridge = new ExcelBridge(context, output);
    context.subscriptions.push(bridge);

    for (const def of TOOLS) {
        context.subscriptions.push(
            vscode.lm.registerTool(def.toolName, new ExcelTool(def, bridge)),
        );
    }

    context.subscriptions.push(
        vscode.commands.registerCommand('axcelerator.checkEnvironment', async () => {
            try {
                const info = await bridge.call('check_environment');
                const msg = `OS: ${info.os} • xlwings ${info.xlwingsVersion} • Excel running: ${info.excelRunning}` +
                    (info.excelVersion ? ` • Excel ${info.excelVersion}` : '') +
                    `\nCOM/VBA/Power Query: ${info.comAvailable ? 'available' : 'unavailable (non-Windows)'}`;
                vscode.window.showInformationMessage(`Axcelerator: ${msg}`);
                output.appendLine(msg);
            } catch (err: any) {
                vscode.window.showErrorMessage(`Axcelerator check failed: ${err.message}`);
            }
        }),
        vscode.commands.registerCommand('axcelerator.restartBridge', () => {
            bridge.restart();
            vscode.window.showInformationMessage('Axcelerator bridge restarted.');
        }),
    );

    output.appendLine(`Axcelerator activated. ${TOOLS.length} LM tools registered.`);
}

export function deactivate(): void {
    /* no-op; disposables handle cleanup */
}
