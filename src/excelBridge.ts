import { spawn, ChildProcessWithoutNullStreams } from 'child_process';
import * as path from 'path';
import * as vscode from 'vscode';

interface PendingRequest {
    resolve: (value: any) => void;
    reject: (err: Error) => void;
    timer: NodeJS.Timeout;
}

/**
 * Long-running JSON-RPC bridge to the Python xlwings process.
 *
 * One process per extension activation. Requests are line-delimited JSON; the
 * Python side responds in the same order it receives them, but we still tag
 * requests with an id and route by id so concurrent requests work.
 */
export class ExcelBridge implements vscode.Disposable {
    private proc: ChildProcessWithoutNullStreams | undefined;
    private buffer = '';
    private nextId = 1;
    private pending = new Map<number, PendingRequest>();
    private readyPromise: Promise<void> | undefined;
    private readonly output: vscode.OutputChannel;

    constructor(private readonly context: vscode.ExtensionContext, output: vscode.OutputChannel) {
        this.output = output;
    }

    public scriptPath(): string {
        return path.join(this.context.extensionPath, 'python', 'excel_bridge.py');
    }

    public pythonPath(): string {
        return vscode.workspace.getConfiguration('axcelerator').get<string>('pythonPath') || 'python3';
    }

    public timeoutMs(): number {
        return vscode.workspace.getConfiguration('axcelerator').get<number>('requestTimeoutMs') ?? 120000;
    }

    private async ensureStarted(): Promise<void> {
        if (this.proc && !this.proc.killed) {
            return this.readyPromise!;
        }
        const py = this.pythonPath();
        const script = this.scriptPath();
        this.output.appendLine(`[bridge] starting: ${py} ${script}`);
        const proc = spawn(py, ['-u', script], {
            stdio: ['pipe', 'pipe', 'pipe'],
            env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
        });
        this.proc = proc;

        proc.stdout.setEncoding('utf-8');
        proc.stderr.setEncoding('utf-8');
        proc.stdout.on('data', (chunk: string) => this.onStdout(chunk));
        proc.stderr.on('data', (chunk: string) => this.output.append(`[py-stderr] ${chunk}`));
        proc.on('exit', (code, signal) => {
            this.output.appendLine(`[bridge] exited code=${code} signal=${signal}`);
            for (const p of this.pending.values()) {
                clearTimeout(p.timer);
                p.reject(new Error(`Excel bridge exited (code=${code}). See "Axcelerator" output for details.`));
            }
            this.pending.clear();
            this.proc = undefined;
            this.readyPromise = undefined;
        });

        this.readyPromise = new Promise<void>((resolve, reject) => {
            const timer = setTimeout(() => reject(new Error('Excel bridge did not signal ready in time.')), 15000);
            // The python side emits an initial {"id":null,"ok":true,"result":{"ready":true...}}.
            this.pending.set(0, {
                resolve: () => { clearTimeout(timer); resolve(); },
                reject: (e) => { clearTimeout(timer); reject(e); },
                timer,
            });
        });
        return this.readyPromise;
    }

    private onStdout(chunk: string): void {
        this.buffer += chunk;
        let nl: number;
        while ((nl = this.buffer.indexOf('\n')) !== -1) {
            const line = this.buffer.slice(0, nl).trim();
            this.buffer = this.buffer.slice(nl + 1);
            if (!line) {
                continue;
            }
            let msg: any;
            try {
                msg = JSON.parse(line);
            } catch (err) {
                this.output.appendLine(`[bridge] non-JSON line: ${line}`);
                continue;
            }
            // The "ready" message has id === null. We mapped it to id 0 in pending.
            const id = msg.id === null || msg.id === undefined ? 0 : Number(msg.id);
            const pend = this.pending.get(id);
            if (!pend) {
                this.output.appendLine(`[bridge] unmatched response id=${id}`);
                continue;
            }
            this.pending.delete(id);
            clearTimeout(pend.timer);
            if (msg.ok) {
                pend.resolve(msg.result);
            } else {
                const err = new Error(msg.error || 'Excel bridge error');
                (err as any).trace = msg.trace;
                pend.reject(err);
            }
        }
    }

    public async call(method: string, params: Record<string, any> = {}): Promise<any> {
        await this.ensureStarted();
        if (!this.proc) {
            throw new Error('Excel bridge is not running.');
        }
        const id = this.nextId++;
        const payload = JSON.stringify({ id, method, params }) + '\n';
        return new Promise((resolve, reject) => {
            const timer = setTimeout(() => {
                this.pending.delete(id);
                reject(new Error(`Excel operation timed out after ${this.timeoutMs()}ms: ${method}`));
            }, this.timeoutMs());
            this.pending.set(id, { resolve, reject, timer });
            this.proc!.stdin.write(payload, (err) => {
                if (err) {
                    this.pending.delete(id);
                    clearTimeout(timer);
                    reject(err);
                }
            });
        });
    }

    public restart(): void {
        if (this.proc && !this.proc.killed) {
            this.output.appendLine('[bridge] restart requested, killing process');
            this.proc.kill();
        }
        this.proc = undefined;
        this.readyPromise = undefined;
    }

    public dispose(): void {
        if (this.proc && !this.proc.killed) {
            this.proc.kill();
        }
    }
}
