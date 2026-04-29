import { spawn, ChildProcessWithoutNullStreams } from 'child_process';
import * as path from 'path';
import * as vscode from 'vscode';

interface PendingRequest {
    resolve: (value: any) => void;
    reject: (err: Error) => void;
    timer: NodeJS.Timeout;
    /** Disposable for any cancellation listener attached to this request. */
    cancelSub?: vscode.Disposable;
    /** Wall-clock when the request was sent, used to decide whether to kill the worker. */
    sentAt: number;
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
        proc.on('error', (err) => {
            const msg = `Failed to start Python bridge using "${py}": ${err.message}. ` +
                `Set "axcelerator.pythonPath" in VS Code settings to a Python interpreter that has xlwings installed.`;
            this.output.appendLine(`[bridge] spawn error: ${msg}`);
            for (const p of this.pending.values()) {
                clearTimeout(p.timer);
                p.cancelSub?.dispose();
                p.reject(new Error(msg));
            }
            this.pending.clear();
            this.proc = undefined;
            this.readyPromise = undefined;
        });
        proc.on('exit', (code, signal) => {
            this.output.appendLine(`[bridge] exited code=${code} signal=${signal}`);
            for (const p of this.pending.values()) {
                clearTimeout(p.timer);
                p.cancelSub?.dispose();
                p.reject(new Error(`Excel bridge exited (code=${code}). See "Axcelerator" output for details.`));
            }
            this.pending.clear();
            this.proc = undefined;
            this.readyPromise = undefined;
        });

        this.readyPromise = new Promise<void>((resolve, reject) => {
            const fail = (msg: string) => {
                // Tear the bridge down so the next call respawns instead of
                // re-awaiting a permanently-rejected promise.
                this.pending.delete(0);
                if (this.proc && !this.proc.killed) {
                    try { this.proc.kill(); } catch { /* ignore */ }
                }
                this.proc = undefined;
                this.readyPromise = undefined;
                reject(new Error(msg));
            };
            const timer = setTimeout(() => fail('Excel bridge did not signal ready in time.'), 15000);
            // The python side emits an initial {"id":null,"ok":true,"result":{"ready":true...}}.
            this.pending.set(0, {
                resolve: () => { clearTimeout(timer); resolve(); },
                reject: (e) => { clearTimeout(timer); fail(e.message); },
                timer,
                sentAt: Date.now(),
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
            pend.cancelSub?.dispose();
            if (msg.ok) {
                pend.resolve(msg.result);
            } else {
                const err = new Error(msg.error || 'Excel bridge error');
                (err as any).trace = msg.trace;
                pend.reject(err);
            }
        }
    }

    /**
     * Send a JSON-RPC request to the Python worker.
     *
     * If `token` is provided and fires while the request is in flight:
     *  - the pending entry is removed and rejected with a CancellationError;
     *  - if it is the only in-flight request and has been running long enough
     *    that the worker is almost certainly inside a synchronous xlwings /
     *    AppleScript / COM call, the Python process is killed (SIGTERM). It
     *    will be respawned on the next call. This is the only reliable way
     *    to interrupt a stuck Excel automation call \u2014 xlwings calls are not
     *    cooperatively cancellable.
     *  - if other requests are also in flight we only reject the cancelled one
     *    and log a warning, to avoid taking down unrelated work.
     */
    public async call(method: string, params: Record<string, any> = {}, token?: vscode.CancellationToken): Promise<any> {
        if (token?.isCancellationRequested) {
            throw new vscode.CancellationError();
        }
        await this.ensureStarted();
        if (!this.proc) {
            throw new Error('Excel bridge is not running.');
        }
        const id = this.nextId++;
        const payload = JSON.stringify({ id, method, params }) + '\n';
        return new Promise((resolve, reject) => {
            const timer = setTimeout(() => {
                const p = this.pending.get(id);
                this.pending.delete(id);
                p?.cancelSub?.dispose();
                // The worker is almost certainly stuck inside a synchronous
                // xlwings/AppleScript/COM call. Treat the bridge as tainted:
                // kill the process so a delayed response can't land on a
                // future request, and so Excel mutations stop. The bridge
                // auto-respawns on the next call.
                if (this.proc && !this.proc.killed) {
                    this.output.appendLine(`[bridge] timeout: killing worker (request "${method}" exceeded ${this.timeoutMs()}ms)`);
                    try { this.proc.kill(); } catch { /* ignore */ }
                }
                reject(new Error(`Excel operation timed out after ${this.timeoutMs()}ms: ${method}`));
            }, this.timeoutMs());
            const entry: PendingRequest = { resolve, reject, timer, sentAt: Date.now() };
            if (token) {
                entry.cancelSub = token.onCancellationRequested(() => {
                    const pend = this.pending.get(id);
                    if (!pend) {
                        return;
                    }
                    this.pending.delete(id);
                    clearTimeout(pend.timer);
                    pend.cancelSub?.dispose();
                    const inFlightMs = Date.now() - pend.sentAt;
                    const otherInFlight = this.pending.size > 0;
                    if (!otherInFlight && inFlightMs > 2000 && this.proc && !this.proc.killed) {
                        this.output.appendLine(`[bridge] cancellation: killing worker (request "${method}" was in-flight for ${inFlightMs}ms)`);
                        this.proc.kill();
                    } else if (otherInFlight) {
                        this.output.appendLine(`[bridge] cancellation: "${method}" cancelled but ${this.pending.size} other request(s) in flight; not killing worker`);
                    }
                    pend.reject(new vscode.CancellationError());
                });
            }
            this.pending.set(id, entry);
            this.proc!.stdin.write(payload, (err) => {
                if (err) {
                    const p = this.pending.get(id);
                    this.pending.delete(id);
                    clearTimeout(timer);
                    p?.cancelSub?.dispose();
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
