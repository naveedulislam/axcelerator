#!/usr/bin/env node
/**
 * Contract test: the three sources of truth for tool names must agree.
 *
 *   1. package.json contributes.languageModelTools[].name  (what VS Code sees)
 *   2. src/extension.ts TOOLS[].toolName / .method          (what the extension dispatches)
 *   3. python/excel_bridge.py METHODS dict keys             (what the worker handles)
 *
 * Drift between any of these is a release-blocker bug, so this script exits
 * non-zero on any mismatch. Wired into `npm test`.
 */
'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');

function readText(p) {
    return fs.readFileSync(path.join(ROOT, p), 'utf-8');
}

// 1. package.json
const pkg = JSON.parse(readText('package.json'));
const pkgToolNames = new Set(
    (pkg.contributes && pkg.contributes.languageModelTools || []).map((t) => t.name),
);

// 2. extension.ts — pull both toolName and method from each TOOLS entry.
const extSrc = readText('src/extension.ts');
const toolNameMatches = [...extSrc.matchAll(/toolName:\s*'([^']+)'/g)].map((m) => m[1]);
const methodMatches = [...extSrc.matchAll(/method:\s*'([^']+)'/g)].map((m) => m[1]);
const extToolNames = new Set(toolNameMatches);
const extMethods = new Set(methodMatches);

// 3. excel_bridge.py — keys of the METHODS dict literal.
const pySrc = readText('python/excel_bridge.py');
const methodsBlock = pySrc.match(/METHODS\s*:\s*Dict[^=]*=\s*\{([\s\S]*?)\n\}/);
if (!methodsBlock) {
    console.error('contract: could not locate METHODS dict in python/excel_bridge.py');
    process.exit(2);
}
const pyMethods = new Set(
    [...methodsBlock[1].matchAll(/"([a-z_]+)"\s*:/g)].map((m) => m[1]),
);

// "tool name" in the schema is `excel_<method>` for every entry in this codebase.
// Compare on the method axis so the three sets are directly comparable.
const pkgMethods = new Set([...pkgToolNames].map((n) => n.replace(/^excel_/, '')));

const failures = [];

function diff(label, a, b, aName, bName) {
    const onlyA = [...a].filter((x) => !b.has(x)).sort();
    const onlyB = [...b].filter((x) => !a.has(x)).sort();
    if (onlyA.length || onlyB.length) {
        failures.push(
            `${label}\n  only in ${aName}: ${JSON.stringify(onlyA)}\n  only in ${bName}: ${JSON.stringify(onlyB)}`,
        );
    }
}

diff('package.json vs extension.ts (toolName)', pkgToolNames, extToolNames, 'package.json', 'extension.ts');
diff('extension.ts (method) vs python METHODS', extMethods, pyMethods, 'extension.ts', 'python');
diff('package.json (method-suffix) vs python METHODS', pkgMethods, pyMethods, 'package.json', 'python');

// extension.ts internal: toolName must equal `excel_${method}`.
const extInternal = [];
for (let i = 0; i < toolNameMatches.length; i++) {
    const tn = toolNameMatches[i];
    const m = methodMatches[i];
    if (tn !== `excel_${m}`) {
        extInternal.push(`  ${tn}  vs  excel_${m}`);
    }
}
if (extInternal.length) {
    failures.push(`extension.ts: toolName != excel_<method>\n${extInternal.join('\n')}`);
}

if (failures.length) {
    console.error('CONTRACT FAILED:');
    for (const f of failures) {
        console.error('\n' + f);
    }
    process.exit(1);
}

console.log(
    `contract OK: ${pkgToolNames.size} tools, ${extToolNames.size} extension entries, ${pyMethods.size} python methods.`,
);
