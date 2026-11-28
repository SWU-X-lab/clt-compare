const assert = require('node:assert/strict');
const { test } = require('node:test');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');

const script = fs.readFileSync(path.join(__dirname, '..', 'exp1_clt_compare.js'), 'utf8');
function app({ realExport = false, onSave = null, workbookError = null, autoDownloadError = null } = {}) {
    const calls = [];
    const captured = { autoClicks: 0, workbooks: [] };
    const timers = new Map();
    const listeners = {};
    const elements = new Map();
    let nextTimer = 1;
    const quiet = { log() {}, warn() {}, error() {} };
    const store = Object.fromEntries(['initializeStorage', 'saveTrial', 'saveCheckpoint', 'loadProgress',
        'saveProgress', 'saveFinalResult', 'clearProgress', 'finishExperiment'].map(name => [
        name, async (...args) => {
            calls.push([name, ...args]);
            return name === 'saveFinalResult' && onSave ? onSave(...args) : null;
        }
    ]));
    const node = id => {
        if (!elements.has(id)) elements.set(id, {
            value: '', style: {}, classList: { add() {}, remove() {} },
            addEventListener(type, fn) { listeners[id + ':' + type] = fn; },
            reportValidity: () => true, getBoundingClientRect: () => ({ width: 800, height: 600 }),
            getContext() { return new Proxy({}, { get: () => () => {} }); }, focus() {}
        });
        return elements.get(id);
    };
    const document = {
        documentElement: {}, hidden: false, fullscreenElement: null, pointerLockElement: null,
        getElementById: node,
        body: { appendChild() {} },
        createElement() {
            return { style: {}, remove() {}, click() { captured.autoClicks++; if (autoDownloadError) throw autoDownloadError; } };
        },
        addEventListener(type, fn) { listeners['document:' + type] = fn; },
        exitPointerLock() { document.pointerLockElement = null; },
        exitFullscreen: async () => { document.fullscreenElement = null; }
    };
    const canvas = node('expCanvas');
    canvas.width = 800;
    canvas.height = 600;
    canvas.requestPointerLock = async () => {
        document.pointerLockElement = canvas;
        listeners['document:pointerlockchange']();
    };
    document.documentElement.requestFullscreen = async () => { document.fullscreenElement = document.documentElement; };
    const g = {
        document,
        window: {
            CLTStorage: store, innerWidth: 800, innerHeight: 600,
            addEventListener(type, fn) {
                // 保留同类型的所有处理器，并按注册顺序发出。
                const prev = listeners['window:' + type];
                listeners['window:' + type] = prev ? (...a) => { prev(...a); return fn(...a); } : fn;
            }
        },
        history: { pushState() {} },
        Blob,
        URL: { createObjectURL: () => 'blob:local-first' },
        XLSX: {
            utils: {
                json_to_sheet: (rows, options) => ({ rows, options }),
                book_new: () => ({ SheetNames: [], Sheets: {} }),
                book_append_sheet: (wb, sheet, name) => { wb.SheetNames.push(name); wb.Sheets[name] = sheet; }
            },
            write(wb) {
                if (workbookError) throw workbookError;
                captured.workbooks.push(wb);
                return Uint8Array.from([0x50, 0x4b, 1, 2]);
            }
        },
        console: quiet,
        performance: { now: () => 1000 },
        setTimeout(fn, ms) { timers.set(nextTimer, [fn, ms]); return nextTimer++; },
        clearTimeout(id) { timers.delete(id); },
        setInterval() { return nextTimer++; },
        clearInterval() {},
        alert() { throw Error('unexpected alert'); }
    };
    const context = vm.createContext(g);
    const run = code => vm.runInContext(code, context);
    run(script);
    if (!realExport) run('downloadXLSX = () => { window.localDownloads = (window.localDownloads || 0) + 1; };');
    return { calls, listeners, node, document, canvas, timers, run, g, captured };
}
const row = a => JSON.parse(JSON.stringify(a));
const start = async (a, group, code) => {
    a.node('subGroup').value = group;
    a.node('participantCode').value = code;
    await a.listeners['start-btn:click']();
};
const probe = a => a.run("trialData = { setSize: 4, studyDuration: 250, block: 2, TrialGlobal: 21 }; labelMap = [0,1,2,3]; probeColorArr = [1,2,3,4]; correctLabel = 1; changeColorIdx = 1; showTestArray({ locs: [{x:410,y:300}], probeColors:[1] });");

test('空条件 / 全超时 / 真正零正确率 / 浏览器中断分别统计', () => {
    const a = app();
    const evalK = records => row(a.run('buildKRow(4,250,' + JSON.stringify(records) + ')'));
    const empty = evalK([]);
    assert.equal(empty.K_Exclude, '');
    assert.equal(empty.K_TimeoutErr, '');
    assert.equal(empty.Accuracy_Exclude, '');
    const timeout = evalK([{ accuracy: 0, timeout: true, TrialStatus: 'timeout', CountedInK: 0 }]);
    assert.equal(timeout.K_Exclude, '');
    assert.equal(timeout.Accuracy_TimeoutErr, '0.0000');
    assert.equal(timeout.K_TimeoutErr, '-1.3333');
    assert.equal(timeout.TimeoutTrials, 1);
    const interrupted = evalK([{ accuracy: null, timeout: false, TrialStatus: 'interrupted', CountedInK: 0 }]);
    assert.equal(interrupted.InterruptedTrials, 1);
    assert.equal(interrupted.TimeoutTrials, 0);
    assert.equal(interrupted.K_TimeoutErr, '');
    const wrong = evalK([{ accuracy: 0, timeout: false, TrialStatus: 'valid', CountedInK: 1, rt: 1 }]);
    assert.equal(wrong.K_Exclude, '-1.3333');
    assert.equal(wrong.MeanRT, '');
});

test('组0从开始到完成仅下载本地文件，不调用存储', async () => {
    const a = app();
    await start(a, '0', 'ABCDEF');
    assert.equal(a.run('p.subjectID'), '1111');
    probe(a);
    a.run('recordResponse(null); endBlock();');
    await new Promise(setImmediate);
    assert.equal(a.g.window.localDownloads, 1);
    assert.deepEqual(a.calls, []);
});

test('正式匿名码稳定，运行会话独立，点击坐标及距离保留', async () => {
    const a = app(), b = app();
    await start(a, '1', 'Stable-42');
    await start(b, '2', 'Stable-42');
    assert.equal(a.run('p.subjectID'), 'Stable-42');
    assert.equal(a.run('p.rndSeed'), b.run('p.rndSeed'));
    assert.notEqual(a.run('SESSION_ID'), b.run('SESSION_ID'));
    probe(a);
    a.listeners['window:mousemove']({ movementX: 10, movementY: 0, clientX: 0, clientY: 0 });
    a.listeners['expCanvas:mousedown']({ button: 0, preventDefault() {} });
    const trial = row(a.run('stimData[0]'));
    assert.equal(trial.TrialStatus, 'valid');
    assert.equal(trial.TrialGlobal, 21);
    assert.equal(trial.clickedItemIndex, 0);
    assert.equal(trial.cursorStartX, 400);
    assert.equal(trial.cursorEndX, 410);
    assert.equal(trial.cursorTravelPx, 10);
    assert.ok(a.run("TRIAL_EXPORT_COLUMNS.some(c => c.header === 'CursorTravelPx')"));
});

test('探测中失去鼠标锁或切后台只记录一次中断并取消超时', async () => {
    for (const event of ['pointerlockchange','visibilitychange','fullscreenchange']) {
        const a = app();
        await start(a, '1', 'Stable-42');
        probe(a);
        const ordinaryTimeout = [...a.timers.keys()][0];
        if (event === 'pointerlockchange') a.document.pointerLockElement = null;
        if (event === 'visibilitychange') a.document.hidden = true;
        if (event === 'fullscreenchange') a.document.fullscreenElement = null;
        a.listeners['document:' + event]();
        a.listeners['document:' + event]();
        await new Promise(setImmediate);
        const interrupted = row(a.run('stimData'));
        assert.equal(a.timers.has(ordinaryTimeout), false);
        assert.equal(interrupted.length, 1);
        assert.equal(interrupted[0].TrialStatus, 'interrupted');
        assert.equal(interrupted[0].timeout, false);
        assert.equal(interrupted[0].accuracy, null);
        assert.ok(interrupted[0].interruptionReason);
        assert.equal(a.calls.find(([name]) => name === 'finishExperiment')[1].successful, false);
        assert.ok(!a.calls.some(([name]) => name === 'clearProgress'));
        assert.equal(a.g.window.localDownloads, 1);
    }
});

test('Supabase 基础列回退明确标记部分保存，不访问线上服务', async () => {
    const pending = [];
    const window = { supabase: { createClient: () => ({
        from: () => ({ async insert(rows) {
            pending.push(rows);
            return { error: pending.length === 1 ? Error('missing columns') : null };
        } })
    }) } };
    const adapter = fs.readFileSync(path.join(__dirname, '..', 'src/storage/supabase-storage.js'), 'utf8');
    vm.runInNewContext(adapter, { window, console: { error() {} } });
    const storage = window.CLTStorageFactories.supabase();
    await storage.initialize();
    const result = await storage.saveFinalResult({ trials: [{subjectID:'stable42', cursorTravelPx:10}] });
    assert.equal(result.partial, true);
    assert.match(result.message, /本地 Excel/);
    assert.equal(pending[0][0].cursorTravelPx, 10);
    assert.equal(pending[1][0].cursorTravelPx, undefined);
});

test('正式组：未完成的云端保存不阻塞本地工作簿、上传页下载链接和单次自动下载', async () => {
    let settle;
    const a = app({ realExport: true, onSave: () => new Promise(resolve => { settle = resolve; }) });
    await start(a, '1', 'Stable-42');
    probe(a);
    a.run('recordResponse(null);');
    const done = a.run('finishExperiment()');
    assert.equal(a.captured.workbooks.length, 1);
    assert.equal(a.captured.autoClicks, 1);
    const xlsx = a.captured.workbooks[0].Sheets['Trial Data'].rows[0];
    assert.equal(xlsx.SubjectID, 'Stable-42');
    assert.equal(xlsx.TrialGlobal, 21);
    assert.equal(xlsx.TrialStatus, 'timeout');
    assert.equal(a.node('upload-manual-download-link').href, 'blob:local-first');
    assert.equal(a.node('manual-download-link').href, 'blob:local-first');
    assert.equal(a.document.pointerLockElement, null);
    assert.match(a.node('upload-local-status').innerText, /本地 Excel：已生成/);
    assert.match(a.node('upload-status').innerText, /云端：正在上传/);
    assert.equal(a.node('finished-text').innerText, undefined);
    assert.equal(a.calls.find(([name]) => name === 'saveFinalResult')[1].trials[0].TrialGlobal, 21);
    settle({ partial: true, message: '服务器仅保存基础字段' });
    await done;
    assert.match(a.node('finished-text').innerText, /本地 Excel：已生成[\s\S]*云端：服务器仅保存基础字段/);
    assert.equal(a.captured.autoClicks, 1);
});

test('正式组：云端失败时本地链接保留，自动下载异常不损坏手动链接', async () => {
    const a = app({
        realExport: true,
        autoDownloadError: new Error('自动下载被浏览器阻止'),
        onSave: async () => { throw Error('连接断开'); }
    });
    await start(a, '2', 'Stable-42');
    probe(a);
    a.run('recordResponse(null);');
    await a.run('finishExperiment()');
    assert.equal(a.captured.workbooks.length, 1);
    assert.equal(a.node('upload-manual-download-link').href, 'blob:local-first');
    assert.match(a.node('finished-text').innerText, /本地 Excel：已生成[\s\S]*云端：保存失败（连接断开）/);
});

test('正式组：未选本地导出或工作簿失败时仍保存云端且无虚假下载入口', async () => {
    for (const workbookError of [null, new Error('示例导出失败')]) {
        const a = app({ realExport: true, workbookError, onSave: async () => ({ message: '数据上传服务器成功。' }) });
        await start(a, '1', 'Stable-42');
        if (!workbookError) a.run('p.saveLocal = false;');
        probe(a);
        a.run('recordResponse(null);');
        await a.run('finishExperiment()');
        assert.equal(a.captured.workbooks.length, 0);
        assert.equal(a.captured.autoClicks, 0);
        assert.equal(a.node('upload-manual-download-link').href, undefined);
        assert.match(a.node('finished-text').innerText,
            workbookError ? /本地 Excel：生成失败[\s\S]*云端：数据上传服务器成功/ : /本地 Excel：未选择导出[\s\S]*云端：数据上传服务器成功/);
    }
});
