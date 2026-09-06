// update_talk.py가 호출하는 헤드리스 Chrome 스크린샷 드라이버.
// 사용: node _shot_driver.js <url> <evalJs> <출력png경로>
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

function findChrome() {
  const candidates = [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  ];
  for (const c of candidates) if (fs.existsSync(c)) return c;
  throw new Error('Chrome을 찾지 못했습니다');
}

const CHROME = findChrome();
const PORT = 9377;
const URL = process.argv[2];
const EVAL_JS = process.argv[3] || '';
const OUT_PNG = process.argv[4];
const userDataDir = path.join(require('os').tmpdir(), 'update_talk_shot_' + Date.now());

function wait(ms) { return new Promise(r => setTimeout(r, ms)); }

async function main() {
  const proc = spawn(CHROME, [
    '--headless=new', '--disable-gpu', '--no-sandbox',
    `--remote-debugging-port=${PORT}`, `--user-data-dir=${userDataDir}`,
    '--window-size=1280,2000', 'about:blank',
  ], { stdio: 'ignore' });

  try {
    let targets;
    for (let i = 0; i < 30; i++) {
      await wait(400);
      try {
        const res = await fetch(`http://127.0.0.1:${PORT}/json`);
        targets = await res.json();
        if (targets && targets.length) break;
      } catch (e) {}
    }
    const target = targets.find(t => t.type === 'page');
    const ws = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => { ws.onopen = resolve; ws.onerror = reject; });

    let id = 0;
    const pending = new Map();
    ws.onmessage = (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.id && pending.has(msg.id)) { pending.get(msg.id)(msg); pending.delete(msg.id); }
    };
    function send(method, params = {}) {
      return new Promise((resolve) => {
        const thisId = ++id;
        pending.set(thisId, resolve);
        ws.send(JSON.stringify({ id: thisId, method, params }));
      });
    }

    await send('Page.enable');
    await send('Page.navigate', { url: URL });
    await wait(1500);
    if (EVAL_JS) {
      await send('Runtime.evaluate', { expression: EVAL_JS });
      await wait(800);
    }
    const shot = await send('Page.captureScreenshot', { format: 'png' });
    fs.writeFileSync(OUT_PNG, Buffer.from(shot.result.data, 'base64'));
    ws.close();
  } finally {
    proc.kill();
    try { fs.rmSync(userDataDir, { recursive: true, force: true }); } catch (e) {}
  }
}

main().catch(e => { console.error('FAILED', e); process.exit(1); });
