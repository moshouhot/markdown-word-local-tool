const fs = require('fs');

const port = process.env.CDP_PORT;
const screenshotPath = process.env.SCREENSHOT_PATH;
const cdpBase = `http://127.0.0.1:${port}`;

async function waitJson(url, timeoutMs = 10000) {
  const deadline = Date.now() + timeoutMs;
  let lastErr;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(url);
      if (res.ok) return await res.json();
      lastErr = new Error(`HTTP ${res.status}`);
    } catch (e) {
      lastErr = e;
    }
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  throw lastErr || new Error(`timeout: ${url}`);
}

class CDP {
  constructor(wsUrl) {
    this.ws = new WebSocket(wsUrl);
    this.nextId = 1;
    this.pending = new Map();
    this.ws.onmessage = (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.id && this.pending.has(msg.id)) {
        const { resolve, reject } = this.pending.get(msg.id);
        this.pending.delete(msg.id);
        msg.error ? reject(new Error(JSON.stringify(msg.error))) : resolve(msg.result);
      }
    };
  }

  open() {
    return new Promise((resolve, reject) => {
      this.ws.onopen = resolve;
      this.ws.onerror = reject;
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    this.ws.send(JSON.stringify({ id, method, params }));
    return new Promise((resolve, reject) => this.pending.set(id, { resolve, reject }));
  }

  close() {
    this.ws.close();
  }
}

async function main() {
  const pages = await waitJson(`${cdpBase}/json/list`);
  const pageInfo =
    pages.find((p) => p.type === 'page' && p.url.includes('127.0.0.1')) ||
    pages.find((p) => p.type === 'page');
  if (!pageInfo) throw new Error('no CDP page found');

  const cdp = new CDP(pageInfo.webSocketDebuggerUrl);
  await cdp.open();
  await cdp.send('Runtime.enable');
  await cdp.send('Page.enable');
  await cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1280,
    height: 720,
    deviceScaleFactor: 1,
    mobile: false,
  });

  async function evalValue(expression) {
    const res = await cdp.send('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (res.exceptionDetails) throw new Error(JSON.stringify(res.exceptionDetails));
    return res.result.value;
  }

  const checks = [];
  function check(name, ok, detail = '') {
    checks.push({ name, ok: !!ok, detail });
    if (!ok) throw new Error(`check failed: ${name} ${detail}`);
  }

  const initial = await evalValue(`(() => ({
    title: document.title,
    hasInput: !!document.querySelector('#mdin'),
    hasExample: [...document.querySelectorAll('button')].some(b => b.textContent.includes('插入示例')),
    modeHelp: document.querySelector('#modeHelp')?.textContent || '',
    advancedOpen: document.querySelector('.advanced-settings')?.open,
    dockVisible: !!document.querySelector('.action-dock'),
    externalAnchors: [...document.querySelectorAll('a[href^="http"]')].map(a => a.href),
    externalResources: performance.getEntriesByType('resource').map(e => e.name).filter(n => /^https?:/.test(n) && !n.startsWith(location.origin))
  }))()`);
  check('页面标题正确', initial.title === 'Markdown → Word', initial.title);
  check('输入区存在', initial.hasInput);
  check('插入示例按钮存在', initial.hasExample);
  check('模式说明初始为模板模式', initial.modeHelp.includes('模板模式'), initial.modeHelp);
  check('高级标题设置默认折叠', initial.advancedOpen === false, String(initial.advancedOpen));
  check('底部操作区存在', initial.dockVisible);
  check('无外部 HTTP 链接', initial.externalAnchors.length === 0, initial.externalAnchors.join(','));
  check('无外部资源请求', initial.externalResources.length === 0, initial.externalResources.join(','));

  await evalValue(`document.querySelector('.input-actions .mini-btn').click(); true`);
  const afterExample = await evalValue(`(() => ({
    value: document.querySelector('#mdin').value,
    status: document.querySelector('#stBar').textContent
  }))()`);
  check('插入示例填充文本', afterExample.value.includes('项目周报'), afterExample.value.slice(0, 20));
  check('插入示例状态提示正确', afterExample.status.includes('已插入示例内容'), afterExample.status);

  await evalValue(`document.querySelector('#mb-clean').click(); true`);
  const cleanHelp = await evalValue(`document.querySelector('#modeHelp').textContent`);
  check('净化模式说明切换', cleanHelp.includes('净化模式'), cleanHelp);

  await evalValue(`document.querySelector('#mb-reflow').click(); true`);
  const reflowState = await evalValue(`(() => ({
    help: document.querySelector('#modeHelp').textContent,
    pdfDisplay: getComputedStyle(document.querySelector('#pdfBtn')).display,
    convertText: document.querySelector('#cvBtn').textContent
  }))()`);
  check('重排模式说明切换', reflowState.help.includes('重排模式'), reflowState.help);
  check('重排模式隐藏 PDF 按钮', reflowState.pdfDisplay === 'none', reflowState.pdfDisplay);
  check('重排模式主按钮文案切换', reflowState.convertText.includes('按模板重排'), reflowState.convertText);

  await evalValue(`document.querySelector('#mb-tpl').click(); true`);
  const tplState = await evalValue(`(() => ({
    help: document.querySelector('#modeHelp').textContent,
    pdfDisplay: getComputedStyle(document.querySelector('#pdfBtn')).display,
    convertText: document.querySelector('#cvBtn').textContent
  }))()`);
  check('返回模板模式说明正确', tplState.help.includes('模板模式'), tplState.help);
  check('模板模式显示 PDF 按钮', tplState.pdfDisplay !== 'none', tplState.pdfDisplay);
  check('模板模式主按钮文案恢复', tplState.convertText.includes('一键转换'), tplState.convertText);

  await evalValue(`document.querySelector('.btn-clear').click(); true`);
  const cleared = await evalValue(`(() => ({
    value: document.querySelector('#mdin').value,
    status: document.querySelector('#stBar').textContent,
    undoExists: !![...document.querySelectorAll('#stBar button')].find(b => b.textContent.includes('撤销'))
  }))()`);
  check('清空后输入区为空', cleared.value === '', cleared.value);
  check('清空后显示撤销', cleared.status.includes('已清空内容') && cleared.undoExists, cleared.status);

  await evalValue(`[...document.querySelectorAll('#stBar button')].find(b => b.textContent.includes('撤销')).click(); true`);
  const undone = await evalValue(`(() => ({
    value: document.querySelector('#mdin').value,
    status: document.querySelector('#stBar').textContent
  }))()`);
  check('撤销恢复示例文本', undone.value.includes('项目周报'), undone.value.slice(0, 20));
  check('撤销状态提示正确', undone.status.includes('已恢复清空前内容'), undone.status);

  const docx = await evalValue(`(async () => {
    const blocks = parseMarkdown(document.querySelector('#mdin').value);
    const blob = buildTplDocx(blocks, readParams());
    const buf = new Uint8Array(await blob.arrayBuffer());
    return {blocks: blocks.length, size: buf.length, signature: String.fromCharCode(buf[0]) + String.fromCharCode(buf[1])};
  })()`);
  check('浏览器内 DOCX 构建成功', docx.signature === 'PK' && docx.size > 1000 && docx.blocks >= 3, JSON.stringify(docx));

  const shot = await cdp.send('Page.captureScreenshot', {
    format: 'png',
    captureBeyondViewport: false,
  });
  fs.writeFileSync(screenshotPath, Buffer.from(shot.data, 'base64'));
  cdp.close();
  console.log(JSON.stringify({ ok: true, checks, docx, screenshotPath }, null, 2));
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : err);
  process.exit(1);
});
