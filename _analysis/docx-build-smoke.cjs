const fs = require('fs');
const html = fs.readFileSync('index.html', 'utf8');
const start = html.indexOf('// ══════════════ MiniZip');
const end = html.indexOf('// ══════════════ 核心变量');
if (start < 0 || end < 0) throw new Error('script slice markers not found');
const code = html.slice(start, end) + `
(async () => {
  const blocks = parseMarkdown('# 标题\\n\\n正文内容\\n\\n|列1|列2|\\n|-|-|\\n|A|B|');
  const blob = buildTplDocx(blocks, {
    bodyFont: '仿宋_GB2312', bodySize: 16, bodyLine: 28, bodyIndent: true,
    h1Font: '黑体', h1Size: 16, h1Bold: true,
    h2Font: '楷体', h2Size: 16, h2Bold: true,
    h3Font: '仿宋_GB2312', h3Size: 16, h3Bold: true
  });
  const buf = Buffer.from(await blob.arrayBuffer());
  if (buf.slice(0, 2).toString() !== 'PK') throw new Error('DOCX zip signature missing');
  if (!buf.includes(Buffer.from('word/document.xml'))) throw new Error('document.xml missing');
  console.log('docx-build-ok', blocks.length, buf.length);
})().catch(err => { console.error(err); process.exit(1); });`;
eval(code);
