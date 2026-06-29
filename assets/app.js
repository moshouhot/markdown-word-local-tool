'use strict';

// ══════════════ MiniZip ══════════════
const MiniZip = (() => {
  const ENC = new TextEncoder();
  const CRC_TBL = (() => { const t = new Uint32Array(256); for (let i = 0; i < 256; i++) { let c = i; for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[i] = c >>> 0; } return t; })();
  function crc32(data) { let c = 0xFFFFFFFF; for (let i = 0; i < data.length; i++) c = CRC_TBL[(c ^ data[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  const w16 = v => { v >>>= 0; return [v & 0xFF, (v >>> 8) & 0xFF]; };
  const w32 = v => { v >>>= 0; return [v&0xFF,(v>>>8)&0xFF,(v>>>16)&0xFF,(v>>>24)&0xFF]; };
  return class {
    constructor(){ this.entries=[]; this.offset=0; this.parts=[]; }
    add(name, content){
      const nameB=ENC.encode(name), data=typeof content==='string'?ENC.encode(content):content;
      const crc=crc32(data), sz=data.length;
      const lh=new Uint8Array([0x50,0x4B,0x03,0x04,20,0,0,0,0,0,0,0,0x21,0x4A,...w32(crc),...w32(sz),...w32(sz),...w16(nameB.length),0,0,...nameB]);
      this.entries.push({nameB, crc, sz, offset:this.offset}); this.parts.push(lh, data); this.offset+=lh.length+sz;
    }
    toBlob(){
      const cdParts=[]; for(const e of this.entries){ cdParts.push(new Uint8Array([0x50,0x4B,0x01,0x02,20,0,20,0,0,0,0,0,0,0,0x21,0x4A,...w32(e.crc),...w32(e.sz),...w32(e.sz),...w16(e.nameB.length),0,0,0,0,0,0,0,0,0,0,0,0,...w32(e.offset),...e.nameB])); }
      const cdSize=cdParts.reduce((s,a)=>s+a.length,0);
      const eocd=new Uint8Array([0x50,0x4B,0x05,0x06,0,0,0,0,...w16(this.entries.length),...w16(this.entries.length),...w32(cdSize),...w32(this.offset),0,0]);
      return new Blob([...this.parts,...cdParts,eocd],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
    }
  };
})();

// ══════════════ ZIP 读取器 ══════════════
async function readZipEntry(arrayBuffer, targetName) {
  const view=new DataView(arrayBuffer), buf=new Uint8Array(arrayBuffer), len=buf.length;
  let eocdPos=-1;
  for(let i=len-22; i>=Math.max(0, len-65558); i--){ if(view.getUint32(i, true)===0x06054b50){ eocdPos=i; break; } }
  if(eocdPos<0) throw new Error('不是有效的 ZIP 文件');
  const cdOffset=view.getUint32(eocdPos+16, true), cdEntries=view.getUint16(eocdPos+10, true);
  let pos=cdOffset;
  for(let i=0; i<cdEntries; i++){
    if(view.getUint32(pos, true)!==0x02014b50) break;
    const method=view.getUint16(pos+10, true), compSize=view.getUint32(pos+20, true);
    const nameLen=view.getUint16(pos+28, true), extraLen=view.getUint16(pos+30, true), commentLen=view.getUint16(pos+32, true);
    const lhOffset=view.getUint32(pos+42, true);
    const name=new TextDecoder().decode(buf.slice(pos+46, pos+46+nameLen));
    pos+=46+nameLen+extraLen+commentLen;
    if(name!==targetName) continue;
    if(view.getUint32(lhOffset, true)!==0x04034b50) throw new Error('本地文件头损坏');
    const lhNameLen=view.getUint16(lhOffset+26, true), lhExtraLen=view.getUint16(lhOffset+28, true);
    const dataStart=lhOffset+30+lhNameLen+lhExtraLen;
    const compData=buf.slice(dataStart, dataStart+compSize);
    if(method===0){ return new TextDecoder('utf-8').decode(compData); }
    else if(method===8){
      const ds=new DecompressionStream('deflate-raw'), w=ds.writable.getWriter();
      w.write(compData); w.close();
      const chunks=[], r=ds.readable.getReader();
      while(true){ const {done, value}=await r.read(); if(done) break; chunks.push(value); }
      const total=chunks.reduce((s,c)=>s+c.length,0), out=new Uint8Array(total);
      let off=0; for(const c of chunks){ out.set(c, off); off+=c.length; }
      return new TextDecoder('utf-8').decode(out);
    } else throw new Error(`不支持的压缩方式 (method=${method})`);
  }
  throw new Error(`文件中未找到 "${targetName}"`);
}

// ══════════════ DOCX 生成器 ══════════════
const xe=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
function spacingXml(line){ if(!line||line<=0) return '<w:spacing w:before="0" w:after="0"/>'; if(line>=6) return `<w:spacing w:line="${Math.round(line*20)}" w:lineRule="exact" w:before="0" w:after="0"/>`; return `<w:spacing w:line="${Math.round(240*line)}" w:lineRule="auto" w:before="0" w:after="0"/>`; }
const fontXml=f=>f?`<w:rFonts w:ascii="${xe(f)}" w:eastAsia="${xe(f)}" w:hAnsi="${xe(f)}" w:cs="${xe(f)}"/>`:'';
function runXml({text,bold,italic,strike,code,font,size}){
  const f=code?'Consolas':(font||'宋体'), sz=(code?Math.max(9,(size||12)-1):(size||12));
  return `<w:r><w:rPr>${fontXml(f)}<w:sz w:val="${sz*2}"/><w:szCs w:val="${sz*2}"/>${bold?'<w:b/><w:bCs/>':''}${italic?'<w:i/><w:iCs/>':''}${strike?'<w:strike/>':''}</w:rPr><w:t xml:space="preserve">${xe(text)}</w:t></w:r>`;
}
function paraXml({runs,line,font,size,indTwips,leftTwips=0,align}){
  const sp=spacingXml(line);
  let ind=''; if(indTwips>0&&leftTwips>0) ind=`<w:ind w:firstLine="${indTwips}" w:firstLineChars="200" w:left="${leftTwips}"/>`; else if(indTwips>0) ind=`<w:ind w:firstLine="${indTwips}" w:firstLineChars="200"/>`; else if(leftTwips>0) ind=`<w:ind w:left="${leftTwips}"/>`;
  const jc=`<w:jc w:val="${align||'both'}"/>`;
  const ppr=`<w:pPr>${sp}${ind}${jc}</w:pPr>`;
  const rx=runs.map(r=>runXml({...r,font:r.font||font,size:r.size||size})).join('');
  return `<w:p>${ppr}${rx}</w:p>`;
}
function headingXml({level,runs,line,font,size,indTwips,bold=true}){
  const styleId=['Heading1','Heading2','Heading3'][Math.min(level,3)-1];
  const sp=spacingXml(line);
  const ind=indTwips>0?`<w:ind w:firstLine="${indTwips}"/>`:'';
  const ppr=`<w:pPr><w:pStyle w:val="${styleId}"/>${sp}${ind}</w:pPr>`;
  const rx=runs.map(r=>runXml({...r,font:r.font||font,size:r.size||size,bold})).join('');
  return `<w:p>${ppr}${rx}</w:p>`;
}
function tableXml({rows,nc,font,size,line}){
  const W=9000, cw=Math.floor(W/nc), cws=Array(nc).fill(cw); cws[nc-1]=W-cw*(nc-1);
  const bd=['top','left','bottom','right','insideH','insideV'].map(s=>`<w:${s} w:val="single" w:sz="6" w:space="0" w:color="888888"/>`).join('');
  const sp=spacingXml(line);
  const tblPr=`<w:tblPr><w:tblW w:w="${W}" w:type="dxa"/><w:tblBorders>${bd}</w:tblBorders><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>`;
  const tblGrid=`<w:tblGrid>${cws.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  const rowsXml=rows.map((row,ri)=>{
    const trPr=ri===0?'<w:trPr><w:tblHeader/></w:trPr>':'';
    const cells=row.map((cell,ci)=>{
      const txt=xe(String(cell).replace(/\*\*/g,'').replace(/`/g,'').trim());
      const rpr=`<w:rPr>${fontXml(font)}<w:sz w:val="${size*2}"/><w:szCs w:val="${size*2}"/>${ri===0?'<w:b/><w:bCs/>':''}</w:rPr>`;
      return `<w:tc><w:tcPr><w:tcW w:w="${cws[ci]}" w:type="dxa"/></w:tcPr><w:p><w:pPr>${sp}</w:pPr><w:r>${rpr}<w:t xml:space="preserve">${txt}</w:t></w:r></w:p></w:tc>`;
    }).join('');
    return `<w:tr>${trPr}${cells}</w:tr>`;
  }).join('');
  return `<w:tbl>${tblPr}${tblGrid}${rowsXml}</w:tbl>`;
}
const emptyPara='<w:p><w:pPr></w:pPr></w:p>';
function buildDocumentXml(bodyXml){ return `<?xml version="1.0"?><w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>`; }
const STYLES_XML=`<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体" w:cs="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style></w:styles>`;
const CONTENT_TYPES=`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`;
const RELS=`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
const DOC_RELS=`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`;
const SETTINGS_XML=`<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
function packDocx(bodyXml){
  const zip=new MiniZip();
  zip.add('[Content_Types].xml',CONTENT_TYPES); zip.add('_rels/.rels',RELS);
  zip.add('word/document.xml',buildDocumentXml(bodyXml)); zip.add('word/styles.xml',STYLES_XML);
  zip.add('word/settings.xml',SETTINGS_XML); zip.add('word/_rels/document.xml.rels',DOC_RELS);
  return zip.toBlob();
}

// ══════════════ Markdown 解析（模板模式，未改动） ══════════════
function parseMarkdown(raw) {
  const lines=raw.split('\n');
  let minHash=5; for(const line of lines){ const m=line.match(/^(#{1,4})\s/); if(m) minHash=Math.min(minHash,m[1].length); }
  const offset=minHash<=4?minHash-1:0;
  const blocks=[]; let i=0;
  while(i<lines.length){
    const line=lines[i];
    if(/^\s*\|.+\|\s*$/.test(line)){
      const tl=[];
      while(i<lines.length){
        const cur=lines[i];
        if(!cur.trim()){ i++; continue; }
        if(!/^\s*\|.+\|\s*$/.test(cur)) break;
        tl.push(cur); i++;
      }
      const data=tl.filter(l=>!/^[\|\-\:\s]+$/.test(l.trim()));
      if(data.length){
        const rows=data.map(l=>{ let s=l.trim(); if(s.startsWith('|')) s=s.slice(1); if(s.endsWith('|')) s=s.slice(0,-1); return s.split('|').map(c=>c.trim()); });
        const nc=Math.max(...rows.map(r=>r.length)); rows.forEach(r=>{while(r.length<nc)r.push('');});
        blocks.push({type:'table',rows,nc});
      }
      continue;
    }
    if(!line.trim()){ i++; continue; }
    const hm=line.match(/^(#{1,4})\s+(.+)/);
    if(hm){
      let level=hm[1].length-offset; level=Math.max(1,Math.min(3,level));
      const runs=inlineM(hm[2].trim());
      blocks.push({type:'heading',level,text:runs.map(r=>r.text).join(''),runs});
      i++; continue;
    }
    const bm=line.match(/^[\s　]*([-*]|\d+[.)、])\s+(.*)/);
    const txt=bm?bm[2].trim():line.trim();
    if(txt){ const runs=inlineM(txt); blocks.push({type:'para',text:runs.map(r=>r.text).join(''),runs,isList:!!bm}); }
    i++;
  }
  return blocks;
}
function inlineM(text){
  const runs=[]; let s=text,m;
  while(s.length){
    if((m=s.match(/^\*\*(.*?)\*\*/s))){ runs.push({text:m[1],bold:true}); s=s.slice(m[0].length); continue; }
    if((m=s.match(/^\*(?!\*)(.*?)(?<!\*)\*/s))){ runs.push({text:m[1],italic:true}); s=s.slice(m[0].length); continue; }
    if((m=s.match(/^~~(.*?)~~/s))){ runs.push({text:m[1],strike:true}); s=s.slice(m[0].length); continue; }
    if((m=s.match(/^`([^`]+)`/))){ runs.push({text:m[1],code:true}); s=s.slice(m[0].length); continue; }
    const nx=s.search(/\*\*|\*|~~|`/); if(nx===-1){ runs.push({text:s}); break; }
    if(nx>0){ runs.push({text:s.slice(0,nx)}); s=s.slice(nx); } else{ runs.push({text:s[0]}); s=s.slice(1); }
  }
  return runs.filter(r=>r.text);
}

// ══════════════ 模板/字号 ══════════════
const SZ_LIST=[{n:'初号',p:42},{n:'小初',p:36},{n:'一号',p:26},{n:'小一',p:24},{n:'二号',p:22},{n:'小二',p:18},{n:'三号',p:16},{n:'小三',p:15},{n:'四号',p:14},{n:'小四',p:12},{n:'五号',p:10.5},{n:'小五',p:9}];
function ptToName(pt){ if(!pt) return ''; const f=SZ_LIST.find(s=>Math.abs(s.p-pt)<0.4); if(f) return f.n; return (Number.isInteger(pt)?pt:pt.toFixed(1))+'pt'; }
function nameToPt(val){ if(!val) return 0; const trimmed=val.replace(/（.*?）/g,'').trim(); const f=SZ_LIST.find(s=>s.n===trimmed); if(f) return f.p; return parseFloat(val)||0; }
const BUILT={
  guowen:{name:'公文类',icon:'📄',body:{font:'仿宋_GB2312',size:16,line:28,indent:true},h1:{font:'黑体',size:16,bold:true},h2:{font:'楷体',size:16,bold:true},h3:{font:'仿宋_GB2312',size:16,bold:true}},
  huibao:{name:'职场汇报类',icon:'📊',body:{font:'宋体',size:12,line:1.5,indent:true},h1:{font:'微软雅黑',size:14,bold:true},h2:{font:'微软雅黑',size:13,bold:true},h3:{font:'宋体',size:12,bold:true}},
  xueshu:{name:'学术报告类',icon:'📚',body:{font:'宋体',size:12,line:1.5,indent:true},h1:{font:'宋体',size:14,bold:true},h2:{font:'宋体',size:13,bold:true},h3:{font:'宋体',size:12,bold:true}},
};
const FONT_OPTS=['微软雅黑','宋体','黑体','仿宋_GB2312','仿宋','楷体_GB2312','楷体','微软雅黑 Light','方正小标宋简体','华文中宋','华文仿宋','华文楷体','华文宋体','华文黑体','新宋体','幼圆','隶书','华文细黑','华文新魏','中易宋体','Calibri','Calibri Light','Cambria','Times New Roman','Arial','Verdana','Tahoma'];
const SZ_OPTS=['初号（42pt）','小初（36pt）','一号（26pt）','小一（24pt）','二号（22pt）','小二（18pt）','三号（16pt）','小三（15pt）','四号（14pt）','小四（12pt）','五号（10.5pt）','小五（9pt）'];

function buildTplDocx(blocks,p){
  const line=p.bodyLine,indTwips=p.bodyIndent?Math.round(p.bodySize*40):0;
  const hcfg=[null,{font:p.h1Font,size:p.h1Size,bold:p.h1Bold},{font:p.h2Font,size:p.h2Size,bold:p.h2Bold},{font:p.h3Font,size:p.h3Size,bold:p.h3Bold}];
  const bodyXml=blocks.map(bl=>{
    if(bl.type==='table') return tableXml({rows:bl.rows,nc:bl.nc,font:p.bodyFont,size:p.bodySize,line})+emptyPara;
    if(bl.type==='heading'){ const lv=Math.min(bl.level,3); return headingXml({level:lv,runs:[{text:bl.text}],line,font:hcfg[lv].font,size:hcfg[lv].size,bold:hcfg[lv].bold,indTwips:0}); }
    return paraXml({runs:[{text:bl.text}],line,font:p.bodyFont,size:p.bodySize,indTwips,align:'both'});
  }).join('');
  return packDocx(bodyXml||emptyPara);
}
function buildCleanDocx(blocks){
  const FONT='微软雅黑',BSIZE=12,LINE=1.5; const hSz=[null,18,16,14];
  const bodyXml=blocks.map(bl=>{
    if(bl.type==='table') return tableXml({rows:bl.rows,nc:bl.nc,font:FONT,size:BSIZE,line:LINE})+emptyPara;
    if(bl.type==='heading'){ const sz=hSz[Math.min(bl.level,3)]; return headingXml({level:Math.min(bl.level,3),runs:bl.runs.map(r=>({...r,font:FONT,size:sz})),line:LINE,font:FONT,size:sz,indTwips:0}); }
    return paraXml({runs:bl.runs.map(r=>({...r,font:r.code?'Consolas':FONT,size:r.code?11:BSIZE})),line:LINE,font:FONT,size:BSIZE,indTwips:Math.round(BSIZE*40),align:'both'});
  }).join('');
  return packDocx(bodyXml||emptyPara);
}
function buildReflowDocx(blocks,p){
  const line=p.bodyLine,indTwips=p.bodyIndent?Math.round(p.bodySize*40):0;
  const hcfg=[null,{font:p.h1Font,size:p.h1Size,bold:p.h1Bold},{font:p.h2Font,size:p.h2Size,bold:p.h2Bold},{font:p.h3Font,size:p.h3Size,bold:p.h3Bold}];
  const bodyXml=blocks.map(bl=>{
    if(bl.type==='doctitle') return paraXml({runs:[{text:bl.text}],line,font:p.h1Font,size:p.h1Size+4,indTwips:0,align:'center'});
    if(bl.type==='table') return tableXml({rows:bl.rows,nc:bl.nc,font:p.bodyFont,size:p.bodySize,line})+emptyPara;
    if(bl.type==='heading'){ const lv=Math.min(bl.level,3); return headingXml({level:lv,runs:[{text:bl.text}],line,font:hcfg[lv].font,size:hcfg[lv].size,bold:hcfg[lv].bold,indTwips:0}); }
    return paraXml({runs:[{text:bl.text}],line,font:p.bodyFont,size:p.bodySize,indTwips,align:'both'});
  }).join('');
  return packDocx(bodyXml||emptyPara);
}

// ══════════════ 核心变量 ══════════════
let curMode='tpl', saved={}, cleanBlocks=null;
window._reflowBlocks=null; let reflowFileName='', reflowDocData=null;

// ── 文件上传/拖拽 ─────────────────────────
async function handleMdFile(ev){
  const file=ev.target.files[0]; if(!file) return; ev.target.value='';
  if(curMode==='reflow'&&file.name.endsWith('.docx')){ await loadReflowFile(file); }
  else{ await readFileIntoEditor(file); }
}
async function readFileIntoEditor(file){
  const ext=file.name.split('.').pop().toLowerCase();
  try{
    if(ext==='docx'){
      const ab=await file.arrayBuffer(); const xml=await readZipEntry(ab,'word/document.xml');
      const doc=new DOMParser().parseFromString(xml,'application/xml');
      const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const paras=doc.getElementsByTagNameNS(W,'p'); const lines=[];
      for(const p of paras){ const ts=p.getElementsByTagNameNS(W,'t'); const txt=Array.from(ts).map(r=>r.textContent).join('').trim(); if(txt) lines.push(txt); }
      document.getElementById('mdin').value=lines.join('\n\n');
    }else{ const text=await file.text(); document.getElementById('mdin').value=text; }
    showSt('ok',`✅ 已读取「${file.name}」`);
  }catch(e){ showSt('err',`❌ 读取失败：${e.message}`); }
}
function initDropzone(){
  const dz=document.getElementById('dropzone'), ov=document.getElementById('dropOverlay');
  if(!dz||!ov) return;
  dz.addEventListener('dragover',e=>{ e.preventDefault(); ov.style.display='flex'; });
  dz.addEventListener('dragleave',e=>{ if(!dz.contains(e.relatedTarget)) ov.style.display='none'; });
  dz.addEventListener('drop',e=>{ e.preventDefault(); ov.style.display='none'; const file=e.dataTransfer.files[0]; if(file) readFileIntoEditor(file); });
}

// ── 下拉组件 ──────────────────────────────
function cmbRender(id,opts,filter){
  const drop=document.getElementById(id+'_drop'); if(!drop) return;
  const lower=(filter||'').toLowerCase();
  const filtered=lower?opts.filter(o=>o.toLowerCase().includes(lower)):opts;
  drop.innerHTML=filtered.length?filtered.map(o=>`<div class="cmb-item" onmousedown="cmbPick('${id}','${o}')">${o}</div>`).join(''):'<div class="cmb-item" style="color:var(--muted)">无匹配结果</div>';
}
function cmbToggle(id){
  const drop=document.getElementById(id+'_drop'); if(!drop) return;
  const isOpen=drop.classList.contains('open'); cmbCloseAll();
  if(!isOpen){ const opts=id.includes('font')?FONT_OPTS:SZ_OPTS; cmbRender(id,opts,''); drop.classList.add('open'); document.getElementById(id).select(); }
}
function cmbPick(id,val){ document.getElementById(id).value=val; cmbCloseAll(); document.getElementById(id).dispatchEvent(new Event('change')); }
function cmbCloseAll(){ document.querySelectorAll('.cmb-drop').forEach(d=>d.classList.remove('open')); }
document.addEventListener('mousedown',e=>{ if(!e.target.closest('.cmb')) cmbCloseAll(); });
['b_font','b_size','h1_font','h1_size','h2_font','h2_size','h3_font','h3_size'].forEach(id=>{
  const el=document.getElementById(id); if(!el) return;
  el.addEventListener('input',()=>{ const drop=document.getElementById(id+'_drop'); if(!drop) return; const opts=id.includes('font')?FONT_OPTS:SZ_OPTS; cmbRender(id,opts,el.value); drop.classList.add('open'); });
  el.addEventListener('focus',()=>{ const drop=document.getElementById(id+'_drop'); if(!drop) return; if(!drop.classList.contains('open')){ const opts=id.includes('font')?FONT_OPTS:SZ_OPTS; cmbRender(id,opts,el.value); drop.classList.add('open'); } });
});

// ── 净化模式 HTML 解析（未改动） ─────────────
function parseHtmlToBlocks(html){
  const doc=new DOMParser().parseFromString(html,'text/html');
  const blocks=[];
  function getRuns(node){
    const runs=[];
    node.childNodes.forEach(c=>{
      if(c.nodeType===3) runs.push({text:c.textContent});
      else if(c.nodeType===1){
        const tag=c.tagName.toLowerCase();
        if(tag==='strong'||tag==='b'){ const sub=getRuns(c); sub.forEach(r=>r.bold=true); runs.push(...sub); }
        else if(tag==='em'||tag==='i'){ const sub=getRuns(c); sub.forEach(r=>r.italic=true); runs.push(...sub); }
        else if(tag==='del'||tag==='s'||tag==='strike'){ const sub=getRuns(c); sub.forEach(r=>r.strike=true); runs.push(...sub); }
        else if(tag==='code') runs.push({text:c.textContent,code:true});
        else runs.push(...getRuns(c));
      }
    });
    return runs;
  }
  function mergeRuns(runs){
    if(!runs.length) return [];
    const out=[]; let cur={...runs[0]};
    for(let i=1;i<runs.length;i++){
      const r=runs[i];
      if(!!r.bold===!!cur.bold&&!!r.italic===!!cur.italic&&!!r.strike===!!cur.strike&&!!r.code===!!cur.code) cur.text+=r.text;
      else{ out.push(cur); cur={...r}; }
    }
    out.push(cur); return out.filter(r=>r.text.trim());
  }
  function containsBlockElement(node){
    if(!node) return false;
    if(node.nodeType!==1) return false;
    const tag=node.tagName.toLowerCase();
    if(/^(table|h[1-6]|ul|ol|dl|blockquote|pre|hr)$/.test(tag)) return true;
    for(let i=0;i<node.childNodes.length;i++){ if(containsBlockElement(node.childNodes[i])) return true; }
    return false;
  }
  function walk(node,list){
    if(!node) return;
    if(node.nodeType===1){
      const tag=node.tagName.toLowerCase();
      if(/^h[1-6]$/.test(tag)){
        const level=parseInt(tag[1]); const runs=mergeRuns(getRuns(node)); const text=runs.map(r=>r.text).join('').trim(); if(text) list.push({type:'heading',level,text,runs});
      }else if(tag==='table'){
        const rows=[]; node.querySelectorAll('tr').forEach(tr=>{ const row=[]; tr.querySelectorAll('th,td').forEach(cell=>row.push(cell.textContent.trim())); if(row.length) rows.push(row); });
        if(rows.length){ const nc=Math.max(...rows.map(r=>r.length)); rows.forEach(r=>{while(r.length<nc)r.push('');}); list.push({type:'table',rows,nc}); }
      }else if(tag==='p'||tag==='div'||tag==='li'){
        if(containsBlockElement(node)){ node.childNodes.forEach(c=>walk(c,list)); }
        else{ const runs=mergeRuns(getRuns(node)); const text=runs.map(r=>r.text).join('').trim(); if(text) list.push({type:'para',text,runs,isList:tag==='li'}); }
      }else{ node.childNodes.forEach(c=>walk(c,list)); }
    }
  }
  walk(doc.body,blocks);
  return blocks;
}

// ── 粘贴拦截（未改动） ────────────────────────
let pasteFlag=false;
function initPasteListener(){
  const ta=document.getElementById('mdin');
  ta.addEventListener('paste',e=>{
    if(curMode!=='clean') return;
    e.preventDefault(); const cd=e.clipboardData; const html=cd.getData('text/html');
    if(!html){ ta.value=cd.getData('text/plain'); showSt('ok','已粘贴纯文本'); return; }
    try{
      const blocks=parseHtmlToBlocks(html); if(!blocks.length) throw new Error('未解析到块');
      cleanBlocks=blocks;
      const stats={chars:0,headings:0,paras:0,bolds:0,tables:0};
      blocks.forEach(b=>{
        if(b.type==='heading'){ stats.headings++; if(b.runs) b.runs.forEach(r=>{ stats.chars+=r.text.length; if(r.bold) stats.bolds++; }); }
        else if(b.type==='para'){ stats.paras++; if(b.runs) b.runs.forEach(r=>{ stats.chars+=r.text.length; if(r.bold) stats.bolds++; }); }
        else if(b.type==='table'){ stats.tables++; if(b.rows) b.rows.forEach(r=>r.forEach(c=>stats.chars+=c.length)); }
      });
      ta.value=cd.getData('text/plain');
      showSt('ok',`✅ 已读取 AI 网页内容（约 ${stats.chars} 字，${stats.headings} 个标题，${stats.paras} 段正文，${stats.tables} 个表格，${stats.bolds} 处加粗）`);
      document.getElementById('pTitle').textContent='✨ 已识别 AI 网页内容';
      document.getElementById('pStats').innerHTML=`约 ${stats.chars} 字<br>📌 标题 ${stats.headings} 个 · 正文 ${stats.paras} 段<br>🔤 加粗 ${stats.bolds} 处 · 表格 ${stats.tables} 个`;
      document.getElementById('previewOverlay').classList.add('show'); pasteFlag=true;
    }catch(err){ ta.value=cd.getData('text/plain'); showSt('err','❌ HTML 解析失败，已粘贴纯文本'); }
  });
  ta.addEventListener('input',()=>{ if(curMode!=='clean') return; if(pasteFlag){ pasteFlag=false; return; } cleanBlocks=null; });
}
function closePreviewOverlay(){ document.getElementById('previewOverlay').classList.remove('show'); }

// ══════════════ 重排模式 ══════════════
let curReflowSub='auto';
function setReflowSub(sub){
  curReflowSub=sub; ['auto','doubao'].forEach(k=>document.getElementById('rsb-'+k).classList.toggle('on',k===sub));
  document.getElementById('reflow-info-text').innerHTML=sub==='auto'?'<b>通用模式</b>：读取 Word 文件的大纲级别，按右侧模板参数重新排版。':'<b>🐮 豆包专用</b>：使用字号+文本量综合识别标题与正文，精准区分总标题、一级/二级/三级标题和正文。';
  if(reflowDocData&&reflowFileName) analyzeReflowDoc(reflowDocData,reflowFileName);
}
async function loadReflowFile(file){
  if(!file.name.endsWith('.docx')){ showSt('err','⚠️ 请上传 .docx 格式的 Word 文件'); return; }
  try{ 
    reflowDocData=await file.arrayBuffer(); 
    reflowFileName=file.name; 
    await analyzeReflowDoc(reflowDocData,file.name); 
  }
  catch(e){ showSt('err','❌ 读取文件失败：'+(e.message||'文件损坏或不兼容')); }
}
async function analyzeReflowDoc(ab,fname){
  const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  let xml;
  try { xml = await readZipEntry(ab,'word/document.xml'); } 
  catch(e) { showSt('err','❌ 无法解析 Word 文件，请确认是标准 .docx 格式'); return; }
  const doc = new DOMParser().parseFromString(xml,'application/xml');
  const body = doc.getElementsByTagNameNS(W,'body')[0];
  if(!body) { showSt('err','❌ 文档结构异常，缺少 body 元素'); return; }
  const blocks=[], paraInfos=[];
  const cellText=tc=>{ const ts=tc.getElementsByTagNameNS(W,'t'); return Array.from(ts).map(t=>t.textContent||'').join('').trim(); };
  function safeGetText(p) { try { const ts = p.getElementsByTagNameNS(W,'t'); return Array.from(ts).map(t=>t.textContent||'').join('').trim(); } catch(e) { return ''; } }
  function safeGetSz(p) { try { const rpr = p.getElementsByTagNameNS(W,'rPr')[0]; if(!rpr) return 0; const sz = rpr.getElementsByTagNameNS(W,'sz')[0] || rpr.getElementsByTagNameNS(W,'szCs')[0]; return sz ? parseFloat(sz.getAttributeNS(W,'val')||'0')/2 : 0; } catch(e) { return 0; } }
  for(let n=body.firstChild; n; n=n.nextSibling){
    if(n.nodeType!==1) continue;
    if(n.localName==='p'){
      const text = safeGetText(n); if(!text) continue;
      const sz = safeGetSz(n);
      paraInfos.push({text, sz}); blocks.push({type:'para', text});
    } else if(n.localName==='tbl'){
      try {
        const rows=[]; const trs=n.getElementsByTagNameNS(W,'tr');
        for(let tr of trs){ const row=[]; const tcs=tr.getElementsByTagNameNS(W,'tc'); for(let tc of tcs) row.push(cellText(tc)); if(row.length) rows.push(row); }
        if(rows.length){ const nc=Math.max(...rows.map(r=>r.length)); rows.forEach(r=>{while(r.length<nc) r.push('');}); blocks.push({type:'table', rows, nc}); }
      } catch(e) {}
    }
  }
  if(paraInfos.length===0){
    window._reflowBlocks=blocks;
    const fullText=blocks.map(b=>b.type==='table'?b.rows.map(r=>r.join(' ')).join('\n'):b.text||'').join('\n\n');
    document.getElementById('mdin').value=fullText;
    document.getElementById('reflowFileInfo').style.display='';
    document.getElementById('rfName').textContent=fname;
    document.getElementById('rfStat').textContent='识别到：表格 '+blocks.filter(b=>b.type==='table').length+' 个';
    showSt('ok',`✅ 已解析「${fname}」，点击下方按钮按模板排版输出`); return;
  }
  const sizeStats = new Map();
  for(const info of paraInfos) { if(!info.sz) continue; if(!sizeStats.has(info.sz)) sizeStats.set(info.sz, { count:0, totalChars:0 }); const s = sizeStats.get(info.sz); s.count++; s.totalChars += info.text.length; }
  let bodySize = null;
  if(sizeStats.size > 0) {
    const sortedSizes = [...sizeStats.entries()].sort((a,b) => b[1].count - a[1].count || b[1].totalChars - a[1].totalChars);
    bodySize = sortedSizes[0][0];
    const minSize = Math.min(...sizeStats.keys());
    if(minSize !== bodySize && (sizeStats.get(minSize).count >= 3 || minSize < bodySize - 1)) { bodySize = minSize; }
  }
  const headingCandidates = [...sizeStats.keys()].filter(s => s > bodySize).sort((a,b) => b-a);
  let hasTitle = false, titleSize = 0;
  if(paraInfos.length > 0 && headingCandidates.length > 0) {
    const firstSz = paraInfos[0].sz;
    const maxHeadingSize = headingCandidates[0];
    const maxCount = sizeStats.get(maxHeadingSize)?.count || 0;
    if(firstSz === maxHeadingSize && maxCount === 1) { hasTitle = true; titleSize = maxHeadingSize; const idx = headingCandidates.indexOf(titleSize); if(idx > -1) headingCandidates.splice(idx,1); }
  }
  const sizeToLevel = new Map();
  for(let i = 0; i < headingCandidates.length && i < 3; i++) { sizeToLevel.set(headingCandidates[i], i+1); }
  let paraIdx = 0;
  for(let i = 0; i < blocks.length; i++) {
    if(blocks[i].type === 'table') continue;
    const info = paraInfos[paraIdx++];
    if(hasTitle && info.sz === titleSize && info === paraInfos[0]) { blocks[i] = { type:'doctitle', text: info.text }; }
    else if(sizeToLevel.has(info.sz)) { const lv = sizeToLevel.get(info.sz); blocks[i] = { type:'heading', level: lv, text: info.text }; }
    else { blocks[i] = { type:'para', text: info.text }; }
  }
  window._reflowBlocks = blocks;
  const fullText = blocks.map(b => b.text || (b.rows?b.rows.map(r=>r.join(' ')).join('\n') : '')).join('\n\n');
  document.getElementById('mdin').value = fullText;
  const hCount = blocks.filter(b=>b.type==='heading').length;
  const pCount = blocks.filter(b=>b.type==='para').length;
  const dtCount = blocks.filter(b=>b.type==='doctitle').length;
  const tbCount = blocks.filter(b=>b.type==='table').length;
  document.getElementById('reflowFileInfo').style.display='';
  document.getElementById('rfName').textContent = fname;
  let statStr = `识别到：${dtCount?'总标题 1 个 · ':''}标题 ${hCount} 个 · 正文 ${pCount} 段 · 表格 ${tbCount} 个`;
  const levelCounts = {}; blocks.filter(b=>b.type==='heading').forEach(b=>{ levelCounts[b.level] = (levelCounts[b.level]||0)+1; });
  if(Object.keys(levelCounts).length) { statStr += ' ('; for(let lv=1; lv<=3; lv++) if(levelCounts[lv]) statStr += `H${lv}×${levelCounts[lv]} `; statStr = statStr.trim() + ')'; }
  document.getElementById('rfStat').textContent = statStr;
  showSt('ok', `✅ 已解析「${fname}」，点击下方按钮按模板排版输出。正文识别字号：${bodySize}pt`);
}

// ── 模式切换 ───────────────────────────────
function setMode(m){
  curMode=m; ['tpl','clean','reflow'].forEach(k=>document.getElementById('mb-'+k).classList.toggle('on',k===m));
  document.getElementById('sec-tpl').style.display=m==='tpl'?'':'none';
  document.getElementById('sec-clean').style.display=m==='clean'?'':'none';
  document.getElementById('sec-reflow').style.display=m==='reflow'?'':'none';
  const help={
    tpl:'<b>模板模式</b>：适合 Markdown 文本，按右侧模板参数生成 Word。',
    clean:'<b>净化模式</b>：适合从网页复制的富文本，尽量保留标题、加粗、列表和表格。',
    reflow:'<b>重排模式</b>：适合上传已有 .docx 后重新识别层级并套用当前模板。'
  };
  document.getElementById('modeHelp').innerHTML=help[m]||help.tpl;
  const cvBtn=document.getElementById('cvBtn'),pdfBtn=document.getElementById('pdfBtn');
  if(m==='reflow'){ cvBtn.innerHTML='⚡ 按模板重排并下载 Word'; pdfBtn.style.display='none'; }
  else{ cvBtn.innerHTML='⚡ 一键转换下载 Word'; pdfBtn.style.display=''; }
  if(m!=='clean'){ cleanBlocks=null; closePreviewOverlay(); }
}

const EXAMPLE_MD=`# 项目周报

这是一段正文，支持 **加粗**、*斜体*、~~删除线~~ 和 \`行内代码\`。

## 本周进展

- 完成 Markdown 内容整理
- 统一 Word 输出格式

| 项目 | 状态 | 备注 |
|---|---|---|
| 模板转换 | 完成 | 可直接下载 docx |
| 表格识别 | 正常 | 首行自动加粗 |

### 下周计划

继续完善文档排版细节。`;
let lastClearedState=null, clearUndoTimer=null;
function cloneState(v){ try{return structuredClone(v);}catch(_){ return v==null?v:JSON.parse(JSON.stringify(v)); } }
function insertExample(){
  const input=document.getElementById('mdin');
  if(input.value.trim()&&!confirm('当前输入区已有内容，是否用示例覆盖？')) return;
  input.value=EXAMPLE_MD; cleanBlocks=null; window._reflowBlocks=null; closePreviewOverlay();
  document.getElementById('reflowFileInfo').style.display='none';
  if(curMode==='reflow') setMode('tpl');
  input.focus(); showSt('ok','✅ 已插入示例内容，可直接转换测试');
}

// ── 清除 ───────────────────────────────────
function clearInput(){
  const input=document.getElementById('mdin'), info=document.getElementById('reflowFileInfo');
  if(!input.value&&!cleanBlocks&&!window._reflowBlocks){ showSt('err','⚠️ 当前没有可清空的内容'); return; }
  lastClearedState={
    value:input.value,
    cleanBlocks:cloneState(cleanBlocks),
    reflowBlocks:cloneState(window._reflowBlocks),
    reflowDisplay:info.style.display,
    rfName:document.getElementById('rfName').textContent,
    rfStat:document.getElementById('rfStat').textContent
  };
  input.value=''; cleanBlocks=null; window._reflowBlocks=null;
  closePreviewOverlay(); info.style.display='none';
  if(clearUndoTimer) clearTimeout(clearUndoTimer);
  showStHtml('ok','✅ 已清空内容。<button type="button" onclick="undoClear()">撤销</button>');
  clearUndoTimer=setTimeout(()=>{ lastClearedState=null; clearUndoTimer=null; const st=document.getElementById('stBar'); if(st.textContent.includes('已清空内容')) clearSt(); },10000);
}
function undoClear(){
  if(!lastClearedState){ showSt('err','⚠️ 没有可撤销的清空操作'); return; }
  document.getElementById('mdin').value=lastClearedState.value||'';
  cleanBlocks=cloneState(lastClearedState.cleanBlocks); window._reflowBlocks=cloneState(lastClearedState.reflowBlocks);
  const info=document.getElementById('reflowFileInfo'); info.style.display=lastClearedState.reflowDisplay||'none';
  document.getElementById('rfName').textContent=lastClearedState.rfName||'—';
  document.getElementById('rfStat').textContent=lastClearedState.rfStat||'—';
  lastClearedState=null; if(clearUndoTimer) clearTimeout(clearUndoTimer); clearUndoTimer=null;
  showSt('ok','✅ 已恢复清空前内容'); document.getElementById('mdin').focus();
}

// ── 主转换 ─────────────────────────────────
async function doConvert(){
  const btn=document.getElementById('cvBtn'); btn.disabled=true; btn.innerHTML='⏳ 转换中…'; clearSt();
  try{
    let blob;
    if(curMode==='tpl'){
      const raw=document.getElementById('mdin').value.trim(); if(!raw) throw new Error('请先粘贴 Markdown 内容');
      const blocks=parseMarkdown(raw); if(!blocks.length) throw new Error('未解析到有效内容');
      blob=buildTplDocx(blocks,readParams());
      const tb=blocks.filter(b=>b.type==='table').length,hd=blocks.filter(b=>b.type==='heading').length,pa=blocks.filter(b=>b.type==='para').length;
      showSt('ok',`✅ 转换完成！正文 ${pa} 段 · 标题 ${hd} 个 · 表格 ${tb} 个`);
    }else if(curMode==='clean'){
      if(!cleanBlocks){ const raw=document.getElementById('mdin').value.trim(); if(!raw) throw new Error('请先在净化模式下粘贴 AI 网页内容'); cleanBlocks=parseMarkdown(raw); }
      blob=buildCleanDocx(cleanBlocks);
      const tb=cleanBlocks.filter(b=>b.type==='table').length,hd=cleanBlocks.filter(b=>b.type==='heading').length,pa=cleanBlocks.filter(b=>b.type==='para').length;
      showSt('ok',`✅ 转换完成！正文 ${pa} 段 · 标题 ${hd} 个 · 表格 ${tb} 个`);
    }else if(curMode==='reflow'){
      if(!window._reflowBlocks||!window._reflowBlocks.length) throw new Error('请先在重排模式下上传 Word 文件，解析后左侧会显示内容');
      blob=buildReflowDocx(window._reflowBlocks,readParams());
      const hd=window._reflowBlocks.filter(b=>b.type==='heading').length,pa=window._reflowBlocks.filter(b=>b.type==='para').length,dt=window._reflowBlocks.filter(b=>b.type==='doctitle').length;
      showSt('ok',`✅ 重排完成！总标题 ${dt} 个 · 标题 ${hd} 个 · 正文 ${pa} 段`);
    }
    const url=URL.createObjectURL(blob); const a=document.createElement('a');
    a.href=url; a.download=`转换结果_${new Date().toLocaleDateString('zh-CN').replace(/\//g,'-')}.docx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(()=>URL.revokeObjectURL(url),2000);
  }catch(err){ console.error(err); showSt('err','❌ 转换失败：'+(err.message||'请查看控制台')); }
  finally{ btn.disabled=false; btn.innerHTML=curMode==='reflow'?'⚡ 按模板重排并下载 Word':'⚡ 一键转换下载 Word'; }
}

// ── PDF 导出 ──────────────────────────────
async function doPdf(){
  const raw=document.getElementById('mdin').value.trim(); if(!raw&&curMode!=='clean'&&curMode!=='reflow'){ showSt('err','⚠️ 请先粘贴内容'); return; }
  let blocks; if(curMode==='clean'&&cleanBlocks) blocks=cleanBlocks; else if(curMode==='reflow'&&window._reflowBlocks) blocks=window._reflowBlocks; else blocks=parseMarkdown(raw);
  if(!blocks.length){ showSt('err','⚠️ 未解析到有效内容'); return; }
  const p=curMode==='tpl'?readParams():null; const font=p?p.bodyFont:'微软雅黑',size=p?p.bodySize:12,line=p?p.bodyLine:1.5;
  const ind=p?(p.bodyIndent?`${size*2}pt`:'0'):'0'; const hFnt=p?[p.h1Font,p.h2Font,p.h3Font]:['微软雅黑','微软雅黑','微软雅黑']; const hSz=p?[p.h1Size,p.h2Size,p.h3Size]:[18,16,14];
  const lh=(!line||line<=0)?'1.5':(line>=6?`${line}pt`:String(line));
  let body='';
  for(const bl of blocks){
    if(bl.type==='table'){ const rows=bl.rows.map((row,ri)=>'<tr>'+row.map(c=>ri===0?`<th>${xe(c)}</th>`:`<td>${xe(c)}</td>`).join('')+'</tr>').join(''); body+=`<table>${rows}</table>`; }
    else if(bl.type==='heading'){ const lv=Math.min(bl.level,3); const txt=(curMode==='clean')?bl.runs.map(r=>{ let t=xe(r.text); if(r.bold) t=`<b>${t}</b>`; if(r.italic) t=`<i>${t}</i>`; return t; }).join(''):`<b>${xe(bl.text)}</b>`; body+=`<h${lv} style="font-family:'${hFnt[lv-1]}',sans-serif;font-size:${hSz[lv-1]}pt;line-height:${lh};margin:.5em 0 .2em">${txt}</h${lv}>`; }
    else{ const txt=(curMode==='clean')?bl.runs.map(r=>{ let t=xe(r.text); if(r.bold) t=`<b>${t}</b>`; if(r.italic) t=`<i>${t}</i>`; if(r.strike) t=`<del>${t}</del>`; if(r.code) t=`<code>${t}</code>`; return t; }).join(''):xe(bl.text); body+=`<p style="font-family:'${font}',sans-serif;font-size:${size}pt;line-height:${lh};text-indent:${ind};text-align:justify;margin:0 0 .15em">${txt}</p>`; }
  }
  const html=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>转换预览</title><style>body{margin:0;padding:2cm;font-family:'${font}',sans-serif}table{border-collapse:collapse;width:100%;margin:1em 0}th,td{border:1px solid #888;padding:4pt 6pt;font-family:'${font}',sans-serif;font-size:${size}pt}th{font-weight:bold;background:#f5f5f5}code{font-family:Consolas,monospace;font-size:90%;background:#f0f0f0;padding:1px 4px;border-radius:3px}h1,h2,h3{page-break-after:avoid}p{page-break-inside:avoid}@media print{body{padding:0}@page{margin:2cm}.no-print{display:none}}</style></head><body><div class="no-print" style="text-align:right;margin-bottom:1cm;color:#888;font-size:11pt"><button onclick="window.print()" style="padding:6px 18px;background:#F5A623;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:bold">🖨️ 打印 / 另存为 PDF</button>&nbsp;关闭此窗口后返回工具</div>${body}</body></html>`;
  const w=window.open('','_blank'); if(!w){ showSt('err','⚠️ 弹窗被拦截'); return; } w.document.write(html); w.document.close(); w.focus(); showSt('ok','✅ 预览已打开，点击「打印/另存为PDF」');
}

// ── 模板管理 ──────────────────────────────
function loadSaved(){ try{saved=JSON.parse(localStorage.getItem('laoliuAI_saved')||'{}');}catch(e){saved={};} }
function persistSaved(){ localStorage.setItem('laoliuAI_saved',JSON.stringify(saved)); }
function renderSaved(){ const slot=document.getElementById('savedSlot'); slot.innerHTML=''; Object.entries(saved).forEach(([id,t])=>{ const lb=document.createElement('label'); lb.className='ti'; lb.dataset.id=id; lb.innerHTML=`<input type="radio" name="tpl" value="${id}"><span class="ti-ico">⭐</span><div class="ti-info"><div class="ti-name">${t.name}</div><div class="ti-desc">${t.body.font} ${t.body.size}pt · ${t.body.line>=6?'固定'+t.body.line+'磅':t.body.line+'倍行距'}</div></div><div class="ti-dot"></div><button class="ti-del" onclick="delSaved(event,'${id}')">✕</button>`; slot.appendChild(lb); }); }
function delSaved(e,id){ e.preventDefault(); e.stopPropagation(); if(!confirm(`删除模板「${saved[id]?.name}」？`)) return; delete saved[id]; persistSaved(); renderSaved(); const cur=document.querySelector('input[name="tpl"]:checked'); if(!cur||cur.value===id){ document.querySelector('input[name="tpl"][value="guowen"]').checked=true; refreshTplUI(); fillParams(BUILT.guowen); } }
function openModal(){ if(Object.keys(saved).length>=5){ showSt('err','⚠️ 最多保存 5 个自定义模板'); return; } document.getElementById('moIn').value=''; document.getElementById('ov').classList.add('open'); setTimeout(()=>document.getElementById('moIn').focus(),50); }
function closeOv(e){ if(e.target===document.getElementById('ov')) closeOvDirect(); }
function closeOvDirect(){ document.getElementById('ov').classList.remove('open'); }
function confirmSave(){ const name=document.getElementById('moIn').value.trim(); if(!name){ document.getElementById('moIn').focus(); return; } const p=readParams(); const id='c_'+Date.now(); saved[id]={name,body:{font:p.bodyFont,size:p.bodySize,line:p.bodyLine,indent:p.bodyIndent},h1:{font:p.h1Font,size:p.h1Size,bold:p.h1Bold},h2:{font:p.h2Font,size:p.h2Size,bold:p.h2Bold},h3:{font:p.h3Font,size:p.h3Size,bold:p.h3Bold}}; persistSaved(); renderSaved(); closeOvDirect(); showSt('ok',`✅ 模板「${name}」已保存`); }
function readParams(){ const s=id=>document.getElementById(id).value.trim(); const sz=id=>nameToPt(s(id)); const n=id=>parseFloat(document.getElementById(id).value)||0; const ck=id=>document.getElementById(id).checked; const tId=document.querySelector('input[name="tpl"]:checked')?.value||'guowen'; const base=BUILT[tId]||saved[tId]||BUILT.guowen; return {bodyFont:s('b_font')||base.body.font,bodySize:sz('b_size')||base.body.size,bodyLine:n('b_line')||base.body.line,bodyIndent:ck('b_ind'),h1Font:s('h1_font')||base.h1.font,h1Size:sz('h1_size')||base.h1.size,h1Bold:ck('h1_bold'),h2Font:s('h2_font')||base.h2.font,h2Size:sz('h2_size')||base.h2.size,h2Bold:ck('h2_bold'),h3Font:s('h3_font')||base.h3.font,h3Size:sz('h3_size')||base.h3.size,h3Bold:ck('h3_bold')}; }
function fillParams(t){ const b=t.body; $v('b_font',b.font||''); $v('b_size',ptToName(b.size)); $v('b_line',b.line||''); document.getElementById('b_ind').checked=!!b.indent; $v('h1_font',t.h1?.font||''); $v('h1_size',ptToName(t.h1?.size)); $v('h2_font',t.h2?.font||''); $v('h2_size',ptToName(t.h2?.size)); $v('h3_font',t.h3?.font||''); $v('h3_size',ptToName(t.h3?.size)); document.getElementById('h1_bold').checked=t.h1?.bold!==false; document.getElementById('h2_bold').checked=t.h2?.bold!==false; document.getElementById('h3_bold').checked=t.h3?.bold!==false; }
function $v(id,val){ document.getElementById(id).value=val; }

// ── 设置导入/导出 ───────────────────────────
const SETTINGS_SCHEMA='markdown-to-word-settings', SETTINGS_VERSION=1;
const SETTINGS_KEYS=['laoliuAI_saved','laoliuAI_lastTpl','laoliuAI_theme','laoliuAI_rpWidth','laoliuAI_noNotice'];
function collectSettingsPayload(){
  const settings={}; SETTINGS_KEYS.forEach(k=>{ settings[k]=localStorage.getItem(k); });
  return {schema:SETTINGS_SCHEMA,version:SETTINGS_VERSION,exportedAt:new Date().toISOString(),origin:location.origin,settings};
}
function normalizeImportedSettings(payload){
  if(!payload||typeof payload!=='object') throw new Error('不是有效的设置文件');
  if(payload.schema&&payload.schema!==SETTINGS_SCHEMA) throw new Error('设置文件类型不匹配');
  const src=payload.settings&&typeof payload.settings==='object'?payload.settings:payload;
  const out={};
  for(const key of SETTINGS_KEYS){
    if(!Object.prototype.hasOwnProperty.call(src,key)) continue;
    const raw=src[key];
    if(raw===null||raw===undefined||raw===''){ out[key]=null; continue; }
    const val=String(raw);
    if(key==='laoliuAI_saved'){
      const parsed=JSON.parse(val);
      if(!parsed||typeof parsed!=='object'||Array.isArray(parsed)) throw new Error('自定义模板数据格式不正确');
      out[key]=JSON.stringify(parsed);
    }else if(key==='laoliuAI_theme'){
      if(!['dark','light','paper'].includes(val)) throw new Error('主题设置不正确');
      out[key]=val;
    }else if(key==='laoliuAI_rpWidth'){
      const w=parseInt(val,10);
      if(!Number.isFinite(w)||w<220||w>560) throw new Error('侧栏宽度超出范围');
      out[key]=String(w);
    }else if(key==='laoliuAI_noNotice'){
      if(val!=='1') throw new Error('提示关闭状态不正确');
      out[key]=val;
    }else{
      if(val.length>128) throw new Error('模板选择记录过长');
      out[key]=val;
    }
  }
  if(!Object.keys(out).length) throw new Error('设置文件中没有可导入的项目');
  return out;
}
function applySettingsPayload(payload){
  const settings=normalizeImportedSettings(payload);
  SETTINGS_KEYS.forEach(k=>{
    if(!Object.prototype.hasOwnProperty.call(settings,k)) return;
    settings[k]===null?localStorage.removeItem(k):localStorage.setItem(k,settings[k]);
  });
  loadSaved(); renderSaved(); restoreTheme(); restoreTpl(); noticeInit();
  const rp=document.getElementById('rpanel'), savedW=parseInt(localStorage.getItem('laoliuAI_rpWidth'),10);
  if(rp&&savedW&&savedW>=220&&savedW<=560) rp.style.width=savedW+'px';
  return settings;
}
function exportSettings(){
  try{
    const blob=new Blob([JSON.stringify(collectSettingsPayload(),null,2)],{type:'application/json;charset=utf-8'});
    const url=URL.createObjectURL(blob), a=document.createElement('a');
    a.href=url; a.download=`markdown-word-settings_${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(()=>URL.revokeObjectURL(url),2000);
    showSt('ok','✅ 设置已导出为 JSON 文件');
  }catch(err){ console.error(err); showSt('err','❌ 导出失败：'+(err.message||'请查看控制台')); }
}
async function handleSettingsImport(ev){
  const file=ev.target.files[0]; if(!file) return;
  try{
    const payload=JSON.parse(await file.text());
    const imported=applySettingsPayload(payload);
    showSt('ok',`✅ 设置导入成功，已恢复 ${Object.keys(imported).length} 项偏好`);
  }catch(err){ console.error(err); showSt('err','❌ 导入失败：'+(err.message||'请确认是本工具导出的 JSON 设置文件')); }
  finally{ ev.target.value=''; }
}

// ══════════════ 上传 docx 识别参数（完整实现） ══════════════
async function handleUpload(ev){
  const file = ev.target.files[0];
  if (!file) return;
  const el = document.getElementById('upStatus');
  el.style.display = 'block';
  el.style.color = 'var(--muted)';
  el.textContent = '⏳ 正在识别参数…';
  try {
    const ab = await file.arrayBuffer();
    const xml = await readZipEntry(ab, 'word/styles.xml');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    // 建立 styleId → 元素映射
    const styleMap = {};
    const docEl = doc.documentElement;
    for (let i = 0; i < docEl.childNodes.length; i++) {
      const s = docEl.childNodes[i];
      if (s.nodeType !== 1 || s.localName !== 'style') continue;
      const sid = s.getAttributeNS(W, 'styleId') || s.getAttribute('w:styleId') || '';
      if (sid) styleMap[sid] = s;
    }

    // 辅助函数
    function ch(el, name) { if (!el) return null; for (let i = 0; i < el.childNodes.length; i++) { const c = el.childNodes[i]; if (c.nodeType === 1 && c.localName === name) return c; } return null; }
    function desc(el, name) { if (!el) return null; for (let i = 0; i < el.childNodes.length; i++) { const c = el.childNodes[i]; if (c.nodeType !== 1) continue; if (c.localName === name) return c; const r = desc(c, name); if (r) return r; } return null; }
    function ga2(el, attr) { if (!el) return ''; return el.getAttributeNS(W, attr) || el.getAttribute('w:' + attr) || el.getAttribute(attr) || ''; }
    function dFont(st) { const rpr = ch(st, 'rPr'); if (!rpr) return ''; const f = ch(rpr, 'rFonts'); if (!f) return ''; return ga2(f, 'eastAsia') || ga2(f, 'ascii') || ''; }
    function dSize(st) { const rpr = ch(st, 'rPr'); if (rpr) { const z = ch(rpr, 'sz') || ch(rpr, 'szCs'); if (z) { const v = ga2(z, 'val'); if (v) return parseFloat(v) / 2; } } const allRpr = desc(st, 'rPr'); if (allRpr) { const z = ch(allRpr, 'sz') || ch(allRpr, 'szCs'); if (z) { const v = ga2(z, 'val'); if (v) return parseFloat(v) / 2; } } return 0; }
    function dBold(st) { const rpr = ch(st, 'rPr'); if (!rpr) return undefined; return ch(rpr, 'b') !== null; }
    function dLine(st) { const ppr = ch(st, 'pPr'); if (!ppr) return 0; const sp = ch(ppr, 'spacing'); if (!sp) return 0; const v = ga2(sp, 'line'), rule = ga2(sp, 'lineRule'); if (!v) return 0; return rule === 'exact' ? parseFloat(v) / 20 : Math.round(parseFloat(v) / 240 * 100) / 100; }
    function dInd(st) { const ppr = ch(st, 'pPr'); if (!ppr) return undefined; const ind = ch(ppr, 'ind'); if (!ind) return undefined; const fl = ga2(ind, 'firstLine'); return fl ? parseInt(fl) > 0 : undefined; }
    function resolve(st, getter, depth = 0) { if (!st || depth > 6) return null; const val = getter(st); if (val) return val; const bo = ch(st, 'basedOn'); if (!bo) return null; const parentId = ga2(bo, 'val'); return resolve(styleMap[parentId] || null, getter, depth + 1); }
    function ddFont() { const dd = desc(doc.documentElement, 'docDefaults'); if (!dd) return ''; const rprDef = desc(dd, 'rPr'); if (!rprDef) return ''; const f = ch(rprDef, 'rFonts'); return f ? (ga2(f, 'eastAsia') || ga2(f, 'ascii') || '') : ''; }
    function ddSize() { const dd = desc(doc.documentElement, 'docDefaults'); if (!dd) return 0; const rprDef = desc(dd, 'rPr'); if (!rprDef) return 0; const z = ch(rprDef, 'sz') || ch(rprDef, 'szCs'); return z ? parseFloat(ga2(z, 'val')) / 2 : 0; }
    function ddLine() { const dd = desc(doc.documentElement, 'docDefaults'); if (!dd) return 0; const pprDef = desc(dd, 'pPr'); if (!pprDef) return 0; const sp = ch(pprDef, 'spacing'); if (!sp) return 0; const v = ga2(sp, 'line'), rule = ga2(sp, 'lineRule'); if (!v) return 0; return rule === 'exact' ? parseFloat(v) / 20 : Math.round(parseFloat(v) / 240 * 100) / 100; }

    const getStyle = (stdId, stdName) => {
      if (styleMap[stdId]) return styleMap[stdId];
      const lower = stdName.toLowerCase();
      for (const s of Object.values(styleMap)) {
        const nm = ch(s, 'name');
        if (nm) { const v = ga2(nm, 'val') || nm.getAttribute('w:val') || ''; if (v.toLowerCase() === lower) return s; }
      }
      return null;
    };

    const nm = getStyle('Normal', 'normal'),
          h1 = getStyle('Heading1', 'heading 1'),
          h2 = getStyle('Heading2', 'heading 2'),
          h3 = getStyle('Heading3', 'heading 3');

    // 清空所有字段
    ['b_font','b_size','b_line','h1_font','h1_size','h2_font','h2_size','h3_font','h3_size'].forEach(id => $v(id, ''));
    let hits = 0;
    function setF(id, val) { if (!val && val !== 0) return; $v(id, String(val)); hits++; }

    // 正文
    const bFont = resolve(nm, dFont) || ddFont();
    const bSize = resolve(nm, dSize) || ddSize();
    const bLine = resolve(nm, dLine) || ddLine();
    setF('b_font', bFont);
    setF('b_size', ptToName(bSize));
    if (bLine) setF('b_line', bLine);
    const ind = dInd(nm);
    if (ind !== undefined) { document.getElementById('b_ind').checked = ind; hits++; }

    // 标题
    function setH(fid, sid, bid, st) {
      const font = dFont(st) || resolve(st, dFont);
      const size = dSize(st) || resolve(st, dSize) || ddSize();
      const bold = dBold(st);
      if (font) setF(fid, font);
      if (size) setF(sid, ptToName(size));
      if (bold !== undefined) { document.getElementById(bid).checked = bold; hits++; }
    }
    setH('h1_font', 'h1_size', 'h1_bold', h1);
    setH('h2_font', 'h2_size', 'h2_bold', h2);
    setH('h3_font', 'h3_size', 'h3_bold', h3);

    el.style.color = 'var(--accent)';
    el.textContent = `✅ 已读取「${file.name}」，识别到 ${hits} 项参数，未识别的请手动补充`;
  } catch (err) {
    el.style.color = 'var(--err)';
    el.textContent = '❌ 识别失败：' + (err.message || '请确认是标准 .docx 文件');
  }
  ev.target.value = '';
}

// ── 状态条 / 主题 / 初始化 ──────────────────
function showSt(type,msg){ const el=document.getElementById('stBar'); el.className='st '+type; el.textContent=msg; }
function showStHtml(type,html){ const el=document.getElementById('stBar'); el.className='st '+type; el.innerHTML=html; }
function clearSt(){ const el=document.getElementById('stBar'); el.className='st'; el.textContent=''; }
function setTheme(t){ document.documentElement.setAttribute('data-theme',t); document.querySelectorAll('.t-btn').forEach(b=>b.classList.toggle('on',b.dataset.t===t)); localStorage.setItem('laoliuAI_theme',t); }
function restoreTheme(){ setTheme(localStorage.getItem('laoliuAI_theme')||'dark'); }
function refreshTplUI(){ document.querySelectorAll('#tplList .ti').forEach(el=>{ const r=el.querySelector('input[type=radio]'); el.classList.toggle('on',r?.checked); }); }
function restoreTpl(){ const last=localStorage.getItem('laoliuAI_lastTpl')||'guowen'; const r=document.querySelector(`input[name="tpl"][value="${last}"]`); if(r){ r.checked=true; refreshTplUI(); } const t=BUILT[last]||saved[last]; if(t) fillParams(t); }
function noticeInit(){ const bar=document.getElementById('noticeBar'); if(!bar) return; bar.classList.toggle('hidden',localStorage.getItem('laoliuAI_noNotice')==='1'); }
function noticeClose(forever){ const bar=document.getElementById('noticeBar'); if(!bar) return; if(forever) localStorage.setItem('laoliuAI_noNotice','1'); bar.style.maxHeight='0'; bar.style.opacity='0'; bar.style.margin='0'; bar.style.overflow='hidden'; setTimeout(()=>bar.classList.add('hidden'),320); }

function init(){
  loadSaved(); renderSaved(); restoreTheme(); restoreTpl(); initDropzone(); initPasteListener(); noticeInit();
  document.getElementById('tplList').addEventListener('input',e=>{ if(e.target.name!=='tpl') return; refreshTplUI(); const t=BUILT[e.target.value]||saved[e.target.value]; if(t) fillParams(t); localStorage.setItem('laoliuAI_lastTpl',e.target.value); });
  document.getElementById('moIn').addEventListener('keydown',e=>{ if(e.key==='Enter') confirmSave(); if(e.key==='Escape') closeOvDirect(); });
  const divider=document.getElementById('divider'),rp=document.getElementById('rpanel'); let isDragging=false,startX=0,startW=0;
  const savedW=parseInt(localStorage.getItem('laoliuAI_rpWidth')); if(savedW&&savedW>=220&&savedW<=560) rp.style.width=savedW+'px';
  divider.addEventListener('mousedown',e=>{ isDragging=true; startX=e.clientX; startW=rp.offsetWidth; divider.classList.add('dragging'); document.body.style.cursor='col-resize'; document.body.style.userSelect='none'; e.preventDefault(); });
  document.addEventListener('mousemove',e=>{ if(!isDragging) return; const delta=startX-e.clientX; const newW=Math.max(220,Math.min(560,startW+delta)); rp.style.width=newW+'px'; });
  document.addEventListener('mouseup',()=>{ if(!isDragging) return; isDragging=false; divider.classList.remove('dragging'); document.body.style.cursor=''; document.body.style.userSelect=''; localStorage.setItem('laoliuAI_rpWidth',rp.offsetWidth); });
}
init();
