/* Shipflow browser-side workbook editor. No external scripts or services. */
(() => {
  'use strict';
  const td = new TextDecoder();
  const te = new TextEncoder();
  const maxCompressedBytes = 2_500_000;
  const maxUncompressedBytes = 14_000_000;
  const maxRows = 1200, maxCols = 35;
  const $ = s => document.querySelector(s);
  const alpha = i => { let out = ''; for (i++; i; i=Math.floor((i-1)/26)) out=String.fromCharCode(65+(i-1)%26)+out; return out; };
  const columnNumber = s => [...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1;
  const state = { data:[], fileName:'shipment_workbook.xlsx', sheetName:'Shipment Data', selected:[0,0], dirty:false, page:0, pageSize:45, usedRows:0, lastError:'' };
  const xmlEscape = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;').replace(/[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]/g,'');
  const parseXML = value => { if (/<!DOCTYPE|<!ENTITY/i.test(value)) throw Error('Unsupported XML declarations in workbook.'); const doc=new DOMParser().parseFromString(value,'application/xml'); if(doc.getElementsByTagName('parsererror').length) throw Error('This workbook contains invalid XML.'); return doc; };
  const nodes = (el,local) => [...el.getElementsByTagName('*')].filter(e=>e.localName===local);
  const first = (el,local) => nodes(el,local)[0];
  const readU16=(v,o)=>v.getUint16(o,true), readU32=(v,o)=>v.getUint32(o,true);
  async function extractZip(buffer, desired) {
    if(buffer.byteLength > maxCompressedBytes) throw Error('Maximum upload size is 2.5 MB.');
    const v=new DataView(buffer), bytes=new Uint8Array(buffer);
    let eocd=-1;
    for(let i=buffer.byteLength-22;i>=Math.max(0,buffer.byteLength-65557);i--) if(readU32(v,i)===0x06054b50){eocd=i;break;}
    if(eocd<0) throw Error('Not a valid Excel .xlsx ZIP file.');
    const n=Math.min(readU16(v,eocd+10),1000), cd=readU32(v,eocd+16);let pos=cd,total=0;
    const files=new Map();
    for(let i=0;i<n;i++){
      if(pos+46>bytes.length || readU32(v,pos)!==0x02014b50) throw Error('Invalid ZIP directory.');
      const method=readU16(v,pos+10), compressed=readU32(v,pos+20), plain=readU32(v,pos+24);
      const nameLen=readU16(v,pos+28), extra=readU16(v,pos+30), comment=readU16(v,pos+32), offset=readU32(v,pos+42);
      const name=td.decode(bytes.slice(pos+46,pos+46+nameLen));pos+=46+nameLen+extra+comment;
      if(!desired(name))continue;
      total+=plain;if(total>maxUncompressedBytes) throw Error('Workbook content is too large for the demo.');
      if(offset+30>bytes.length || readU32(v,offset)!==0x04034b50) throw Error('Invalid ZIP file entry.');
      const dataPos=offset+30+readU16(v,offset+26)+readU16(v,offset+28);
      if(dataPos+compressed>bytes.length)throw Error('Incomplete workbook file.');
      const compressedBytes=bytes.slice(dataPos,dataPos+compressed);
      let out;
      if(method===0)out=compressedBytes;
      else if(method===8){
        if(typeof DecompressionStream==='undefined')throw Error('Please use a recent Chrome, Edge or Firefox browser to import Excel files.');
        try {const stream=new Blob([compressedBytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));out=new Uint8Array(await new Response(stream).arrayBuffer());}
        catch {throw Error('Could not decompress this workbook. Please use a recent browser.');}
      } else throw Error('Unsupported Excel compression method.');
      if(out.length>maxUncompressedBytes || out.length!==plain)throw Error('Invalid or oversized workbook content.');
      files.set(name,td.decode(out));
    }
    return files;
  }
  async function readXlsx(file) {
    const contents=await extractZip(await file.arrayBuffer(), n=>n==='xl/workbook.xml'||n==='xl/_rels/workbook.xml.rels'||n==='xl/sharedStrings.xml'||/^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    if(!contents.has('xl/workbook.xml'))throw Error('Excel workbook metadata is missing.');
    const workbook=parseXML(contents.get('xl/workbook.xml'));
    const sheet=nodes(workbook,'sheet')[0];if(!sheet)throw Error('Excel workbook contains no worksheet.');
    const sheetName=sheet.getAttribute('name')||'Shipment Data';
    const relId=sheet.getAttribute('r:id')||sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    let sheetPath='xl/worksheets/sheet1.xml';
    if(contents.has('xl/_rels/workbook.xml.rels') && relId) {
      const rels=parseXML(contents.get('xl/_rels/workbook.xml.rels'));
      const rel=nodes(rels,'Relationship').find(r=>r.getAttribute('Id')===relId);
      if(rel){let t=rel.getAttribute('Target')||'';t=t.replace(/^\//,'');if(!t.startsWith('xl/'))t='xl/'+t.replace(/^\.\//,'');if(t.startsWith('xl/worksheets/'))sheetPath=t;}
    }
    const sheetText=contents.get(sheetPath)||contents.get('xl/worksheets/sheet1.xml');
    if(!sheetText)throw Error('The first worksheet is missing or unsupported.');
    let strings=[];
    if(contents.has('xl/sharedStrings.xml')){
      const ss=parseXML(contents.get('xl/sharedStrings.xml'));
      strings=nodes(ss,'si').map(si=>nodes(si,'t').map(t=>t.textContent).join(''));
    }
    const doc=parseXML(sheetText), out=[];
    let rowCount=0;
    for(const row of nodes(doc,'row')) {
      let r=Number(row.getAttribute('r')||(++rowCount))-1;
      if(!Number.isInteger(r)||r<0||r>=maxRows)continue;
      rowCount=r+1;
      const cells=[...row.children].filter(el=>el.localName==='c');let prev=-1;
      for(const cell of cells){
        const ref=cell.getAttribute('r');let col=ref?columnNumber((ref.match(/[A-Z]+/i)||[''])[0].toUpperCase()):prev+1;prev=col;
        if(col<0||col>=maxCols)continue;
        const type=cell.getAttribute('t')||'';
        let value=type==='inlineStr'?nodes(cell,'t').map(x=>x.textContent).join(''):(first(cell,'v')?.textContent||'');
        if(type==='s')value=strings[Number(value)]??'';
        if(type==='b')value=value==='1'?'TRUE':'FALSE';
        if(!out[r])out[r]=[];
        out[r][col]=String(value);
      }
    }
    return {rows:out, sheetName};
  }
  function parseCSV(text) {
    text=text.replace(/^\uFEFF/,'');const rows=[];let row=[],cell='',quoted=false;
    for(let i=0;i<text.length;i++){
      const ch=text[i];
      if(ch==='"'){if(quoted&&text[i+1]==='"'){cell+='"';i++;}else quoted=!quoted;}
      else if(ch===','&&!quoted){row.push(cell);cell='';}
      else if((ch==='\n'||ch==='\r')&&!quoted){if(ch==='\r'&&text[i+1]==='\n')i++;row.push(cell);rows.push(row);row=[];cell='';}
      else cell+=ch;
    }
    row.push(cell);if(row.length>1||row[0]!==''||!rows.length)rows.push(row);
    return rows.slice(0,maxRows).map(r=>r.slice(0,maxCols));
  }
  function defaultRows() {return [
    ['Shipment ID','Origin','Destination','Customer','Product','Quantity','Weight (kg)','Status','Dispatch date','Notes'],
    ['SHP-1001','London, UK','Dubai, UAE','Atlas Trading','Electronic parts','24','112','Ready','2026-10-23','Priority shipment'],
    ['SHP-1002','Manchester, UK','Toronto, CA','Maple Logistics','Machine parts','12','340','In transit','2026-10-24','Air freight'],
    ['SHP-1003','Berlin, DE','Singapore, SG','Nova Supply','Textiles','60','185','Pending','2026-10-25','Awaiting documents'],
    ['SHP-1004','Paris, FR','New York, US','Global Retail','Accessories','42','96','Ready','2026-10-27','Standard delivery']
  ];}
  const dims=()=>({rows:Math.max(state.data.length,12),cols:Math.max(10,Math.min(maxCols,Math.max(...state.data.map(r=>r?.length||0),0)))});
  function render() {
    const {rows,cols}=dims(), pages=Math.max(1,Math.ceil(rows/state.pageSize));state.page=Math.max(0,Math.min(state.page,pages-1));
    const head=$('#sheetGrid thead'),body=$('#sheetGrid tbody');if(!head||!body)return;
    head.replaceChildren();body.replaceChildren();
    const hr=document.createElement('tr');const corner=document.createElement('th');corner.className='sheet-corner';corner.textContent='#';hr.append(corner);
    for(let col=0;col<cols;col++){const th=document.createElement('th');th.textContent=alpha(col);hr.append(th);}head.append(hr);
    const frag=document.createDocumentFragment(),begin=state.page*state.pageSize,end=Math.min(rows,begin+state.pageSize);
    for(let r=begin;r<end;r++){
      const tr=document.createElement('tr'),num=document.createElement('th');num.className='row-number';num.textContent=r+1;tr.append(num);
      for(let c=0;c<cols;c++){
        const cell=document.createElement('td');cell.dataset.row=r;cell.dataset.col=c;cell.tabIndex=0;cell.textContent=state.data[r]?.[c]??'';
        if(r===0)cell.classList.add('table-header-row');
        if(r===state.selected[0]&&c===state.selected[1])cell.classList.add('selected-cell');
        tr.append(cell);
      }frag.append(tr);
    }body.append(frag);
    $('#sheetFileName').textContent=state.fileName+(state.dirty?' • Unsaved changes':'');
    $('#sheetTabName').textContent=state.sheetName;
    $('#sheetCounter').textContent=`${rows} rows · ${cols} columns`;
    $('#sheetPage').textContent=`Page ${state.page+1} of ${pages}`;
    $('#sheetPrevious').disabled=state.page===0;$('#sheetNext').disabled=state.page>=pages-1;
    syncSelection();
  }
  function syncSelection() {
    const [r,c]=state.selected;$('#cellAddress').textContent=`${alpha(c)}${r+1}`;
    $('#formulaInput').value=state.data[r]?.[c]??'';
    $('#sheetGrid').querySelectorAll('td.selected-cell').forEach(e=>e.classList.remove('selected-cell'));
    const current=$(`#sheetGrid td[data-row="${r}"][data-col="${c}"]`);if(current)current.classList.add('selected-cell');
  }
  function select(r,c,focus=false){const {rows,cols}=dims();r=Math.max(0,Math.min(rows-1,r));c=Math.max(0,Math.min(cols-1,c));state.selected=[r,c];
    const pg=Math.floor(r/state.pageSize);if(pg!==state.page){state.page=pg;render();}
    else syncSelection();
    if(focus)$(`#sheetGrid td[data-row="${r}"][data-col="${c}"]`)?.focus();
  }
  function edit(r,c,value) {if(r>=maxRows||c>=maxCols)return; if(!state.data[r])state.data[r]=[];state.data[r][c]=String(value).slice(0,5000);state.dirty=true;
    const dom=$(`#sheetGrid td[data-row="${r}"][data-col="${c}"]`);if(dom&&dom.textContent!==state.data[r][c])dom.textContent=state.data[r][c];
    $('#sheetFileName').textContent=state.fileName+' • Unsaved changes';if(state.selected[0]===r&&state.selected[1]===c)$('#formulaInput').value=state.data[r][c];
  }
  function reset() {state.data=defaultRows();state.fileName='shipment_workbook.xlsx';state.sheetName='Shipment Data';state.selected=[0,0];state.page=0;state.dirty=false;render();}
  async function load(file){if(!file)return;const csv=/\.csv$/i.test(file.name),xlsx=/\.xlsx$/i.test(file.name);
    if(!csv&&!xlsx)throw Error('Please upload an .xlsx or .csv file.');
    if(file.size>maxCompressedBytes)throw Error('File must be smaller than 2.5 MB.');
    let parsed;
    if(csv) parsed={rows:parseCSV(await file.text()),sheetName:'Shipment Data'};
    else parsed=await readXlsx(file);
    if(!parsed.rows.some(r=>r?.some(v=>v!=='')))throw Error('This workbook has no data in its first sheet.');
    state.data=parsed.rows;state.fileName=file.name.replace(/\.csv$/i,'.xlsx');state.sheetName=parsed.sheetName;
    state.selected=[0,0];state.page=0;state.dirty=false;render();
    return getPreview();
  }
  function getPreview(){const rows=state.data.filter(Boolean).filter(row=>row.some(x=>x!==undefined&&x!==''));const [header,...rest]=rows;
    return {sheetName:state.sheetName,rowCount:rest.length,headers:(header||[]).slice(0,20),preview:rest.slice(0,5).map(r=>(header||[]).slice(0,20).map((_,i)=>r[i]??''))};
  }
  // Minimal, valid XLSX exporter. Exports values from the first sheet; styles and formulas are not preserved.
  const makeSheetXML=()=>{
    const xml=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>','<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'];
    for(let r=0;r<state.data.length;r++){
      const row=state.data[r]||[];if(!row.some(v=>v!==undefined&&v!==''))continue;
      xml.push(`<row r="${r+1}">`);
      for(let c=0;c<Math.min(row.length,maxCols);c++){
        const v=row[c];if(v===undefined||v==='')continue;
        const address=alpha(c)+(r+1),str=String(v);
        if(/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(str)&&!(str.length>1&&str[0]==='0')&&str.length<=14)xml.push(`<c r="${address}"><v>${str}</v></c>`);
        else xml.push(`<c r="${address}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(str)}</t></is></c>`);
      }xml.push('</row>');
    }xml.push('</sheetData></worksheet>');return xml.join('');
  };
  const crcTable=new Uint32Array(256).map((_,n)=>{let c=n;for(let i=0;i<8;i++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);return c>>>0;});
  function crc32(bytes){let c=0xFFFFFFFF;for(const x of bytes)c=crcTable[(c^x)&255]^(c>>>8);return(c^0xFFFFFFFF)>>>0;}
  function makeZip(entries) {
    const out=[],central=[];let offset=0;
    const u16=n=>[n&255,(n>>>8)&255],u32=n=>[n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255];
    for(const [name,content] of entries){
      const filename=te.encode(name),file=te.encode(content),crc=crc32(file),n=file.length;
      const local=new Uint8Array([...u32(0x04034b50),...u16(20),...u16(0),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(n),...u32(n),...u16(filename.length),...u16(0)]);
      out.push(local,filename,file);
      central.push(new Uint8Array([...u32(0x02014b50),...u16(20),...u16(20),...u16(0),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(n),...u32(n),...u16(filename.length),...u16(0),...u16(0),...u16(0),...u16(0),...u32(0),...u32(offset)]),filename);
      offset+=local.length+filename.length+file.length;
    }
    const cdLength=central.reduce((n,x)=>n+x.length,0);
    const end=new Uint8Array([...u32(0x06054b50),...u16(0),...u16(0),...u16(entries.length),...u16(entries.length),...u32(cdLength),...u32(offset),...u16(0)]);
    return new Blob([...out,...central,end],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  function toBlob() {
    const name=xmlEscape(state.sheetName.slice(0,31));
    return makeZip([
      ['[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
      ['_rels/.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml',`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${name}" sheetId="1" r:id="rId1"/></sheets></workbook>`],
      ['xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
      ['xl/styles.xml','<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'],
      ['xl/worksheets/sheet1.xml',makeSheetXML()]
    ]);
  }
  function save(){const blob=toBlob(),href=URL.createObjectURL(blob),a=document.createElement('a');a.href=href;a.download=state.fileName.replace(/\.csv$/i,'.xlsx');document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(href),2500);state.dirty=false;render();}
  function openEdit(cell) {
    if(!cell||cell.firstElementChild?.tagName==='INPUT')return;
    const r=Number(cell.dataset.row),c=Number(cell.dataset.col);select(r,c);
    const field=document.createElement('input');field.className='grid-edit';field.type='text';field.value=state.data[r]?.[c]??'';cell.textContent='';cell.append(field);field.focus();field.select();
    const finish=move=>{if(!field.isConnected)return;edit(r,c,field.value);cell.textContent=state.data[r]?.[c]??'';if(move)select(r+move[0],c+move[1],true);};
    field.addEventListener('blur',()=>finish(null));field.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();finish([1,0]);}else if(e.key==='Tab'){e.preventDefault();finish([0,e.shiftKey?-1:1]);}else if(e.key==='Escape'){cell.textContent=state.data[r]?.[c]??'';cell.focus();}});
  }
  function pasteCells(text){const [r,c]=state.selected,lines=text.replace(/\r/g,'').split('\n');if(lines.at(-1)==='')lines.pop();
    lines.forEach((line,i)=>line.split('\t').forEach((v,j)=>edit(r+i,c+j,v)));render();select(r,c,true);
  }
  function init(){
    $('#sheetOpen').addEventListener('click',()=>$('#sheetInput').click());
    $('#sheetInput').addEventListener('change',async e=>{try{await load(e.target.files[0]);window.shipflowNotify?.('Workbook opened. You can edit cells directly.');}catch(err){window.shipflowNotify?.(err.message,true);}finally{e.target.value='';}});
    $('#sheetSample').addEventListener('click',()=>{reset();window.shipflowNotify?.('Sample shipment workbook opened.');});
    $('#sheetNew').addEventListener('click',()=>{state.data=[['Shipment ID','Origin','Destination','Customer','Product','Quantity','Weight (kg)','Status']];state.fileName='new_shipments.xlsx';state.sheetName='Shipments';state.page=0;state.dirty=true;render();});
    $('#sheetAddRow').addEventListener('click',()=>{if(state.data.length>=maxRows)return window.shipflowNotify?.('Demo limit is 1,200 rows.',true);state.data.push([]);state.dirty=true;select(state.data.length-1,0);render();});
    $('#sheetAddCol').addEventListener('click',()=>{const c=dims().cols;if(c>=maxCols)return window.shipflowNotify?.('Demo limit is 35 columns.',true);state.data.forEach(row=>{if(row)row[c]='';});state.dirty=true;render();select(0,c);});
    $('#sheetDownload').addEventListener('click',save);
    $('#sheetPrevious').addEventListener('click',()=>{state.page--;render();});$('#sheetNext').addEventListener('click',()=>{state.page++;render();});
    $('#sheetSend').addEventListener('click',async()=>{try{const blob=toBlob();const file=new File([blob],state.fileName.replace(/\.csv$/i,'.xlsx'),{type:blob.type});window.shipflowSelectFile?.(file);window.shipflowNavigate?.('new');window.shipflowNotify?.('Edited workbook added to the dispatch form.');}catch(err){window.shipflowNotify?.(err.message,true);}});
    $('#formulaInput').addEventListener('change',e=>{edit(...state.selected,e.target.value);render();});
    $('#sheetGrid').addEventListener('click',e=>{const cell=e.target.closest('td[data-row]');if(cell)select(+cell.dataset.row,+cell.dataset.col);});
    $('#sheetGrid').addEventListener('dblclick',e=>openEdit(e.target.closest('td[data-row]')));
    $('#sheetGrid').addEventListener('keydown',e=>{const cell=e.target.closest('td[data-row]');if(!cell||e.target.tagName==='INPUT')return;const r=+cell.dataset.row,c=+cell.dataset.col;
      const moves={ArrowRight:[0,1],ArrowLeft:[0,-1],ArrowDown:[1,0],ArrowUp:[-1,0],Enter:[1,0],Tab:[0,e.shiftKey?-1:1]};
      if(moves[e.key]){e.preventDefault();select(r+moves[e.key][0],c+moves[e.key][1],true);}else if(e.key==='F2'){e.preventDefault();openEdit(cell);}else if(e.key==='Delete'||e.key==='Backspace'){e.preventDefault();edit(r,c,'');}else if(e.key.length===1&&!e.ctrlKey&&!e.metaKey&&!e.altKey){openEdit(cell);const f=cell.querySelector('input');if(f){f.value=e.key;f.setSelectionRange(1,1);}}
    });
    $('#sheetGrid').addEventListener('paste',e=>{if(e.target.closest('input'))return;e.preventDefault();pasteCells(e.clipboardData.getData('text/plain'));});
    reset();
  }
  window.ShipflowWorkbook={init,load,getPreview,toBlob,save,get data(){return state.data},get fileName(){return state.fileName},get dirty(){return state.dirty}};
})();
