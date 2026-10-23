/* Static, read-only preview. Document text is never interpreted as executable markup. */
(async()=>{
  'use strict';
  if(new URLSearchParams(location.search).get('embedded')==='1'){document.body.classList.add('embedded');}
  const root=document.getElementById('preview-content'),status=document.getElementById('preview-status');
  const el=(tag,cls,text)=>{const x=document.createElement(tag);if(cls)x.className=cls;if(text!==undefined)x.textContent=String(text);return x;};
  const btn=(text,fn)=>{const b=el('button','',text);b.type='button';b.onclick=fn;return b;};
  const note=text=>root.append(el('p','preview-note',text));
  const base=new URL('./',location.href), raw=new URLSearchParams(location.search).get('file');
  let url,path;
  const ready=()=>{document.body.dataset.ready='true';};
  try {
    if(!raw)throw Error('缺少文件路径，请从审阅台点击文件。');
    url=new URL(raw,base);
    if(url.origin!==base.origin||!url.pathname.startsWith(base.pathname)||/^(data|javascript):/.test(url.protocol))throw Error('此预览仅支持本审阅站内的文件。');
    path=decodeURIComponent(url.pathname.slice(base.pathname.length));
    fetch('input-origin-index.json',{cache:'no-store'}).then(r=>r.ok?r.json():{}).then(index=>{const origin=index[path];if(!origin)return;const banner=el('div','origin-banner '+origin.origin_kind);banner.append(el('strong','',origin.origin_label+' · '+origin.origin_detail));if(origin.source_url){const link=el('a','',origin.source_url);link.href=origin.source_url;link.target='_blank';link.rel='noopener';banner.append(link);}root.before(banner);}).catch(()=>{});
    const name=path.split('/').pop(),ext=name.includes('.')?name.split('.').pop().toLowerCase():'';
    const domain=new URLSearchParams(location.search).get('domain');
    const labels={GO:'通用办公',FIN:'金融',LAW:'法律',DA:'数据分析',GA:'游戏'};
    document.title=(labels[domain]?labels[domain]+' · ':'')+name+' · 在线预览';document.getElementById('preview-title').textContent=name;document.getElementById('preview-path').textContent=(labels[domain]?labels[domain]+' · '+domain+' / ':'')+path;
    if(labels[domain])document.querySelector('.preview-brand').href='./?domain='+domain+'#inputs';
    const download=document.getElementById('preview-download');download.href=url.href;download.download=name;
    async function json(href){const target=new URL(href,base);target.searchParams.set('v','colleague-20261022');const r=await fetch(target);if(!r.ok)throw Error('文件加载失败（HTTP '+r.status+'）');return r.json();}
    function table(rows,header=false,total){
      const controls=el('div','preview-controls'),search=el('input','preview-search');search.type='search';search.placeholder='筛选当前预览中的内容';search.setAttribute('aria-label','筛选表格');controls.append(search);
      const wrap=el('div','preview-table-wrap'),pagination=el('div','preview-pagination');let page=0,filtered=rows;
      const headers=header?rows[0]:null,body=header?rows.slice(1):rows;
      function render(){
        const term=search.value.toLowerCase();filtered=term?body.filter(row=>row.some(c=>String(c??'').toLowerCase().includes(term))):body;
        const pages=Math.max(1,Math.ceil(filtered.length/50));page=Math.min(page,pages-1);
        const t=el('table');const head=el('thead'),hr=el('tr');hr.append(el('th','','行'));
        const width=Math.max(headers?.length||0,...filtered.slice(page*50,page*50+50).map(r=>r.length),1);
        for(let i=0;i<width;i++)hr.append(el('th','',headers?headers[i]??'':String(i+1)));head.append(hr);t.append(head);
        const tbody=el('tbody');filtered.slice(page*50,page*50+50).forEach((row,i)=>{const tr=el('tr');tr.append(el('th','',page*50+i+(header?2:1)));for(let k=0;k<width;k++){const cell=el('td','',row[k]??'');if(String(row[k]??'').startsWith('='))cell.title='原工作簿公式（未在预览中重算）';tr.append(cell);}tbody.append(tr);});t.append(tbody);wrap.replaceChildren(t);
        const prev=btn('← 上一页',()=>{page--;render();}),next=btn('下一页 →',()=>{page++;render();});prev.disabled=page===0;next.disabled=page>=pages-1;
        pagination.replaceChildren(prev,el('span','','第 '+(page+1)+' / '+pages+' 页 · '+filtered.length+' 行'+(total?' · 原表 '+total+' 行':'')),next);
      }
      search.oninput=()=>{page=0;render();};root.append(controls,wrap,pagination);render();
    }
    function csv(text,partial){
      const rows=[];let row=[],value='',quoted=false;
      for(let i=0;i<text.length;i++){
        const c=text[i];
        if(c==='"'){if(quoted&&text[i+1]==='"'){value+='"';i++;}else quoted=!quoted;}
        else if(c===','&&!quoted){row.push(value);value='';}
        else if((c==='\n'||c==='\r')&&!quoted){if(c==='\r'&&text[i+1]==='\n')i++;row.push(value);if(row.some(x=>x!==''))rows.push(row);row=[];value='';}
        else value+=c;
      }
      if(!partial&&(value!==''||row.length)){row.push(value);rows.push(row);}return rows;
    }
    async function text(limit){
      let r=await fetch(url,{headers:{Range:'bytes=0-'+limit}});
      if(r.status===416)r=await fetch(url);
      if(!r.ok)throw Error('文件加载失败（HTTP '+r.status+'）');
      const reader=r.body.getReader(),decoder=new TextDecoder('utf-8');let result='',bytes=0,partial=false;
      while(true){const {done,value}=await reader.read();if(done){result+=decoder.decode();break;}const remaining=limit-bytes;if(value.length>remaining){result+=decoder.decode(value.subarray(0,Math.max(0,remaining)),{stream:true});partial=true;await reader.cancel();break;}result+=decoder.decode(value,{stream:true});bytes+=value.length;}
      return {text:result,partial};
    }
    async function renderText(limit=2*1024*1024){
      status.textContent='正在载入…';const data=await text(limit);root.replaceChildren();
      if(data.partial){const n=el('div','preview-note','大文件当前预览前 '+Math.round(limit/1048576)+' MB，原件完整可下载。');n.append(btn('继续加载更多',()=>renderText(Math.min(limit*4,64*1048576)).catch(fail)));root.append(n);}
      if(ext==='csv'){table(csv(data.text,data.partial),!/^dota2(?:Train|Test)\.csv$/.test(name));}
      else if(ext==='html'||ext==='htm'){
        note('原始 HTML 在隔离视图内展示，脚本不运行。');const wrap=el('div','preview-html'),frame=el('iframe');frame.title=name;frame.setAttribute('sandbox','');frame.srcdoc=data.text;wrap.append(frame);root.append(wrap);
      }else{
        let content=data.text;
        if(ext==='json'&&!data.partial&&content.length<2*1048576){try{content=JSON.stringify(JSON.parse(content),null,2);}catch{note('文件未能解析为完整 JSON，按原文展示。');}}
        const lines=content.split('\n'), controls=el('div','preview-controls'),search=el('input','preview-search');search.type='search';search.placeholder='查找文本（仅已载入部分）';search.setAttribute('aria-label','查找文本');controls.append(search);
        const wrap=el('div','preview-code'),pre=el('pre');wrap.append(pre);const pager=el('div','preview-pagination');let page=0;
        function show(){const term=search.value.toLowerCase();const filtered=lines.map((line,i)=>({line,i})).filter(x=>!term||x.line.toLowerCase().includes(term));const pages=Math.max(1,Math.ceil(filtered.length/250));page=Math.min(page,pages-1);pre.replaceChildren();for(const item of filtered.slice(page*250,page*250+250)){const row=el('div','preview-line');row.append(el('span','preview-line-number',item.i+1),el('span','',item.line.length>50000?item.line.slice(0,50000)+' …（单行预览限 50,000 字符；完整内容见原件）':item.line||' '));pre.append(row);}const prev=btn('← 上一页',()=>{page--;show();}),next=btn('下一页 →',()=>{page++;show();});prev.disabled=page===0;next.disabled=page===pages-1;pager.replaceChildren(prev,el('span','','第 '+(page+1)+' / '+pages+' 页 · '+filtered.length+' 行'),next);}
        controls.append(btn('切换换行',()=>wrap.classList.toggle('wrap')));search.oninput=()=>{page=0;show();};root.append(controls,wrap,pager);show();
      }
      status.textContent=data.partial?'部分内容预览 · 原文件保持完整':'文件内容预览';ready();
    }
    function fail(error){status.textContent='预览未完成';root.replaceChildren(el('p','preview-error',error.message+'。可使用右上角按钮下载原文件。'));ready();}
    if(['docx','xlsx','pptx'].includes(ext)){
      const index=await json('file-preview-index.json');if(!index[path])throw Error('此文档尚无预览转换结果');const data=await json(index[path]);
      if(data.type==='docx'){
        note('文档内容预览：按原文顺序展示段落与表格；原始排版请下载 DOCX 查看。');const doc=el('article','preview-doc');
        for(const block of data.blocks){if(block.type==='table'){const wrap=el('div','preview-table-wrap'),t=el('table');for(const row of block.rows){const tr=el('tr');row.forEach(c=>tr.append(el('td','',c??'')));t.append(tr);}wrap.append(t);doc.append(wrap);}else{const heading=String(block.style||'').match(/(?:Heading|标题)\s*(\d)/i);doc.append(el(heading?'h'+Math.min(3,Number(heading[1])):'p','',block.text||'\u00a0'));}}root.append(doc);
      }else if(data.type==='xlsx'){
        const tabs=el('div','preview-tabs');tabs.setAttribute('role','tablist');root.append(tabs);const body=el('div');root.append(body);let serial=0;
        async function select(sheet,index){const current=++serial;[...tabs.children].forEach((b,i)=>b.setAttribute('aria-selected',String(i===index)));body.textContent='正在加载工作表…';const values=await json(sheet.href);if(current!==serial)return;body.replaceChildren();const keep=[...root.children];for(const child of keep)if(child!==tabs&&child!==body)child.remove();if(values.truncated)note('此表共 '+values.total_rows+' 行、'+values.total_columns+' 列；在线预览前 '+values.rows.length+' 行、最多 '+(values.column_limit||50)+' 列。完整数据请下载原文件。');else note('工作表：'+sheet.name+' · '+values.total_rows+' 行、'+values.total_columns+' 列。公式按原式显示，未重算。');table(values.rows,false,values.total_rows);body.remove();}
        data.sheets.forEach((sheet,i)=>tabs.append(btn(sheet.name,()=>select(sheet,i).catch(fail))));if(data.sheets.length)await select(data.sheets[0],0);else note('工作簿没有可预览的工作表。');
      }
      status.textContent='文档内容预览 · 原文件可下载';ready();
    }else if(ext==='pdf'){
      const response=await fetch(new URL('pdf-preview-index.json?v=colleague-20261022',base));
      const index=response.ok?await response.json():{};
      const document=index[path];
      if(document?.pages?.length){
        const pager=el('div','preview-pagination'),wrap=el('div','preview-pdf-pages'),image=el('img');
        image.style.cssText='display:block;width:100%;height:auto;background:white;box-shadow:0 2px 12px #17263d15';
        wrap.append(image);root.append(pager,wrap);let page=0;
        const select=el('select');select.setAttribute('aria-label','选择 PDF 页码');
        document.pages.forEach((p,i)=>{const option=el('option','','第 '+(i+1)+' 页');option.value=i;select.append(option);});
        function show(){image.src=new URL(document.pages[page].href,base).href;image.alt=name+' · 第 '+(page+1)+' 页';select.value=String(page);const prev=btn('← 上一页',()=>{page--;show();}),next=btn('下一页 →',()=>{page++;show();});prev.disabled=page===0;next.disabled=page===document.pages.length-1;pager.replaceChildren(prev,select,el('span','','共 '+document.pages.length+' 页'),next);}
        select.onchange=()=>{page=Number(select.value);show();};show();
        note('原 PDF 全部 '+document.pages.length+' 页均可预览；翻页或选择页码查看，原文件保持完整。');
        status.textContent='PDF 逐页在线预览';
      }else{
        const wrap=el('div','preview-pdf'),frame=el('iframe');frame.title=name;frame.src=url.href;wrap.append(frame);root.append(wrap);note('若浏览器未显示 PDF，可使用右上角“下载原文件”查看。');status.textContent='PDF 在线预览';
      }
      ready();
    }else if(['png','jpg','jpeg','gif','webp','svg'].includes(ext)){
      const img=el('img');img.src=url.href;img.alt=name;img.style.maxWidth='100%';root.append(img);status.textContent='图片预览';ready();
    }else if(ext==='zip'){
      note('此文件为完整审阅压缩包，请使用右上角按钮下载；其中的文档可在审阅台逐一在线预览。');status.textContent='压缩包';ready();
    }else await renderText();
  }catch(error){status.textContent='预览未完成';root.replaceChildren(el('p','preview-error',error.message));ready();}
})();
