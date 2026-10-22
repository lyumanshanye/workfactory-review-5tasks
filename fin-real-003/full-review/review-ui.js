/* Presentation only. Original task text, evidence and file bytes remain intact. */
(() => {
  'use strict';
  const legacy = document.querySelector('main');
  if (!legacy) return;
  const domains = ['GO','FIN','LAW','DA','GA'];
  const labels = {GO:'行政与采购',FIN:'金融分析',LAW:'法律与合规',DA:'数据分析',GA:'游戏分析'};
  const state = {domain:'GO', view:'colleague', source:'current'};
  let inputs = {}, bundle = {};
  const make = (tag, cls, text) => {
    const el = document.createElement(tag);
    if (cls) el.className = cls;
    if (text !== undefined) el.textContent = text;
    return el;
  };
  const a = (href, text, cls) => {
    const el = make('a', cls, text); el.href = href; previewLink(el); return el;
  };
  function previewLink(link) {
    const raw=link.getAttribute('href');
    if(!raw||raw.startsWith('#')||/^(?:https?:|mailto:|javascript:|data:)/i.test(raw)||raw.includes('file-preview.html'))return;
    if(raw.endsWith('.zip')) {link.dataset.originalArchive='true';return;}
    link.dataset.originalHref=raw;
    link.href='file-preview.html?file='+encodeURIComponent(raw);
    link.target='_blank';link.rel='noopener';link.removeAttribute('download');
  }
  const button = (text, fn, cls) => {
    const el = make('button', cls, text); el.type='button'; el.addEventListener('click',fn); return el;
  };
  const size = bytes => bytes >= 1048576 ? (bytes/1048576).toFixed(1)+' MB' : bytes >= 1024 ? (bytes/1024).toFixed(1)+' KB' : bytes+' B';
  document.body.classList.add('review-app');
  document.title = '工作类任务审阅台 · 五领域题目与文件';
  const app = make('div','review-root');
  app.innerHTML = `
    <header class="review-topbar"><a class="review-brand" href="#model-inputs"><span class="review-dot"></span>WORKFACTORY <span>人工审阅</span></a><span class="review-badge">5 个领域 · 独立预览</span></header>
    <main class="review-shell">
      <header class="review-hero"><p class="review-eyebrow">TASK REVIEW / 工作类任务</p><h1 class="review-title">五领域工作任务 · 审阅台</h1><p class="review-subtitle">按领域查看完整题面、输入附件、参考答案、Rubric 与模型产出。</p></header>
      <section class="review-contract" id="model-inputs" aria-labelledby="input-title">
        <div class="review-section-head"><div><p class="review-eyebrow">01 / 模型输入范围</p><h2 id="input-title">根据规范，给解题模型哪些文件？</h2></div><span class="review-badge">交付规范 §1.2–1.3</span></div>
        <div class="review-visibility">
          <div class="review-access-card"><span class="review-count">01</span><h3>题面作为 Prompt</h3><code>instruction.md</code><p>完整任务书直接作为提示词交给解题 Agent。</p></div>
          <div class="review-access-card"><span class="review-count">02</span><h3>随题原始附件</h3><code>/app/input_files/</code><p>题包的 <code>environment/input_files/</code> 挂载到此目录，只读。下方列出本题实际文件。</p></div>
          <div class="review-access-card"><span class="review-count">03</span><h3>运行环境与交付目录</h3><code>/app/output/</code><p>模型在此写交付物。依赖预装；如提供 Skill，须在题面声明其镜像内路径。</p></div>
        </div>
        <p class="review-privacy"><strong>裁判专用：</strong><code>rubrics.json</code>、<code>tests/</code>、<code>solution/</code> 和参考答案不提供给解题模型；<code>task.toml</code> 文件不进入解题容器。本页包含这些审阅材料，不能整页交给盲解模型。</p>
        <div class="review-domain-tabs" role="tablist" aria-label="选择任务领域"></div>
        <div class="review-input-card"><div class="review-input-head"><div><span class="review-eyebrow" id="input-task-id"></span><h3 id="input-files-title">实际输入文件</h3></div><div class="review-actions" id="input-actions"></div></div><div class="review-source-tabs" role="group" aria-label="选择题目来源"></div><div id="input-manifest"></div></div>
        <details class="review-spec"><summary>查看规范依据与路径说明</summary><p>依据《外发版-评测题包交付规范（rewardkit）20260913版本-1009修改》§1.2、§1.3。模型工作目录为 <code>/app</code>；唯一交付目录为 <code>/app/output/</code>。<code>solution/</code> 仅用于 Oracle，<code>tests/</code> 仅在评分阶段上传；模型仅能看到 <code>[environment].env</code> 的环境变量值，不能读取配置文件本身。</p><p>“同事原题”保留原文件命名，<code>TASK_INSTRUCTIONS.txt</code> 作为题面单列，业务附件另列；<code>state.json</code>、<code>RUBRICS.json</code>、<code>difficulty.json</code> 属于审阅证据。</p></details>
      </section>
      <div class="review-workspace"><aside class="review-sidebar"><p class="review-eyebrow">02 / 审阅内容</p><nav aria-label="审阅模块"></nav><p class="review-muted">顶部选择领域，下方内容同步切换。</p></aside><div class="review-panels"></div></div>
      <footer class="review-footer">WORKFACTORY · 独立审阅预览 · 原始材料与验收状态以各题证据为准</footer>
    </main>`;
  legacy.before(app);
  const panels = app.querySelector('.review-panels');
  const panelMap = {};
  const panelDefs = [['colleague','题目对比','同事原题与当前题，按领域并排阅读。'],['all-tasks','完整题面','当前任务书保持原文，可直接下载。'],['task-files','本题全部文件','题面、输入、参考、rubric、产出和证据，逐题对应。'],['fin-review','金融完整证据','原页面的金融专项材料与评分记录，保留完整内容。'],['pipeline','验证概况','原页面记录的四题验证快照。']];
  for (const [id,title,desc] of panelDefs) {
    const panel=make('section','review-panel'); panel.id='view-'+id;
    panel.append(make('h2','review-panel-title',title),make('p','review-muted',desc));
    panels.append(panel); panelMap[id]=panel;
    const link=a('#'+id,title); link.dataset.view=id;
    app.querySelector('.review-sidebar nav').append(link);
  }
  const compareBox=make('div','review-comparisons');
  panelMap.colleague.append(compareBox);
  const compareNote = legacy.querySelector('#colleague')?.nextElementSibling;
  if (compareNote?.tagName==='P') { compareNote.classList.add('review-context'); panelMap.colleague.insertBefore(compareNote,compareBox); }
  for (const domain of domains) {
    const section=legacy.querySelector('#compare-'+domain);
    if (!section) continue;
    section.classList.add('review-compare'); section.dataset.domain=domain;
    const comparison=section.querySelector('.comparison');
    if (comparison) {
      comparison.classList.add('review-comparison');
      for (const card of [...comparison.children]) {
        card.classList.add('review-comparison-card');
        const head=card.querySelector(':scope > h4');
        if (head) head.classList.add('review-comparison-head');
        const body=make('div','review-comparison-body');
        for (const node of [...card.childNodes]) if(node!==head) body.append(node);
        card.append(body);
      }
    }
    compareBox.append(section);
  }
  for (const task of [...legacy.querySelectorAll('article[id^="task-"]')]) {
    task.classList.add('review-task');
    task.dataset.domain=task.id.split('-')[1];
    if (task.id==='task-FIN-REAL-002') {
      const d=make('details','review-history'); d.dataset.domain='FIN';
      d.append(make('summary','', '历史对照：FIN-REAL-002'),task); panelMap['all-tasks'].append(d);
    } else panelMap['all-tasks'].append(task);
  }
  const status=legacy.querySelector('.status');
  if(status) {status.classList.add('review-status');panelMap['fin-review'].append(status);}
  for (const id of ['instruction','inputs','reference','rubric','outputs','scores','pipeline']) {
    const heading=legacy.querySelector(':scope > h2#'+id);
    if (!heading) continue;
    const wrapper=make('section','review-evidence-section');
    let node=heading;
    while(node) {
      if(node!==heading && node.tagName==='H2' && node.id) break;
      const next=node.nextSibling; wrapper.append(node); node=next;
    }
    (id==='pipeline' ? panelMap.pipeline : panelMap['fin-review']).append(wrapper);
  }
  // Keep all remaining source content available without a duplicate page header.
  const remainder=make('details','review-source-remainder');
  remainder.append(make('summary','', '原页面导航与附加说明'));
  for(const node of [...legacy.childNodes]) remainder.append(node);
  panelMap['fin-review'].append(remainder);
  legacy.remove();
  app.querySelectorAll('a[href]').forEach(previewLink);
  app.querySelectorAll('a[data-original-archive]').forEach(link=>link.remove());
  app.querySelectorAll('iframe').forEach(frame=>frame.loading='lazy');
  const filePanel=make('div','review-file-groups'); panelMap['task-files'].append(filePanel);
  for(const domain of domains) {
    const b=button(domain+' · '+labels[domain],()=>{
      state.domain=domain;state.source='current';render();
      history.replaceState(null,'','#'+(state.view==='colleague'?'compare-'+domain:state.view));
    });
    b.role='tab';b.dataset.domain=domain;app.querySelector('.review-domain-tabs').append(b);
  }
  function fileRow(file, folderHint=false) {
    const row=make('div','review-file-row');
    const ext=file.name.split('.').pop().toUpperCase();
    row.append(make('span','review-file-kind',ext.length<8?ext:'FILE'));
    const center=make('div','review-file-label');
    center.append(file.href?a(file.href,file.name):make('span','',file.name));
    if(file.sha256) {
      const meta=make('span','review-file-meta','SHA256 '+file.sha256.slice(0,12)+'…');meta.title=file.sha256;center.append(meta);
    }
    row.append(center,make('span','review-file-meta',size(file.bytes)));
    if(file.href) {const open=a(file.href,'预览 ↗','review-file-open');row.append(open);}
    return row;
  }
  function fileList(files) {
    const list=make('div','review-file-list');
    if(files.length>15) {
      const search=make('input','review-file-search');search.type='search';search.placeholder='按文件名或路径筛选 '+files.length+' 个文件';search.setAttribute('aria-label','筛选文件');
      list.append(search);
      const rows=make('div','review-file-results');files.forEach(f=>rows.append(fileRow(f)));
      search.addEventListener('input',()=>{
        const term=search.value.trim().toLowerCase();[...rows.children].forEach((r,i)=>r.hidden=!files[i].name.toLowerCase().includes(term));
      }); list.append(rows);
    } else files.forEach(f=>list.append(fileRow(f)));
    return list;
  }
  function renderInputs() {
    const source=inputs[state.domain]; if(!source) return;
    const item=state.source==='colleague'?source.colleague:source;
    const choices=app.querySelector('.review-source-tabs');choices.replaceChildren();
    for(const [key,label] of [['current','当前题'],['colleague','同事原题']]) {
      const b=button(label,()=>{state.source=key;renderInputs();});b.disabled=key==='colleague'&&!source.colleague;
      b.setAttribute('aria-pressed',String(state.source===key));choices.append(b);
    }
    app.querySelector('#input-task-id').textContent=item?.task||source.task;
    app.querySelector('#input-files-title').textContent='解题模型输入 · '+(item?.files.length||0)+' 个附件';
    const actions=app.querySelector('#input-actions');actions.replaceChildren();
    if(item?.instruction) actions.append(a(item.instruction,'预览题面 ↗','review-button'));
    const root=app.querySelector('#input-manifest');root.replaceChildren();
    if(!item) {root.append(make('p','review-muted','该领域暂未找到可对照的同事原题。'));return;}
    const prompt=make('p','review-input-prompt');
    prompt.append(make('strong','','Prompt：'),a(item.instruction,state.source==='current'?'instruction.md':'TASK_INSTRUCTIONS.txt'));
    root.append(prompt,make('p','review-input-path','附件挂载目录：/app/input_files/（只读）'),fileList(item.files));
  }
  function renderFiles() {
    filePanel.replaceChildren();
    const item=bundle[state.domain];if(!item) {filePanel.append(make('p','review-muted','正在载入文件索引…'));return;}
    const map={instruction:['题面','解题模型可见'],inputs:['原始输入','解题模型可见 · 只读'],reference:['参考答案','裁判与人工审阅'],rubric:['Rubric 与判分配置','裁判与人工审阅'],outputs:['模型实际产出','待评分的交付物'],evidence:['执行与评分记录','审阅证据']};
    filePanel.append(make('h3','review-task-label',item.task));
    const groups=item.groups;
    if(!groups.instruction) groups.instruction=item.instruction?[item.instruction]:[];
    for(const [key,[title,visibility]] of Object.entries(map)) {
      const files=groups[key]||[];
      const section=make('section','review-file-group');
      const head=make('div','review-section-head');head.append(make('h3','',title+' · '+files.length),make('span','review-badge',visibility));
      section.append(head);
      if(files.length) section.append(fileList(files));else section.append(make('p','review-muted','当前审阅包未收录此类文件。'));
      filePanel.append(section);
    }
    const colleague=inputs[state.domain]?.colleague;
    if(colleague) {
      const section=make('section','review-file-group');section.append(make('h3','','同事原题 · '+colleague.task));
      section.append(a(colleague.instruction,'TASK_INSTRUCTIONS.txt'),fileList(colleague.files));filePanel.append(section);
    }
  }
  function render() {
    app.querySelectorAll('.review-domain-tabs button').forEach(b=>{b.setAttribute('aria-selected',String(b.dataset.domain===state.domain));b.tabIndex=b.dataset.domain===state.domain?0:-1;});
    app.querySelectorAll('.review-compare,.review-task,.review-history').forEach(el=>el.hidden=el.dataset.domain!==state.domain);
    for(const [key,panel] of Object.entries(panelMap)) panel.hidden=key!==state.view;
    app.querySelectorAll('.review-sidebar nav a').forEach(link=>{
      if(link.dataset.view===state.view) link.setAttribute('aria-current','page');else link.removeAttribute('aria-current');
    });
    renderInputs();renderFiles();
  }
  function route(scroll=false) {
    const hash=decodeURIComponent(location.hash.slice(1));
    if(hash.startsWith('compare-')&&domains.includes(hash.slice(8))) {state.domain=hash.slice(8);state.view='colleague';}
    else if(hash.startsWith('task-')&&domains.includes(hash.split('-')[1])) {state.domain=hash.split('-')[1];state.view='all-tasks';}
    else if(['instruction','inputs','reference','rubric','outputs','scores'].includes(hash)) {state.view='fin-review';state.domain='FIN';}
    else if(panelMap[hash]) state.view=hash;
    render();
    if(scroll&&hash&&hash!=='model-inputs') {
      const target=document.getElementById(hash)||panelMap[state.view];
      requestAnimationFrame(()=>target?.scrollIntoView({behavior:'smooth',block:'start'}));
    }
  }
  app.addEventListener('click',event=>{
    const link=event.target.closest('a[href^="#"]');
    if(link&&link.hash===location.hash) {event.preventDefault();route(true);}
  });
  window.addEventListener('hashchange',()=>route(true));
  app.querySelector('.review-domain-tabs').addEventListener('keydown',e=>{
    if(!['ArrowLeft','ArrowRight','Home','End'].includes(e.key))return;
    e.preventDefault();const i=domains.indexOf(state.domain);
    state.domain=e.key==='Home'?domains[0]:e.key==='End'?domains.at(-1):domains[(i+(e.key==='ArrowRight'?1:4))%5];
    state.source='current';render();app.querySelector('.review-domain-tabs button[aria-selected="true"]').focus();
  });
  const embedded=document.getElementById('review-metadata');
  const data=embedded?Promise.resolve(JSON.parse(embedded.textContent)):Promise.all([fetch('review-inputs.json').then(r=>{if(!r.ok)throw Error('输入索引加载失败');return r.json();}),fetch('review-bundle.json').then(r=>{if(!r.ok)throw Error('文件索引加载失败');return r.json();})]);
  data
    .then(([i,b])=>{inputs=i;bundle=b;route(false);app.dataset.ready='true';})
    .catch(error=>{
      app.querySelector('#input-manifest').append(make('p','review-load-error',error.message+'。请通过 HTTP 预览此文件夹，原始证据仍可在“金融完整证据”中阅读。'));
      console.error(error);
    });
  route(false);
})();
