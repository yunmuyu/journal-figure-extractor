const $ = (s) => document.querySelector(s);
const wordFolder = $('#wordFolder');
const revisedFolder = $('#revisedFolder');
const wordPicked = $('#wordPicked');
const revisedPicked = $('#revisedPicked');
const wordStats = $('#wordStats');
const revisedStats = $('#revisedStats');
const wordCount = $('#wordCount');
const revisedCount = $('#revisedCount');
const extractOnly = $('#extractOnly');
const prepare = $('#prepare');
const prepareStatus = $('#prepareStatus');
const prepareProgress = $('#prepareProgress');
const prepareError = $('#prepareError');
const pairSummary = $('#pairSummary');
const unmatched = $('#unmatched');
const resultsCard = $('#resultsCard');
const pairsEl = $('#pairs');
const providerMode = $('#providerMode');
const precisionMode = $('#precisionMode');
const persistKey = $('#persistKey');
const compareAll = $('#compareAll');
const stopCompare = $('#stopCompare');
const exportReport = $('#exportReport');
const aiProgress = $('#aiProgress');

const providers = ['bailian','groq','openrouter'];
const providerLabels = {bailian:'阿里云百炼 Qwen',groq:'Groq Qwen',openrouter:'OpenRouter Free'};
const providerConfigured = {bailian:false,groq:false,openrouter:false};
let currentJob = null;
let pairData = [];
let stopRequested = false;

for (const p of providers) {
  const sel = $(`#model-${p}`);
  const saved = localStorage.getItem(`jfe_v11_model_${p}`);
  if (saved && [...sel.options].some(o => o.value === saved)) sel.value = saved;
  sel.addEventListener('change', () => localStorage.setItem(`jfe_v11_model_${p}`, sel.value));
}
const savedMode = localStorage.getItem('jfe_v12_provider_mode');
if (savedMode && [...providerMode.options].some(o => o.value === savedMode)) providerMode.value = savedMode;
providerMode.addEventListener('change', () => localStorage.setItem('jfe_v12_provider_mode', providerMode.value));
const savedPrecision = localStorage.getItem('jfe_v12_precision_mode'); if(savedPrecision && precisionMode && [...precisionMode.options].some(o=>o.value===savedPrecision)) precisionMode.value=savedPrecision; if(precisionMode) precisionMode.addEventListener('change',()=>localStorage.setItem('jfe_v12_precision_mode', precisionMode.value));

function wordFiles(){ return [...wordFolder.files].filter(f => /\.(doc|docx)$/i.test(f.name)); }
function revisedFiles(){ return [...revisedFolder.files].filter(f => /\.(pdf|jpe?g|png)$/i.test(f.name)); }
function rootName(files){ return files.length ? ((files[0].webkitRelativePath || files[0].name).split('/')[0]) : ''; }
function updateReady(){ prepare.disabled = !(wordFiles().length && revisedFiles().length); }
function modelsPayload(){ return Object.fromEntries(providers.map(p => [p, $(`#model-${p}`).value])); }
function anyConfigured(){ return providers.some(p => providerConfigured[p]); }
function selectedProviderReady(){ const m=providerMode.value; return m==='auto' ? anyConfigured() : !!providerConfigured[m]; }

wordFolder.addEventListener('change', () => {
  const fs=wordFiles(); wordPicked.textContent=fs.length?`${rootName(fs)}（${fs.length} 个 Word）`:'没有找到 Word';
  wordCount.textContent=fs.length; wordStats.hidden=!fs.length; extractOnly.disabled=!fs.length; updateReady();
});
revisedFolder.addEventListener('change', () => {
  const fs=revisedFiles(); revisedPicked.textContent=fs.length?`${rootName(fs)}（${fs.length} 个重制图）`:'没有找到 PDF/JPG/PNG';
  revisedCount.textContent=fs.length; revisedStats.hidden=!fs.length; updateReady();
});

async function refreshProviderStatus(){
  try{
    const r=await fetch('/api/providers/status'); const j=await r.json();
    for(const p of providers){
      const info=(j.providers||{})[p]||{}; providerConfigured[p]=!!info.configured;
      const el=$(`#status-${p}`); el.textContent=info.configured?`已配置 · ${info.source}`:'未配置'; el.className=`pill ${info.configured?'ok':'neutral'}`;
    }
  }catch(e){ for(const p of providers){ const el=$(`#status-${p}`); el.textContent='状态未知'; } }
}
refreshProviderStatus();

document.querySelectorAll('[data-save]').forEach(btn => btn.addEventListener('click', async () => {
  const p=btn.dataset.save; const key=$(`#key-${p}`).value.trim(); const err=$(`#error-${p}`); const st=$(`#status-${p}`);
  err.hidden=true; if(!key){err.textContent='请输入 API Key。';err.hidden=false;return;}
  btn.disabled=true; st.textContent='正在测试…'; st.className='pill neutral';
  try{
    const r=await fetch('/api/providers/key',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({provider:p,api_key:key,model:$(`#model-${p}`).value,persist:persistKey.checked})});
    const j=await r.json(); if(!r.ok) throw new Error(j.error||`HTTP ${r.status}`);
    providerConfigured[p]=true; st.textContent=`Key 可用 · ${j.source}`; st.className='pill ok'; $(`#key-${p}`).value='';
  }catch(e){ providerConfigured[p]=false; st.textContent='测试失败'; st.className='pill bad'; err.textContent=e.message; err.hidden=false; }
  finally{btn.disabled=false;}
}));

document.querySelectorAll('[data-clear]').forEach(btn => btn.addEventListener('click', async () => {
  const p=btn.dataset.clear; await fetch(`/api/providers/key/${encodeURIComponent(p)}`,{method:'DELETE'});
  $(`#key-${p}`).value=''; providerConfigured[p]=false; const st=$(`#status-${p}`); st.textContent='未配置'; st.className='pill neutral';
}));

extractOnly.addEventListener('click', async () => {
  const fs=wordFiles(); if(!fs.length)return; extractOnly.disabled=true; const fd=new FormData(); fs.forEach(f=>fd.append('files',f,f.name));
  try{ const r=await fetch('/extract',{method:'POST',body:fd}); if(!r.ok){let j={};try{j=await r.json()}catch{};throw new Error(j.error||`HTTP ${r.status}`)} const blob=await r.blob();downloadBlob(blob,'图件提取结果.zip'); }
  catch(e){alert('提取失败：'+e.message)} finally{extractOnly.disabled=false;}
});

prepare.addEventListener('click', async () => {
  prepareError.hidden=true; unmatched.hidden=true; pairSummary.hidden=true; resultsCard.hidden=true;
  const w=wordFiles(),r=revisedFiles(); if(!w.length||!r.length)return;
  prepare.disabled=true; prepareProgress.hidden=false; prepareStatus.textContent='Word 正在后台提取、重制图正在渲染并配对…';
  const fd=new FormData(); w.forEach(f=>fd.append('word_files',f,f.name)); r.forEach(f=>fd.append('revised_files',f,f.name)); fd.append('revised_relpaths',JSON.stringify(r.map(f=>f.webkitRelativePath||f.name)));
  try{
    const res=await fetch('/api/prepare-compare',{method:'POST',body:fd}); const j=await res.json(); if(!res.ok)throw new Error(j.error||`HTTP ${res.status}`);
    currentJob=j.job_id; pairData=j.pairs||[]; pairSummary.innerHTML=`<b>${j.article_count}</b> 篇原稿　·　<b>${j.original_figure_count}</b> 张原稿图　·　<b>${j.revised_count}</b> 张重制图　·　<b>${j.matched_count}</b> 对已配对`+((j.unmatched_original_count||j.unmatched_revised_count)?`　·　<span class="danger">${j.unmatched_original_count+j.unmatched_revised_count} 项未匹配</span>`:''); pairSummary.hidden=false;
    if(j.unmatched_original_count||j.unmatched_revised_count||(j.word_failures||[]).length){const lines=[];(j.unmatched_originals||[]).forEach(x=>lines.push(`原稿未匹配：${x.author} 图${x.figure_no} ${x.caption||''}`));(j.unmatched_revised||[]).forEach(x=>lines.push(`重制图未匹配：${x.relative_path||x.source_name||''} —— ${x.reason||(x.warnings||[]).join('；')}`));(j.word_failures||[]).forEach(x=>lines.push(`Word 处理失败：${x.file} —— ${x.error}`));unmatched.textContent=lines.join('\n');unmatched.hidden=false;}
    renderPairs(); resultsCard.hidden=!pairData.length; prepareStatus.textContent='配对完成。先肉眼确认左右图，再开始 AI 校对。'; resultsCard.scrollIntoView({behavior:'smooth',block:'start'});
  }catch(e){prepareError.textContent=e.message;prepareError.hidden=false;prepareStatus.textContent='未完成';}
  finally{prepare.disabled=false;prepareProgress.hidden=true;}
});

function imageUrl(p,side){return `/api/job/${encodeURIComponent(currentJob)}/pair/${encodeURIComponent(p.id)}/image/${side}?t=${Date.now()}`;}
function escapeHtml(s){return String(s??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));}
function statusLabel(status){return ({PENDING:'待校对',RUNNING:'校对中…',PASS:'PASS · 一致',REVIEW:'REVIEW · 需复核',FAIL:'FAIL · 有差异',ERROR:'ERROR · 调用失败'})[status]||status;}
function resultHtml(comp,status){
  if(status==='PENDING')return '<p class="empty">尚未调用 AI。</p>';
  if(status==='RUNNING')return '<p class="empty">AI 正在读取两张图并逐项比较…</p>';
  if(status==='ERROR')return `<div class="error-inline">${escapeHtml((comp||{}).error||'调用失败')}</div>`;
  comp=comp||{};
  const diffs=(comp.differences||[]).map(d=>`<div class="diff"><div class="diff-top"><b>${escapeHtml(d.type||'差异')}</b><span>${escapeHtml(d.severity||'')}</span></div><div class="loc">${escapeHtml(d.location||'')}</div><div><em>原稿</em> ${escapeHtml(d.original||'—')}</div><div><em>重制</em> ${escapeHtml(d.revised||'—')}</div>${d.reason?`<p>${escapeHtml(d.reason)}</p>`:''}</div>`).join('');
  const uncertain=(comp.uncertainties||[]).length?`<div class="uncertain"><b>无法确认</b><br>${comp.uncertainties.map(escapeHtml).join('<br>')}</div>`:'';
  const ignored=(comp.ignored_changes||[]).length?`<details><summary>已忽略的样式变化</summary><p>${comp.ignored_changes.map(escapeHtml).join('；')}</p></details>`:'';
  const ta=comp.text_audit||null;
  let auditHtml='';
  if(ta){
    const tds=(ta.text_differences||[]).map(d=>`<div class="diff"><div class="diff-top"><b>逐字核对 · ${escapeHtml(d.status||'差异')}</b></div><div class="loc">${escapeHtml(d.location||'')}</div><div><em>原稿</em> ${escapeHtml(d.original||'—')}</div><div><em>重制</em> ${escapeHtml(d.revised||'—')}</div>${d.reason?`<p>${escapeHtml(d.reason)}</p>`:''}</div>`).join('');
    const oa=(ta.backend_only_original||[]).map(escapeHtml).join('；'); const ob=(ta.backend_only_revised||[]).map(escapeHtml).join('；');
    const inv=ta.backend_inventory_match?'<span class="pill ok">文字库存一致</span>':`<span class="pill bad">文字库存未对齐</span>${oa?`<p><b>仅原稿：</b>${oa}</p>`:''}${ob?`<p><b>仅重制：</b>${ob}</p>`:''}`;
    auditHtml=`<details open><summary>逐字文字清点</summary>${inv}${tds||'<p class="empty">AI 未列出明确逐字差异。</p>'}</details>`;
  }
  const who=[comp.provider_label,comp.model].filter(Boolean).join(' · '); const usage=comp.usage&&comp.usage.total_tokens?`<div class="usage">${escapeHtml(who)} · ${comp.usage.total_tokens} tokens${comp.precision==='strict'?' · 精确双阶段':''}</div>`:`<div class="usage">${escapeHtml(who)}</div>`;
  const fallback=comp.fallback_used?`<div class="uncertain"><b>自动容错</b><br>${escapeHtml(comp.fallback_note||'已切换备用 AI 服务完成校对。')}</div>`:'';
  return `<p class="summary-text">${escapeHtml(comp.summary||'')}</p>${auditHtml}${diffs||'<p class="empty">未列出实质差异。</p>'}${uncertain}${ignored}${fallback}${usage}`;
}

function renderPairs(){pairsEl.innerHTML='';pairData.forEach(p=>{const el=document.createElement('article');el.className='pair';el.id=`pair-${p.id}`;el.innerHTML=`<div class="pair-title"><div><span class="author">${escapeHtml(p.author)}</span><b>图${escapeHtml(p.figure_no)}</b><span>${escapeHtml(p.caption||'')}</span></div><span class="badge ${String(p.status||'PENDING').toLowerCase()}" data-role="badge">${statusLabel(p.status||'PENDING')}</span></div><div class="pair-grid"><div class="visual"><div class="col-title">作者原稿 <small>${escapeHtml(p.original_method||'')}</small></div><a href="${imageUrl(p,'original')}" target="_blank"><img src="${imageUrl(p,'original')}" loading="lazy"></a></div><div class="visual"><div class="col-title">重制图 <small>${escapeHtml(p.revised_source||'')}</small></div><a href="${imageUrl(p,'revised')}" target="_blank"><img src="${imageUrl(p,'revised')}" loading="lazy"></a></div><div class="ai-col"><div class="col-title">AI 结论 <button class="tiny ghost" data-action="compare">校对这一张</button></div><div data-role="result">${resultHtml(p.comparison,p.status||'PENDING')}</div></div></div>`;el.querySelector('[data-action="compare"]').addEventListener('click',()=>compareOne(p.id));pairsEl.appendChild(el);});updateAiProgress();}
function updatePairUI(p){const el=$(`#pair-${CSS.escape(p.id)}`);if(!el)return;const badge=el.querySelector('[data-role="badge"]');badge.className=`badge ${String(p.status||'PENDING').toLowerCase()}`;badge.textContent=statusLabel(p.status||'PENDING');el.querySelector('[data-role="result"]').innerHTML=resultHtml(p.comparison,p.status||'PENDING');updateAiProgress();}
function updateAiProgress(){const c={PASS:0,FAIL:0,REVIEW:0,ERROR:0,PENDING:0,RUNNING:0};pairData.forEach(p=>c[p.status||'PENDING']=(c[p.status||'PENDING']||0)+1);aiProgress.textContent=`共 ${pairData.length} 对：PASS ${c.PASS} ｜ FAIL ${c.FAIL} ｜ REVIEW ${c.REVIEW} ｜ 待处理 ${c.PENDING+c.RUNNING} ｜ 错误 ${c.ERROR}`;}

async function compareOne(id){
  if(!currentJob)return; if(!selectedProviderReady()){alert('第 03 步还没有配置当前模式所需的 API Key。');return false;}
  const p=pairData.find(x=>x.id===id);if(!p)return; p.status='RUNNING';p.comparison=null;updatePairUI(p);
  try{const r=await fetch(`/api/job/${encodeURIComponent(currentJob)}/pair/${encodeURIComponent(id)}/compare`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({mode:providerMode.value,models:modelsPayload(),precision:precisionMode?precisionMode.value:'strict'})});const j=await r.json();if(!r.ok)throw new Error(j.error||`HTTP ${r.status}`);p.status=j.status;p.comparison=j.comparison;updatePairUI(p);return true;}
  catch(e){p.status='ERROR';p.comparison={error:e.message};updatePairUI(p);if(/429|频率|额度|超时|无法连接|所有可用 AI/i.test(e.message))throw e;return false;}
}
compareAll.addEventListener('click',async()=>{if(!currentJob||!pairData.length)return;if(!selectedProviderReady()){alert('先在第 03 步至少配置一个可用 API Key，或切换到已配置的服务。');return;}stopRequested=false;compareAll.disabled=true;stopCompare.disabled=false;try{for(const p of pairData){if(stopRequested)break;if(['PASS','FAIL','REVIEW'].includes(p.status))continue;try{await compareOne(p.id)}catch(e){alert(e.message+'\n\n批量校对已暂停。');break}await new Promise(r=>setTimeout(r,900));}}finally{compareAll.disabled=false;stopCompare.disabled=true;}});
stopCompare.addEventListener('click',()=>{stopRequested=true;stopCompare.disabled=true;});
exportReport.addEventListener('click',()=>{if(currentJob)location.href=`/api/job/${encodeURIComponent(currentJob)}/export`;});
function downloadBlob(blob,name){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),5000);}
