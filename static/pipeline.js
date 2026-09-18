(() => {
  const q = (s) => document.querySelector(s);
  const btn = q('#prepare');
  if (!btn) return;

  const progressBox = q('#prepareProgress');
  const statusText = q('#prepareStatus');
  const errorBox = q('#prepareError');
  const summaryBox = q('#pairSummary');
  const unmatchedBox = q('#unmatched');
  const results = q('#resultsCard');
  const barFill = progressBox?.querySelector('.bar i');
  if (barFill) {
    barFill.style.animation = 'none';
    barFill.style.transform = 'none';
    barFill.style.transition = 'width .25s ease';
    barFill.style.width = '0%';
  }

  let selectedArticleId = '';
  let activeJobId = '';
  let lastPairSignature = '';
  let pipelineState = null;

  let panel = document.getElementById('pipelineArticlePanel');
  if (!panel) {
    panel = document.createElement('section');
    panel.id = 'pipelineArticlePanel';
    panel.hidden = true;
    panel.innerHTML = `
      <style>
        #pipelineArticlePanel{margin:18px 0 22px;border:1px solid #dedad2;background:#fff;border-radius:16px;padding:16px 16px 14px}
        #pipelineArticlePanel .pl-head{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;margin-bottom:12px}
        #pipelineArticlePanel .pl-title{font-size:18px;font-weight:800;color:#181818}
        #pipelineArticlePanel .pl-meta{font-size:13px;color:#666;margin-top:4px;line-height:1.7}
        #pipelineArticlePanel .pl-overall{font-size:24px;font-weight:900;font-variant-numeric:tabular-nums;white-space:nowrap}
        #pipelineArticlePanel .pl-buttons{display:flex;flex-wrap:wrap;gap:8px;max-height:230px;overflow:auto;padding:2px}
        #pipelineArticlePanel .pl-author{border:1px solid #dedbd4;background:#f8f7f4;border-radius:999px;padding:8px 12px;cursor:pointer;color:#333;font:inherit;display:inline-flex;gap:7px;align-items:center;max-width:360px}
        #pipelineArticlePanel .pl-author:hover{border-color:#999}
        #pipelineArticlePanel .pl-author.selected{outline:2px solid #171717;outline-offset:1px;background:#fff}
        #pipelineArticlePanel .pl-author.waiting{color:#888;background:#f7f7f7}
        #pipelineArticlePanel .pl-author.extracting,#pipelineArticlePanel .pl-author.matching,#pipelineArticlePanel .pl-author.ready_partial,#pipelineArticlePanel .pl-author.ready{border-color:#cfbea0;background:#fff9eb}
        #pipelineArticlePanel .pl-author.done{border-color:#a9cfb2;background:#eff9f1;color:#1e6532}
        #pipelineArticlePanel .pl-author.warning{border-color:#e4bf80;background:#fff5df;color:#7d5200}
        #pipelineArticlePanel .pl-author.error{border-color:#e0a9a9;background:#fff0f0;color:#9b2020}
        #pipelineArticlePanel .pl-author .pl-name{font-weight:800;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:220px}
        #pipelineArticlePanel .pl-author .pl-count{font-size:12px;font-variant-numeric:tabular-nums;white-space:nowrap}
        #pipelineArticlePanel .pl-current{margin-top:12px;background:#faf9f6;border-radius:10px;padding:10px 12px;font-size:13px;color:#555;line-height:1.7}
        @media(max-width:700px){#pipelineArticlePanel .pl-head{display:block}.pl-overall{margin-top:8px}}
      </style>
      <div class="pl-head">
        <div>
          <div class="pl-title">按作者实时出稿</div>
          <div class="pl-meta">一篇完成就可以先看，不必等待整批结束。处理中也会逐图增加。</div>
        </div>
        <div class="pl-overall">0 / 0</div>
      </div>
      <div class="pl-buttons"></div>
      <div class="pl-current">等待开始。</div>
    `;
    if (progressBox?.parentNode) progressBox.parentNode.insertBefore(panel, progressBox.nextSibling);
    else document.body.appendChild(panel);
  }

  const buttonsEl = panel.querySelector('.pl-buttons');
  const overallEl = panel.querySelector('.pl-overall');
  const currentEl = panel.querySelector('.pl-current');

  function fmtTime(seconds) {
    seconds = Math.max(0, Number(seconds || 0));
    const m = Math.floor(seconds / 60), s = seconds % 60;
    return m ? `${m}分${String(s).padStart(2,'0')}秒` : `${s}秒`;
  }

  function newProgressId() {
    if (window.crypto?.randomUUID) return window.crypto.randomUUID().replace(/-/g,'').slice(0,20);
    const a = new Uint8Array(10); crypto.getRandomValues(a);
    return Array.from(a).map(x=>x.toString(16).padStart(2,'0')).join('');
  }

  function wordFilesSafe() {
    try { if (typeof wordFiles === 'function') return wordFiles(); } catch (_) {}
    return Array.from(q('#wordFiles')?.files || q('input[type=file][accept*="doc"]')?.files || []);
  }
  function revisedFilesSafe() {
    try { if (typeof revisedFiles === 'function') return revisedFiles(); } catch (_) {}
    const inputs = Array.from(document.querySelectorAll('input[type=file]'));
    const inp = inputs.find(x => String(x.accept || '').toLowerCase().match(/pdf|jpg|jpeg|png/));
    return Array.from(inp?.files || []);
  }

  function xhrJson(url, fd, onUpload) {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open('POST', url, true);
      xhr.responseType = 'json';
      xhr.upload.onprogress = e => {
        if (e.lengthComputable && onUpload) onUpload(e.loaded, e.total);
      };
      xhr.onload = () => {
        let data = xhr.response;
        if (!data) { try { data = JSON.parse(xhr.responseText || '{}'); } catch { data = {}; } }
        if (xhr.status >= 200 && xhr.status < 300) resolve(data || {});
        else reject(new Error((data && data.error) || `HTTP ${xhr.status}`));
      };
      xhr.onerror = () => reject(new Error('连接本机后台失败'));
      xhr.send(fd);
    });
  }

  function setOverallProgress(st) {
    const total = Number(st.total_articles || 0), done = Number(st.completed_articles || 0);
    const pct = total ? Math.round(done * 100 / total) : 0;
    if (barFill) barFill.style.width = `${st.status === 'done' ? 100 : Math.min(99,pct)}%`;
    if (overallEl) overallEl.textContent = `${done} / ${total}`;
    const eta = st.eta_seconds ? ` · 预计剩余 ${fmtTime(st.eta_seconds)}` : '';
    const elapsed = ` · 已用 ${fmtTime(st.elapsed_seconds || 0)}`;
    if (statusText) statusText.textContent = `${st.detail || '处理中'}${elapsed}${eta}`;
    if (currentEl) currentEl.textContent = `已发现 ${Number(st.discovered_figures || 0)} 张原稿图，已有 ${Number(st.ready_pairs || 0)} 对可查看。${st.detail ? ' 当前：' + st.detail : ''}`;
  }

  function displayName(a, all) {
    if (a.author) {
      const same = all.filter(x => x.author && x.author === a.author).length;
      if (same > 1 && a.title) return `${a.author}｜${a.title.slice(0,14)}${a.title.length>14?'…':''}`;
      return a.author;
    }
    if (a.status === 'extracting') return `正在识别：${a.source_name}`;
    return a.source_name || a.id;
  }

  function renderArticleButtons(st) {
    const arts = st.articles || [];
    const oldScroll = buttonsEl.scrollTop;
    buttonsEl.innerHTML = '';
    for (const a of arts) {
      // Once the author is known, the button becomes the requested author button.
      // Waiting entries stay as filename placeholders so the user can see the queue.
      const b = document.createElement('button');
      b.type = 'button';
      b.className = `pl-author ${a.status || 'waiting'}${selectedArticleId===a.id?' selected':''}`;
      b.dataset.articleId = a.id;
      b.title = [a.author, a.title, a.error].filter(Boolean).join('\n');
      const count = Number(a.figure_total || 0) ? `${Number(a.ready_pairs || 0)}/${Number(a.figure_total || 0)}` : (a.status === 'waiting' ? '等待' : '…');
      let suffix = '';
      if (a.status === 'done') suffix = ' ✓';
      else if (a.status === 'warning') suffix = ' ⚠';
      else if (a.status === 'error') suffix = ' ✕';
      else if (['extracting','matching','ready_partial','ready'].includes(a.status)) suffix = ' ⟳';
      b.innerHTML = `<span class="pl-name"></span><span class="pl-count"></span>`;
      b.querySelector('.pl-name').textContent = displayName(a, arts);
      b.querySelector('.pl-count').textContent = count + suffix;
      b.disabled = !(a.pairs || []).length;
      b.addEventListener('click', () => selectArticle(a.id));
      buttonsEl.appendChild(b);
    }
    buttonsEl.scrollTop = oldScroll;
  }

  function selectArticle(articleId) {
    if (!pipelineState) return;
    const a = (pipelineState.articles || []).find(x => x.id === articleId);
    if (!a || !(a.pairs || []).length) return;
    selectedArticleId = articleId;
    activeJobId = pipelineState.job_id;
    try { currentJob = activeJobId; } catch (_) { window.currentJob = activeJobId; }
    try { pairData = a.pairs || []; } catch (_) { window.pairData = a.pairs || []; }
    try {
      if (typeof renderPairs === 'function') renderPairs();
    } catch (e) {
      console.error('renderPairs failed', e);
    }
    if (results) results.hidden = false;
    if (summaryBox) {
      summaryBox.innerHTML = `<b>${a.author || '作者识别中'}</b>　·　${a.title || a.source_name || ''}　·　<b>${(a.pairs||[]).length}</b> 张已就绪` + (a.unmatched_count ? `　·　<span class="danger">${a.unmatched_count} 张待人工匹配</span>` : '');
      summaryBox.hidden = false;
    }
    const sig = `${articleId}:${(a.pairs||[]).map(x=>x.id+':'+x.status).join(',')}`;
    lastPairSignature = sig;
    renderArticleButtons(pipelineState);
    results?.scrollIntoView({behavior:'smooth', block:'start'});
  }

  function refreshSelectedIfNeeded(st) {
    const arts = st.articles || [];
    if (!selectedArticleId) {
      const first = arts.find(a => (a.pairs || []).length);
      if (first) {
        // First article/first pair becomes viewable automatically while the background keeps working.
        selectArticle(first.id);
      }
      return;
    }
    const a = arts.find(x => x.id === selectedArticleId);
    if (!a) return;
    const sig = `${a.id}:${(a.pairs||[]).map(x=>x.id+':'+x.status).join(',')}`;
    if (sig !== lastPairSignature) selectArticle(a.id);
  }

  function renderUnmatched(st) {
    if (!unmatchedBox) return;
    const lines = [];
    for (const a of (st.articles || [])) {
      if (a.error) lines.push(`${a.author || a.source_name}：处理失败 —— ${a.error}`);
      else if (a.unmatched_count) lines.push(`${a.author || a.source_name}：有 ${a.unmatched_count} 张原稿图尚未安全匹配重制图`);
    }
    unmatchedBox.textContent = lines.join('\n');
    unmatchedBox.hidden = !lines.length;
  }

  async function pollPipeline(pid) {
    while (activeJobId === pid) {
      try {
        const r = await fetch(`/api/pipeline/${encodeURIComponent(pid)}/status?t=${Date.now()}`, {cache:'no-store'});
        if (!r.ok) throw new Error(`状态查询 HTTP ${r.status}`);
        const st = await r.json();
        pipelineState = st;
        panel.hidden = false;
        setOverallProgress(st);
        renderArticleButtons(st);
        renderUnmatched(st);
        refreshSelectedIfNeeded(st);
        if (st.status === 'done') {
          if (barFill) barFill.style.width = '100%';
          if (statusText) statusText.textContent = `整批完成：${st.completed_articles}/${st.total_articles} 篇 · 已用 ${fmtTime(st.elapsed_seconds || 0)}`;
          btn.disabled = false;
          return;
        }
        if (st.status === 'error') throw new Error(st.error || st.detail || '流水线失败');
      } catch (e) {
        if (errorBox) { errorBox.textContent = e.message; errorBox.hidden = false; }
        btn.disabled = false;
        return;
      }
      await new Promise(r => setTimeout(r, 550));
    }
  }

  async function runPipeline() {
    const w = wordFilesSafe(), r = revisedFilesSafe();
    if (!w.length || !r.length) return;
    if (errorBox) errorBox.hidden = true;
    if (unmatchedBox) unmatchedBox.hidden = true;
    if (summaryBox) summaryBox.hidden = true;
    if (results) results.hidden = true;
    selectedArticleId = '';
    lastPairSignature = '';
    pipelineState = null;
    panel.hidden = false;
    buttonsEl.innerHTML = '';
    overallEl.textContent = `0 / ${w.length}`;
    currentEl.textContent = `准备上传 ${w.length} 篇 Word 与 ${r.length} 张重制图。`;
    btn.disabled = true;
    if (progressBox) progressBox.hidden = false;

    const pid = newProgressId();
    activeJobId = pid;
    const fd = new FormData();
    w.forEach(f => fd.append('word_files', f, f.name));
    r.forEach(f => fd.append('revised_files', f, f.name));
    fd.append('revised_relpaths', JSON.stringify(r.map(f => f.webkitRelativePath || f.name)));
    fd.append('progress_id', pid);
    try {
      const ack = await xhrJson('/api/prepare-compare', fd, (loaded,total) => {
        const pct = total ? Math.round(loaded * 100 / total) : 0;
        if (statusText) statusText.textContent = `正在发送到本机后台：${pct}%`;
        if (barFill) barFill.style.width = `${Math.min(4, pct * 0.04)}%`;
        currentEl.textContent = `上传 ${pct}% · 上传完成后，后台会按文章逐篇处理，第一篇完成即可查看。`;
      });
      activeJobId = ack.job_id || pid;
      await pollPipeline(activeJobId);
    } catch (e) {
      if (errorBox) { errorBox.textContent = e.message; errorBox.hidden = false; }
      if (statusText) statusText.textContent = '启动失败';
      btn.disabled = false;
    }
  }

  // Capture phase guarantees the old one-shot v1.2/v1.4 handler never receives the click.
  document.addEventListener('click', (ev) => {
    const target = ev.target?.closest?.('#prepare');
    if (!target) return;
    ev.preventDefault();
    ev.stopPropagation();
    ev.stopImmediatePropagation();
    runPipeline();
  }, true);
})();
