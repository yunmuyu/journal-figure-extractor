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

  let detail = document.getElementById('prepareProgressDetail');
  if (!detail && progressBox) {
    detail = document.createElement('div');
    detail.id = 'prepareProgressDetail';
    detail.innerHTML = `
      <style>
        #prepareProgressDetail{margin-top:12px;border:1px solid #e2ded6;background:#faf9f6;border-radius:12px;padding:13px 15px;font-size:13px;color:#555}
        #prepareProgressDetail .jfe-progress-top{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:10px}
        #prepareProgressDetail .jfe-progress-percent{font-size:22px;font-weight:800;color:#171717;min-width:58px}
        #prepareProgressDetail .jfe-progress-phase{font-weight:800;color:#222}
        #prepareProgressDetail .jfe-progress-detail{margin-top:3px;word-break:break-all;color:#666}
        #prepareProgressDetail .jfe-progress-steps{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-top:10px}
        #prepareProgressDetail .jfe-step{border:1px solid #e5e1d9;background:#fff;border-radius:9px;padding:9px 10px;line-height:1.55}
        #prepareProgressDetail .jfe-step b{display:block;color:#222;font-size:12px}
        #prepareProgressDetail .jfe-step span{font-variant-numeric:tabular-nums}
        #prepareProgressDetail .jfe-step.active{border-color:#999;background:#f4f2ed}
        #prepareProgressDetail .jfe-step.done{border-color:#b7d6bf;background:#eef8f0;color:#285d34}
        #prepareProgressDetail .jfe-step.bad{border-color:#efcaca;background:#fff1f1;color:#8c1f1f}
        @media(max-width:720px){#prepareProgressDetail .jfe-progress-steps{grid-template-columns:1fr}}
      </style>
      <div class="jfe-progress-top">
        <div><div class="jfe-progress-phase">等待开始</div><div class="jfe-progress-detail">选择文件后点击“提取并配对”。</div></div>
        <div class="jfe-progress-percent">0%</div>
      </div>
      <div class="jfe-progress-steps">
        <div class="jfe-step" data-jfe-step="word"><b>① Word 原稿</b><span>0 / 0</span></div>
        <div class="jfe-step" data-jfe-step="revised"><b>② 重制图渲染</b><span>0 / 0</span></div>
        <div class="jfe-step" data-jfe-step="pair"><b>③ 自动配对</b><span>等待</span></div>
      </div>
    `;
    progressBox.appendChild(detail);
  }

  const phaseEl = detail?.querySelector('.jfe-progress-phase');
  const detailEl = detail?.querySelector('.jfe-progress-detail');
  const percentEl = detail?.querySelector('.jfe-progress-percent');
  const wordStep = detail?.querySelector('[data-jfe-step="word"]');
  const revisedStep = detail?.querySelector('[data-jfe-step="revised"]');
  const pairStep = detail?.querySelector('[data-jfe-step="pair"]');

  function fmtTime(seconds) {
    seconds = Math.max(0, Number(seconds || 0));
    const m = Math.floor(seconds / 60);
    const s = seconds % 60;
    return m ? `${m}分${String(s).padStart(2,'0')}秒` : `${s}秒`;
  }

  function setStep(el, text, state) {
    if (!el) return;
    const span = el.querySelector('span');
    if (span) span.textContent = text;
    el.classList.remove('active','done','bad');
    if (state) el.classList.add(state);
  }

  function paint(state = {}) {
    const pct = Math.max(0, Math.min(100, Math.round(Number(state.percent || 0))));
    if (barFill) barFill.style.width = `${pct}%`;
    if (percentEl) percentEl.textContent = `${pct}%`;
    if (phaseEl) phaseEl.textContent = state.phase_label || '处理中';
    if (detailEl) {
      const elapsed = state.elapsed_seconds != null ? ` · 已用 ${fmtTime(state.elapsed_seconds)}` : '';
      detailEl.textContent = `${state.detail || ''}${elapsed}`;
    }

    const wt = Number(state.word_total || 0), wd = Number(state.word_done || 0), wf = Number(state.word_failed || 0);
    const rt = Number(state.revised_total || 0), rd = Number(state.revised_done || 0), rf = Number(state.revised_failed || 0);
    const phase = state.phase || '';

    setStep(wordStep, `${wd} / ${wt}${wf ? ` · 失败 ${wf}` : ''}`, wf ? 'bad' : (wt && wd >= wt ? 'done' : (phase === 'word' ? 'active' : '')));
    setStep(revisedStep, `${rd} / ${rt}${rf ? ` · 失败 ${rf}` : ''}`, rf ? 'bad' : (rt && rd >= rt ? 'done' : (phase === 'revised' ? 'active' : '')));
    const ps = state.pair_state || 'waiting';
    setStep(pairStep, ps === 'done' ? '完成' : ps === 'running' ? '正在配对' : '等待', ps === 'done' ? 'done' : ps === 'running' ? 'active' : '');
  }

  function xhrJson(url, formData, onUpload) {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open('POST', url, true);
      xhr.responseType = 'json';
      xhr.upload.onprogress = (e) => {
        if (e.lengthComputable && onUpload) onUpload(e.loaded, e.total);
      };
      xhr.onload = () => {
        let data = xhr.response;
        if (!data) {
          try { data = JSON.parse(xhr.responseText || '{}'); } catch { data = {}; }
        }
        if (xhr.status >= 200 && xhr.status < 300) resolve(data || {});
        else reject(new Error((data && data.error) || `HTTP ${xhr.status}`));
      };
      xhr.onerror = () => reject(new Error('网络连接失败'));
      xhr.ontimeout = () => reject(new Error('请求超时'));
      xhr.send(formData);
    });
  }

  function newProgressId() {
    if (window.crypto?.randomUUID) return window.crypto.randomUUID().replace(/-/g,'').slice(0,20);
    return Array.from(crypto.getRandomValues(new Uint8Array(10))).map(x=>x.toString(16).padStart(2,'0')).join('');
  }

  async function pollProgress(pid, controller) {
    while (!controller.stop) {
      try {
        const r = await fetch(`/api/prepare-progress/${encodeURIComponent(pid)}?t=${Date.now()}`, {cache:'no-store'});
        if (r.ok) {
          const j = await r.json();
          if (j.status !== 'waiting') paint(j);
          if (j.status === 'done' || j.status === 'error') return j;
        }
      } catch (_) {}
      await new Promise(r => setTimeout(r, 450));
    }
    return null;
  }

  function showFinalResult(j) {
    currentJob = j.job_id;
    pairData = j.pairs || [];
    summaryBox.innerHTML = `<b>${j.article_count}</b> 篇原稿　·　<b>${j.original_figure_count}</b> 张原稿图　·　<b>${j.revised_count}</b> 张重制图　·　<b>${j.matched_count}</b> 对已配对` + ((j.unmatched_original_count || j.unmatched_revised_count) ? `　·　<span class="danger">${j.unmatched_original_count + j.unmatched_revised_count} 项未匹配</span>` : '');
    summaryBox.hidden = false;

    if (j.unmatched_original_count || j.unmatched_revised_count || (j.word_failures || []).length) {
      const lines = [];
      (j.unmatched_originals || []).forEach(x => lines.push(`原稿未匹配：${x.author} 图${x.figure_no} ${x.caption || ''}`));
      (j.unmatched_revised || []).forEach(x => lines.push(`重制图未匹配：${x.relative_path || x.source_name || ''} —— ${x.reason || (x.warnings || []).join('；')}`));
      (j.word_failures || []).forEach(x => lines.push(`Word 处理失败：${x.file} —— ${x.error}`));
      unmatchedBox.textContent = lines.join('\n');
      unmatchedBox.hidden = false;
    }

    renderPairs();
    results.hidden = !pairData.length;
    statusText.textContent = '配对完成。先肉眼确认左右图，再开始 AI 校对。';
    if (results && !results.hidden) results.scrollIntoView({behavior:'smooth', block:'start'});
  }

  async function runPrepare() {
    errorBox.hidden = true;
    unmatchedBox.hidden = true;
    summaryBox.hidden = true;
    results.hidden = true;

    const w = wordFiles(), r = revisedFiles();
    if (!w.length || !r.length) return;

    btn.disabled = true;
    progressBox.hidden = false;
    statusText.textContent = '正在上传文件并启动处理……';
    paint({percent:0, phase_label:'上传文件', detail:'正在把所选 Word 与重制图发送到本机后台。', word_total:w.length, revised_total:r.length, word_done:0, revised_done:0, pair_state:'waiting', elapsed_seconds:0});

    const fd = new FormData();
    w.forEach(f => fd.append('word_files', f, f.name));
    r.forEach(f => fd.append('revised_files', f, f.name));
    fd.append('revised_relpaths', JSON.stringify(r.map(f => f.webkitRelativePath || f.name)));
    const pid = newProgressId();
    fd.append('progress_id', pid);

    const controller = {stop:false};
    const started = Date.now();
    const pollPromise = pollProgress(pid, controller);

    try {
      const j = await xhrJson('/api/prepare-compare', fd, (loaded, total) => {
        const ratio = total ? loaded / total : 0;
        const pct = Math.min(3, Math.max(0, Math.round(ratio * 3)));
        paint({percent:pct, phase_label:'上传文件', detail:`正在上传：${Math.round(ratio*100)}%`, word_total:w.length, revised_total:r.length, word_done:0, revised_done:0, pair_state:'waiting', elapsed_seconds:Math.floor((Date.now()-started)/1000)});
      });
      controller.stop = true;
      await pollPromise.catch(()=>null);
      paint({percent:100, phase_label:'完成', detail:'全部提取、渲染与配对已完成。', word_total:w.length, word_done:w.length, revised_total:r.length, revised_done:r.length, pair_state:'done', elapsed_seconds:Math.floor((Date.now()-started)/1000)});
      showFinalResult(j);
    } catch (e) {
      controller.stop = true;
      await pollPromise.catch(()=>null);
      errorBox.textContent = e.message;
      errorBox.hidden = false;
      statusText.textContent = '未完成';
      if (phaseEl) phaseEl.textContent = '处理失败';
      if (detailEl) detailEl.textContent = e.message;
    } finally {
      btn.disabled = false;
      // Keep the finished progress panel visible so the editor can see what happened.
    }
  }

  // Capture on document so the old v1.2 button listener never runs.
  document.addEventListener('click', (ev) => {
    const target = ev.target?.closest?.('#prepare');
    if (!target) return;
    ev.preventDefault();
    ev.stopPropagation();
    ev.stopImmediatePropagation();
    runPrepare();
  }, true);
})();
