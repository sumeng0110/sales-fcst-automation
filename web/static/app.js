const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));

const state = { from: null, to: null, lastOutput: null, priorOutput: null };

async function api(path, { method = "GET", body, form } = {}) {
  const opts = { method };
  if (form) opts.body = form;
  else if (body) {
    opts.headers = { "Content-Type": "application/json" };
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(path, opts);
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); } catch { data = { detail: text }; }
  if (!res.ok) throw new Error(data.detail || res.statusText);
  return data;
}

function pair() {
  if (!state.from || !state.to) throw new Error("请先选择期间");
  return { from_period: state.from, to_period: state.to };
}

function badge(n, text, cls) {
  const el = $("badge" + n);
  if (!el) return;
  el.textContent = text;
  el.className = "badge " + (cls || "");
  const stage = document.querySelector(`.stage[data-stage="${n}"]`);
  if (stage) stage.classList.toggle("done", cls === "ok"), stage.classList.toggle("error", cls === "err");
}

// ---------------------------------------------------------------- job log

let poller = null;

function showLog(title) {
  $("logPanel").classList.remove("hidden");
  $("logTitle").textContent = title;
  $("logBody").textContent = "";
}

function appendEvents(events, seen) {
  const body = $("logBody");
  for (let i = seen; i < events.length; i++) {
    const e = events[i];
    const line = document.createElement("span");
    line.className = "lv-" + e.level;
    line.textContent = e.message + "\n";
    body.appendChild(line);
  }
  body.parentElement.scrollTop = body.parentElement.scrollHeight;
  return events.length;
}

function runJob(title, jobId, onDone) {
  showLog(title);
  let seen = 0;
  clearInterval(poller);
  poller = setInterval(async () => {
    let job;
    try { job = await api(`/api/jobs/${jobId}`); }
    catch { return; }
    seen = appendEvents(job.events, seen);
    $("logElapsed").textContent = job.elapsed + " 秒";
    if (job.status !== "running") {
      clearInterval(poller);
      $("logTitle").textContent = title + (job.status === "done" ? " — 完成" : " — 失败");
      onDone(job);
    }
  }, 1000);
}

// ---------------------------------------------------------------- stage 1

async function plan1() {
  badge(1, "读取中…", "run");
  const d = await api("/api/copy/plan?include_feedback=" + $("includeFeedback").checked,
    { method: "POST", body: pair() });
  const rows = d.renames.map((r) => `<tr><td class="mono wrap">${esc(r.from)}</td><td class="mono wrap">${esc(r.to)}</td></tr>`).join("");
  $("out1").innerHTML = `
    ${d.target_exists ? '<div class="notice warn">目标目录已存在，已存在的文件会被跳过，不会覆盖。</div>' : ""}
    <div class="summary">
      <div><strong>${d.file_count}</strong>个文件</div>
      <div><strong>${d.total_mb}</strong>MB</div>
      <div><strong>${d.renames.length}</strong>个改名</div>
      <div><strong>${d.empty_folders.length}</strong>个只建空壳的目录</div>
    </div>
    <table><thead><tr><th>原文件名</th><th>新文件名</th></tr></thead><tbody>${rows}</tbody></table>
    <p class="dim" style="margin-top:10px">只建空壳：${d.empty_folders.map(esc).join("、") || "无"}</p>
    ${d.derived.length ? `<div class="notice warn">以下 ${d.derived.length} 个文件不复制：
      ${d.derived.map(esc).join("、")}<br>
      它是本期正式版定稿后才裁出来的（清空 销售收入_China D672:O672），
      搬上期的过来等于让 SCM 顶着本期文件名读上期的数。</div>` : ""}`;
  badge(1, "已预览", "warn");
  document.querySelector('[data-act="apply1"]').disabled = false;
}

async function apply1() {
  const d = await api("/api/copy/apply?include_feedback=" + $("includeFeedback").checked,
    { method: "POST", body: pair() });
  badge(1, "复制中…", "run");
  runJob("复制期间文件夹", d.job, (job) => {
    badge(1, job.status === "done" ? "已完成" : "失败", job.status === "done" ? "ok" : "err");
  });
}

// ---------------------------------------------------------------- stage 2

async function plan2() {
  badge(2, "读取中…", "run");
  const d = await api("/api/links/plan", { method: "POST", body: pair() });
  let html = "";
  if (d.unresolved) html += `<div class="notice err">${d.unresolved} 个链接无法判断，必须先处理才能执行。</div>`;
  if (d.waiting) html += `<div class="notice warn">${d.waiting} 个链接指向本月还没到的输入文件，会先写上，文件到了自动生效。</div>`;
  if (!d.unresolved && !d.waiting) html += `<div class="notice ok">所有链接都能确定去向。</div>`;

  for (const wb of d.workbooks) {
    if (wb.missing) {
      html += `<div class="notice err">${esc(wb.workbook)}：文件不存在，请先执行第 1 步。</div>`;
      continue;
    }
    const rows = wb.links.map((l) => {
      const cls = l.status === "unresolved" ? "flag" : l.status === "waiting" ? "soft" : "";
      const target = l.changed
        ? `${esc(l.new_name)}<div class="dim mono wrap">原：${esc(l.old_name)}</div>`
        : `<span class="dim">${esc(l.new_name)}</span>`;
      return `<tr class="${cls}"><td class="num-col">${l.no}</td>
        <td><span class="pill ${l.status}">${l.status}</span></td>
        <td>${target}</td><td class="dim">${esc(l.note)}</td></tr>`;
    }).join("");
    html += `<h3 style="font-size:14px;margin:16px 0 6px">${esc(wb.workbook)} · <span class="dim mono">${esc(wb.file)}</span></h3>
      <table><thead><tr><th class="num-col">#</th><th>处理</th><th>指向</th><th>说明</th></tr></thead><tbody>${rows}</tbody></table>`;
  }
  $("out2").innerHTML = html;
  badge(2, d.unresolved ? "有待处理" : "已预览", d.unresolved ? "err" : "warn");
  document.querySelector('[data-act="apply2"]').disabled = d.unresolved > 0;
}

async function apply2() {
  const d = await api("/api/links/apply", { method: "POST", body: pair() });
  badge(2, "重指中…", "run");
  runJob("重指外部链接", d.job, (job) => {
    badge(2, job.status === "done" ? "已完成" : "失败", job.status === "done" ? "ok" : "err");
  });
}

// ---------------------------------------------------------------- stage 3

async function check3() {
  badge(3, "检查中…", "run");
  const d = await api(`/api/hfm/check?to_period=${encodeURIComponent(state.to)}`);
  let html = d.stale.length
    ? `<div class="notice err">以下网格还停在更早的月份，请在 Smart View 里刷新后重新检查：<br>${d.stale.map(esc).join("<br>")}</div>`
    : `<div class="notice ok">所有在用网格都已刷到 P${String(d.expected_month).padStart(2, "0")}。</div>`;

  for (const wb of d.workbooks) {
    if (wb.missing || !wb.grids.length) continue;
    html += `<h3 style="font-size:14px;margin:16px 0 6px">${esc(wb.workbook)}</h3>`;
    for (const g of wb.grids) {
      const months = g.months.map((m) => {
        const cls = m.invalid ? "empty" : m.populated ? "filled" : "empty";
        return `<span class="month ${cls}">P${String(m.period).padStart(2, "0")}</span>`;
      }).join("");
      html += `<div style="margin-bottom:10px">
        <span class="pill ${g.ok ? "rolled" : "unresolved"}">${esc(g.sheet)}</span>
        <span class="dim" style="margin-left:8px">${esc(g.verdict)}</span>
        <div class="months">${months}</div></div>`;
    }
  }
  $("out3").innerHTML = html;
  badge(3, d.stale.length ? "需要刷新" : "已通过", d.stale.length ? "err" : "ok");
}

// ---------------------------------------------------------------- stage 4

async function run4() {
  const d = await api("/api/recalc?skip_hfm_check=" + $("skipHfm").checked,
    { method: "POST", body: pair() });
  badge(4, "重算中…", "run");
  runJob("Excel 重算", d.job, (job) => {
    badge(4, job.status === "done" ? "已完成" : "失败", job.status === "done" ? "ok" : "err");
    if (job.status === "done" && job.result.saved?.length) {
      state.lastOutput = job.result.saved[job.result.saved.length - 1];
      $("out4").innerHTML = `<div class="notice ok">已重算并保存 ${job.result.saved.length} 个工作簿。</div>`;
    }
  });
}

// ---------------------------------------------------------------- stage 5

async function run5() {
  if (!state.lastOutput) {
    const wb = await api("/api/links/plan", { method: "POST", body: pair() });
    const p110 = wb.workbooks.find((w) => w.workbook === "P110");
    if (!p110 || p110.missing) throw new Error("找不到本期 P110，请先完成前面的步骤");
    state.lastOutput = p110.path;
  }
  const prior = state.lastOutput
    .replace(new RegExp(state.to.replace("+", "\\+"), "g"), state.from);
  badge(5, "比对中…", "run");
  const d = await api("/api/compare", { method: "POST", body: { left: state.lastOutput, right: prior } });
  runJob("逐格比对", d.job, (job) => {
    if (job.status !== "done") { badge(5, "失败", "err"); return; }
    const r = job.result;
    const rows = r.rows.map((x) => `<tr class="${x.diff > 0 ? "soft" : ""}">
      <td>${esc(x.sheet)}</td><td class="num-col">${x.same}</td><td class="num-col">${x.diff}</td>
      <td class="dim mono">${x.examples.map((e) => `${esc(e.cell)}: ${esc(e.left)} / ${esc(e.right)}`).join("<br>")}</td>
    </tr>`).join("");
    const total = r.same + r.diff;
    $("out5").innerHTML = `
      <div class="summary">
        <div><strong>${total.toLocaleString()}</strong>比对单元格</div>
        <div><strong>${r.diff.toLocaleString()}</strong>处差异</div>
        <div><strong>${total ? ((r.same / total) * 100).toFixed(2) : 0}%</strong>一致率</div>
      </div>
      <p class="dim mono wrap">左：${esc(state.lastOutput)}<br>右：${esc(prior)}</p>
      <table><thead><tr><th>工作表</th><th class="num-col">一致</th><th class="num-col">差异</th><th>样例（左 / 右）</th></tr></thead>
      <tbody>${rows}</tbody></table>`;
    badge(5, "已完成", "ok");
  });
}

// ---------------------------------------------------------------- upload

async function uploadFiles(files) {
  const out = $("outUpload");
  out.innerHTML = "";
  for (const f of files) {
    const fd = new FormData();
    fd.append("to_period", state.to);
    fd.append("subfolder", $("uploadFolder").value);
    fd.append("file", f);
    try {
      const d = await api("/api/upload", { method: "POST", form: fd });
      out.innerHTML += `<div class="notice ok">已保存 ${esc(f.name)} → <span class="mono wrap">${esc(d.saved)}</span></div>`;
    } catch (e) {
      out.innerHTML += `<div class="notice err">${esc(f.name)}：${esc(e.message)}</div>`;
    }
  }
}

// ---------------------------------------------------------------- boot

async function loadPeriods() {
  const d = await api("/api/periods");
  $("rootPath").textContent = d.period_root;
  const sel = $("fromPeriod");
  sel.innerHTML = d.periods.map((p) =>
    `<option value="${esc(p.period)}">${esc(p.folder)}</option>`).join("");
  if (d.suggested_from) sel.value = d.suggested_from;
  if (d.suggested_to) $("toPeriod").value = d.suggested_to;
  state.from = sel.value;
  state.to = $("toPeriod").value;
}

function wire() {
  document.body.addEventListener("click", async (ev) => {
    const btn = ev.target.closest("button[data-act]");
    if (!btn) return;
    const handlers = { plan1, apply1, plan2, apply2, check3, run4, run5 };
    const fn = handlers[btn.dataset.act];
    if (!fn) return;
    btn.disabled = true;
    try { await fn(); }
    catch (e) { alert(e.message); }
    finally { if (!btn.dataset.act.startsWith("apply")) btn.disabled = false; }
  });

  $("fromPeriod").addEventListener("change", (e) => { state.from = e.target.value; });
  $("toPeriod").addEventListener("input", (e) => { state.to = e.target.value.trim(); });
  $("reloadBtn").addEventListener("click", () => loadPeriods().catch((e) => alert(e.message)));
  $("logClose").addEventListener("click", () => $("logPanel").classList.add("hidden"));

  const dz = $("dropzone");
  dz.addEventListener("click", () => $("fileInput").click());
  $("fileInput").addEventListener("change", (e) => uploadFiles(e.target.files));
  ["dragenter", "dragover"].forEach((t) => dz.addEventListener(t, (e) => {
    e.preventDefault(); dz.classList.add("hot");
  }));
  ["dragleave", "drop"].forEach((t) => dz.addEventListener(t, (e) => {
    e.preventDefault(); dz.classList.remove("hot");
  }));
  dz.addEventListener("drop", (e) => uploadFiles(e.dataTransfer.files));
}

wire();
loadPeriods().catch((e) => { $("rootPath").textContent = "连接失败：" + e.message; });
