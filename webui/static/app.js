"use strict";
/* The page draws whatever state() says and calls actions; no business logic lives here. */

const DEV = new URLSearchParams(location.search).has("dev");
let S = null;               // last state from Python
let view = "review";        // all | review | matched
let selected = null;        // selected result row id
let polling = false;

async function call(name, ...args) {
  if (DEV) {
    const r = await fetch("/api/" + name, {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(args)});
    return r.json();
  }
  return window.pywebview.api[name](...args);
}

function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]));
}

async function act(name, ...args) {
  try { setState(await call(name, ...args)); }
  catch (e) { showModal({kind: "error", title: "오류", message: String(e && e.message || e)}); }
}

function setState(next) {
  S = next;
  if (S.result) {
    const rows = S.result.rows;
    if (!rows.some(r => r.id === selected)) selected = null;
    if (view === "review" && !rows.some(r => !r.matched)) view = "matched";
  } else { view = "review"; selected = null; }
  render();
  (S.notices || []).forEach(n => queueModal(n));
  if (S.busy.adding || S.busy.running || (S.update && S.update.busy)) schedulePoll();
}

function schedulePoll() {
  if (polling) return;
  polling = true;
  setTimeout(async () => { polling = false; await act("get_state"); }, 300);
}

/* ---------- modal ---------- */
const modalQueue = [];
function queueModal(n) { modalQueue.push(n); if (document.getElementById("modal").hidden) nextModal(); }
function nextModal() {
  const n = modalQueue.shift();
  if (n) showModal(n); else document.getElementById("modal").hidden = true;
}
function showModal(n, onOk) {
  const m = document.getElementById("modal");
  m.hidden = false;
  m.innerHTML = `<div class="dlg ${esc(n.kind)}" role="dialog" aria-modal="true" aria-labelledby="mt">
    <h3 id="mt">${esc(n.title)}</h3><div class="msg">${esc(n.message)}</div>
    <div class="btns">${onOk ? '<button class="btn ghost" id="mno">취소</button>' : ""}<button class="btn primary" id="mok">${onOk ? "확인" : "닫기"}</button></div></div>`;
  const ok = document.getElementById("mok");
  ok.focus();
  ok.onclick = () => { m.hidden = true; if (onOk) onOk(); nextModal(); };
  const no = document.getElementById("mno");
  if (no) no.onclick = () => { m.hidden = true; nextModal(); };
}

/* ---------- render ---------- */
function render() {
  const s = S, p = s.period, f = s.files, u = s.update;
  const done = s.done && !s.inputs_open;
  const running = s.busy.running, adding = s.busy.adding;
  const app = document.getElementById("app");
  const keep = document.activeElement && document.activeElement.id;

  const updateBtn = u.available || u.busy
    ? `<button class="upd" id="upd" ${u.busy ? "disabled" : ""}>${u.busy ? `내려받는 중… ${u.percent}%` : "새 버전 업데이트"}</button>` : "";
  const updErr = u.error ? `<div class="upd-err">${esc(u.error)}</div>` : "";

  const fileRow = (kind, title, hint, list) => {
    const chips = list.map((n, i) => `<div class="chip-file"><span class="ok">✓</span><span class="nm" title="${esc(n)}">${esc(n)}</span>
        <button class="x" data-rm="${kind}:${i}" aria-label="${esc(n)} 제거">×</button></div>`).join("");
    return `<div class="row"><span class="lbl">${title}</span><div class="grow">
      ${list.length ? `<div class="chips">${chips}</div>` : ""}
      <button class="drop" data-add="${kind}" ${adding ? "disabled" : ""}><span class="plus">+</span>${list.length ? "파일 더 추가" : esc(hint)}</button>
      ${list.length > 1 ? `<div class="add-more"><span></span><button class="link plain" data-clear="${kind}">모두 제거</button></div>` : ""}
      ${kind === "prior" && list.length ? `<div class="hint">필요한 명세서: ${esc(s.needed_prior)} 말</div>` : ""}
    </div></div>`;
  };

  const years = p.years.map(y => `<option value="${y}" ${y === p.year ? "selected" : ""}>${y}</option>`).join("");
  const months = Array.from({length: 12}, (_, i) => i + 1).map(m => `<option value="${m}" ${m === p.month ? "selected" : ""}>${m}</option>`).join("");

  const form = `<div class="card" id="form">
    <div class="row"><span class="lbl">회계월</span><div class="grow">
      <div class="period-line"><span class="period">${esc(p.label)}</span><span class="det ${esc(p.detect)}">${p.detect === "auto" ? "● " : ""}${esc(p.detect_text)}</span>
        ${p.forced ? "" : `<button class="link" id="pick">${p.picker_open ? "접기" : "변경"}</button>`}</div>
      <div class="hint">${esc(p.summary)}</div>
      ${p.picker_open ? `<div class="picker"><select id="year" aria-label="연도">${years}</select>년 <select id="month" aria-label="월">${months}</select>월</div>` : ""}
    </div></div>
    ${fileRow("prior", "명세서", "전월 말 명세서 파일을 추가하세요", f.prior)}
    ${fileRow("douzone", "더존 원장", "당월 더존 원장 파일을 추가하세요", f.douzone)}
    ${f.pending.length ? `<div class="row"><span class="lbl"></span><div class="grow"><div class="chip-file pend"><span class="nm">분석 중 · ${esc(f.pending.join(", "))}</span></div></div></div>` : ""}
    <div class="row"><span class="lbl">계정코드</span><div class="grow">
      <input class="input" id="codes" value="${esc(s.account_codes)}" aria-label="미지급금 계정코드" spellcheck="false">
      <div class="hint">미지급금 · 여러 개면 쉼표(,)로 구분</div></div></div>
  </div>`;

  const recap = `<div class="card"><div class="recap"><span>${esc(s.recap)}</span><button class="link" id="edit">입력 수정</button></div></div>`;

  const runLabel = running ? "대사 중…" : "대사 시작";
  const buttons = `<div class="actions">
    <button class="btn ${done ? "ghost" : "primary"}" id="run" ${s.can_run ? "" : "disabled"}>${runLabel}</button>
    <button class="btn ${done ? "primary" : "ghost"}" id="export" ${s.can_export ? "" : "disabled"}>당월 명세서 생성</button></div>`;

  const b = s.banner;
  const status = `<div class="status ${esc(b.kind)}"><b>${esc(b.title)}</b><p>${esc(b.detail)}</p>
    ${(running || adding) ? '<div class="bar-run"><i></i></div>' : ""}</div>`;

  let body = "";
  if (!s.result) {
    body = `<div class="card guide"><h2>이렇게 진행합니다</h2>
      <div class="step"><div class="num">1</div><div><b>파일 추가</b><span>전월 말 명세서와 당월 더존 원장을 추가합니다. 회계월은 파일에서 자동으로 잡힙니다.</span></div></div>
      <div class="step"><div class="num">2</div><div><b>대사 시작</b><span>자동으로 짝을 맞추고, 사람이 봐야 할 항목만 추려 줍니다.</span></div></div>
      <div class="step"><div class="num">3</div><div><b>당월 명세서 생성</b><span>결과를 검토한 뒤 저장합니다. 전월 명세서 양식 그대로 나옵니다.</span></div></div></div>`;
  } else if (!running) {
    const r = s.result, sm = r.summary;
    const all = r.rows, nRev = all.filter(x => !x.matched).length, nAuto = all.length - nRev;
    const rows = view === "all" ? all : view === "matched" ? all.filter(x => x.matched) : all.filter(x => !x.matched);
    if (selected === null && rows.length) selected = rows[0].id;
    const sel = all.find(x => x.id === selected);
    const items = rows.map(x => `<button class="item ${x.id === selected ? "sel" : ""}" data-row="${x.id}">
        <div class="l1"><span class="v">${esc(x.vendor)}</span><span class="amt">${esc(x.amount)}</span></div>
        <div class="l2"><span class="chip ${x.matched ? "ok" : "rv"}">${x.matched ? "✓ " : ""}${esc(x.status)}</span><span class="why">${esc(x.reason)}</span></div></button>`).join("");
    const emptyText = view === "review" ? "검토할 명세서 항목이 없습니다." : view === "matched" ? "자동 대사된 항목이 없습니다." : "대사 결과가 없습니다.";
    body = `<div class="stats">
        <div class="stat"><small>자동 대사</small><strong class="${sm.matched ? "g" : "m"}">${sm.matched}</strong></div>
        <div class="stat"><small>검토 필요</small><strong class="${sm.review ? "a" : "m"}">${sm.review}</strong></div>
        <div class="stat"><small>입력 확인</small><strong class="${sm.issues ? "a" : "m"}">${sm.issues}</strong></div>
        <div class="stat"><small>당월 신규</small><strong>${sm.new}</strong></div></div>
      <div class="bar"><h2>결과 상세</h2><div style="display:flex;gap:8px;align-items:center">
        <button class="link plain" id="save">대사 결과 저장</button>
        <div class="seg" role="tablist"><button data-view="all" class="${view === "all" ? "on" : ""}">전체</button>
        <button data-view="review" class="${view === "review" ? "on" : ""}">검토 필요 ${nRev}</button>
        <button data-view="matched" class="${view === "matched" ? "on" : ""}">자동 대사 ${nAuto}</button></div></div></div>
      <div class="pre ${r.preflight.ok ? "ok" : "bad"}">${esc(r.preflight.text)}</div>
      <div class="card list">${items || `<div class="empty">${emptyText}</div>`}</div>
      <div class="card detail"><div class="h"><span>선택한 항목의 사유</span><button class="copy" id="copy">사유 복사</button></div>
        ${sel ? `<p>${esc(sel.vendor)} · ${esc(sel.amount)}원 · ${esc(sel.status)}\n${esc(sel.reason)}</p>${sel.source ? `<div class="src">원본 명세: ${esc(sel.source)}</div>` : ""}`
              : '<p class="hint">목록에서 항목을 선택하면 사유 전체가 여기에 표시됩니다.</p>'}</div>`;
  }

  app.innerHTML = `${DEV ? "" : `<div class="winbtns"><button class="wb" id="wmin" aria-label="최소화" title="최소화">&#8211;</button><button class="wb close" id="wclose" aria-label="닫기" title="닫기">&#10005;</button></div>`}
    <div class="dragbar" id="dragbar"></div>
    <div class="top"><div><div class="brand"><img class="logo" src="rabbit.png" alt=""><h1>명세서 대사</h1></div>
      <div class="sub">전월 말 명세서를 당월 더존 전표와 대사해 당월 말 명세서 초안을 만듭니다. 원본 파일은 수정하지 않습니다.</div></div>
      <div class="side"><div class="ver">${s.version ? (s.version === "개발 버전" ? "개발 버전" : "버전 " + esc(s.version)) : ""}</div>${updateBtn}${updErr}</div></div>
    ${done ? recap : form}${buttons}${status}${body}
    ${DEV ? "" : '<div class="grip grip-s" data-grip="s"></div><div class="grip grip-se" data-grip="se"></div>'}`;
  bind();
  if (keep) { const el = document.getElementById(keep); if (el && el.focus) el.focus(); }
}

function bind() {
  const $ = id => document.getElementById(id);
  document.querySelectorAll("[data-add]").forEach(el => {
    el.onclick = () => act("add_files", el.dataset.add);
    el.ondragover = e => { e.preventDefault(); el.classList.add("over"); };
    el.ondragleave = () => el.classList.remove("over");
    el.ondrop = e => {
      e.preventDefault(); el.classList.remove("over");
      const paths = [...(e.dataTransfer.files || [])].map(f => f.pywebviewFullPath).filter(Boolean);
      if (paths.length) act("add_paths", el.dataset.add, paths);
    };
  });
  document.querySelectorAll("[data-rm]").forEach(el => el.onclick = () => {
    const [k, i] = el.dataset.rm.split(":"); act("remove_file", k, Number(i));
  });
  document.querySelectorAll("[data-clear]").forEach(el => el.onclick = () => act("clear_files", el.dataset.clear));
  const startDrag = (mode) => e => { e.preventDefault(); call("begin_drag", mode); };
  if ($("dragbar")) $("dragbar").onpointerdown = startDrag("move");
  document.querySelectorAll("[data-grip]").forEach(el => el.onpointerdown = startDrag(el.dataset.grip));
  if ($("wmin")) $("wmin").onclick = () => call("minimize_window");
  if ($("wclose")) $("wclose").onclick = () => call("close_window");
  if ($("pick")) $("pick").onclick = () => act("toggle_picker");
  if ($("year")) { $("year").onchange = $("month").onchange = () => act("set_period", Number($("year").value), Number($("month").value)); }
  if ($("codes")) {
    $("codes").onchange = () => act("set_account_codes", $("codes").value);
    $("codes").onkeydown = e => { if (e.key === "Enter") $("codes").blur(); };
  }
  if ($("edit")) $("edit").onclick = () => act("open_inputs");
  if ($("run")) $("run").onclick = () => act("run_reconciliation");
  if ($("export")) $("export").onclick = () => act("export_month_end");
  if ($("save")) $("save").onclick = () => act("export_result");
  document.querySelectorAll("[data-view]").forEach(el => el.onclick = () => { view = el.dataset.view; selected = null; render(); });
  document.querySelectorAll("[data-row]").forEach(el => el.onclick = () => { selected = Number(el.dataset.row); render(); });
  if ($("copy")) $("copy").onclick = copyReason;
  if ($("upd")) $("upd").onclick = () => {
    const u = S.update;
    showModal({kind: "info", title: "업데이트",
      message: `현재 버전: ${u.current || "알 수 없음"}\n새 버전: ${u.date || "최신"}\n\n새 버전을 내려받아 프로그램을 다시 시작합니다.\n작업 중인 대사 결과는 사라지니, 필요하면 먼저 저장해 주세요.\n\n업데이트할까요?`},
      () => act("start_update"));
  };
}

function copyReason() {
  const sel = S.result && S.result.rows.find(x => x.id === selected);
  if (!sel) return;
  const text = `${sel.vendor} · ${sel.amount}원 · ${sel.status}\n${sel.reason}`;
  const t = document.createElement("textarea");
  t.value = text; document.body.appendChild(t); t.select();
  try { document.execCommand("copy"); } catch (e) {}
  t.remove();
}

/* ---------- start ---------- */
async function start() {
  await act("get_state");
  act("check_update");
  setInterval(() => { if (S && !S.busy.running && !S.busy.adding) act("check_update"); }, 30 * 60 * 1000);
}
if (DEV) start();
else window.addEventListener("pywebviewready", start);
