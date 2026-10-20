/* === Ayla Add-Exc — Chat principal === */

const MAX_ROUNDS = 25;
const MAX_TOOL_RESULT_CHARS = 30000;
const MAX_CONTEXT_CHARS = 8000;

let conversation = [];
let busy = false;
let abortController = null;
let attachedContext = null; // { label, data }
let pendingNotes = [];
let inExcel = false;

const $ = (id) => document.getElementById(id);
const msgContainer = $('messages');
const userInput = $('user-input');
const sendBtn = $('btn-send');

/* ===================== Markdown seguro ===================== */

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

if (typeof marked !== 'undefined') {
  marked.use({ gfm: true, breaks: true, renderer: { html: (token) => escapeHtml(token.text || token.raw || '') } });
}

function renderMarkdown(el, text) {
  el.innerHTML = typeof marked !== 'undefined' ? marked.parse(text || '') : escapeHtml(text || '');
  el.querySelectorAll('a').forEach((a) => {
    if (!/^https?:/i.test(a.getAttribute('href') || '')) a.removeAttribute('href');
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
  });
  el.querySelectorAll('pre code').forEach((b) => { try { hljs.highlightElement(b); } catch { /* sin resaltado */ } });
  linkifyCellRefs(el);
}

/* ===================== Citas de celdas ===================== */

const CELL_REF_RE = /(?:('(?:[^']|'')+'|[A-Za-zÀ-ÿ_][\wÀ-ÿ.]*)!)?(\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?)/g;

function isValidRef(ref) {
  return ref.replace(/\$/g, '').split(':').every((part) => {
    const m = /^([A-Z]{1,3})(\d+)$/.exec(part);
    return m && colToNum(m[1]) <= 16384 && parseInt(m[2], 10) <= 1048576;
  });
}

function linkifyCellRefs(root) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (n.parentElement.closest('pre, a, .cell-ref') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);

  for (const node of nodes) {
    const text = node.nodeValue;
    CELL_REF_RE.lastIndex = 0;
    let m;
    let last = 0;
    const frag = document.createDocumentFragment();
    while ((m = CELL_REF_RE.exec(text))) {
      const before = text[m.index - 1];
      const after = text[m.index + m[0].length];
      if ((before && /[\wÀ-ÿ$]/.test(before)) || (after && /[\w(]/.test(after)) || !isValidRef(m[2])) continue;
      frag.appendChild(document.createTextNode(text.slice(last, m.index)));
      const span = document.createElement('span');
      span.className = 'cell-ref';
      span.textContent = m[0];
      span.dataset.ref = m[0];
      span.title = 'Ir a ' + m[0];
      frag.appendChild(span);
      last = m.index + m[0].length;
    }
    if (last === 0) continue;
    frag.appendChild(document.createTextNode(text.slice(last)));
    node.parentNode.replaceChild(frag, node);
  }
}

msgContainer.addEventListener('click', async (e) => {
  const ref = e.target.closest('.cell-ref');
  if (!ref) return;
  try {
    await selectAddress(ref.dataset.ref);
  } catch (err) {
    ref.classList.add('invalid');
    ref.title = `No se pudo ir a ${ref.dataset.ref}: ${err.message}`;
  }
});

/* ===================== Contexto automático ===================== */

function cellText(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[\t\n\r]+/g, ' ').slice(0, 40);
}

function gridPreview(values, startRow, startCol) {
  const header = ['#', ...values[0].map((_, c) => numToCol(startCol + c))].join('\t');
  const rows = values.map((row, r) => [startRow + r, ...row.map(cellText)].join('\t'));
  return [header, ...rows].join('\n');
}

async function buildAutoContext() {
  if (!inExcel) return '';
  const lines = [];
  try {
    const ov = await getTool('get_workbook_overview').run({});
    lines.push(`Libro: ${ov.workbook || '(sin nombre)'} | Hoja activa: ${ov.activeSheet} | Selección: ${ov.selection}`);
    lines.push('Hojas:');
    for (const s of ov.sheets) {
      let line = `- ${s.name}${s.visible ? '' : ' (oculta)'}: ${s.usedRange ? `rango usado ${splitAddress(s.usedRange).ref} (${s.rows} filas × ${s.cols} columnas)` : 'vacía'}`;
      if (s.firstRows && s.firstRows.length) line += ` | primera fila: ${s.firstRows[0].map(cellText).join(' | ')}`;
      if (s.charts.length) line += ` | gráficos: ${s.charts.join(', ')}`;
      lines.push(line);
    }
    if (ov.tables.length) lines.push('Tablas: ' + ov.tables.map((t) => `${t.name} (${t.address})`).join(', '));
    if (ov.namedRanges.length) lines.push('Nombres definidos: ' + ov.namedRanges.slice(0, 30).map((n) => `${n.name}=${n.refersTo}`).join(', '));

    await Excel.run(async (ctx) => {
      const sel = ctx.workbook.getSelectedRange();
      sel.load('address, rowCount, columnCount');
      await ctx.sync();
      const rows = Math.min(sel.rowCount, 20);
      const cols = Math.min(sel.columnCount, 15);
      const part = sel.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
      part.load('values, formulas');
      await ctx.sync();
      const start = topLeft(sel.address);
      lines.push(`\nSelección actual ${sel.address} (${sel.rowCount}×${sel.columnCount})${sel.rowCount > rows || sel.columnCount > cols ? ', vista parcial' : ''}:`);
      lines.push(gridPreview(part.values, start.row, start.col));
      const f = sparseFormulas(part.formulas, start.row, start.col);
      if (Object.keys(f).length) lines.push('Fórmulas en la selección: ' + JSON.stringify(f));
    });
  } catch (err) {
    lines.push(`(No se pudo leer el contexto: ${err.message})`);
  }
  let text = lines.join('\n');
  if (text.length > MAX_CONTEXT_CHARS) text = text.slice(0, MAX_CONTEXT_CHARS) + '\n…(contexto recortado)';
  return text;
}

/* ===================== UI de mensajes ===================== */

function hideWelcome() {
  const w = $('welcome');
  if (w) w.style.display = 'none';
}

function scrollToBottom() {
  msgContainer.scrollTop = msgContainer.scrollHeight;
}

function addUserMessage(text) {
  hideWelcome();
  const el = document.createElement('div');
  el.className = 'msg user';
  el.textContent = text;
  msgContainer.appendChild(el);
  scrollToBottom();
}

function createTurn() {
  hideWelcome();
  const el = document.createElement('div');
  el.className = 'msg assistant turn';
  el.innerHTML = '<div class="msg-icon">✦</div><div class="turn-body"><div class="thinking">Pensando<span class="loading-dots"></span></div></div>';
  msgContainer.appendChild(el);
  const body = el.querySelector('.turn-body');
  let stepsGroup = null;

  const clearThinking = () => body.querySelector('.thinking')?.remove();

  return {
    el,
    textSegment() {
      const seg = document.createElement('div');
      seg.className = 'msg-text';
      let appended = false;
      return {
        update(text) {
          if (!text) return;
          clearThinking();
          if (!appended) { body.appendChild(seg); appended = true; stepsGroup = null; }
          seg.classList.add('streaming');
          renderMarkdown(seg, text);
          scrollToBottom();
        },
        finish() { seg.classList.remove('streaming'); },
      };
    },
    addStep(name, args) {
      clearThinking();
      if (!stepsGroup) {
        stepsGroup = document.createElement('div');
        stepsGroup.className = 'steps';
        body.appendChild(stepsGroup);
      }
      const step = document.createElement('div');
      step.className = 'step running';
      step.innerHTML = '<div class="step-row"><span class="step-icon"></span><span class="step-label"></span></div><pre class="step-details hidden"></pre>';
      step.querySelector('.step-label').textContent = toolLabel(name, args);
      const details = step.querySelector('.step-details');
      details.textContent = `${name}(${JSON.stringify(args, null, 2)})`;
      step.querySelector('.step-row').addEventListener('click', () => details.classList.toggle('hidden'));
      stepsGroup.appendChild(step);
      scrollToBottom();
      return {
        done(result) {
          const failed = result && result.error;
          step.className = `step ${failed ? 'error' : 'ok'}`;
          if (failed) step.querySelector('.step-label').textContent += ` — ${result.error}`;
          details.textContent += `\n→ ${JSON.stringify(result, null, 2).slice(0, 3000)}`;
          linkifyCellRefs(step.querySelector('.step-label'));
        },
      };
    },
    confirm(message) {
      clearThinking();
      return new Promise((resolve) => {
        const card = document.createElement('div');
        card.className = 'confirm-card';
        card.innerHTML = '<div class="confirm-text"></div><div class="confirm-actions"><button class="btn btn-secondary">Cancelar</button><button class="btn btn-danger">Permitir</button></div>';
        card.querySelector('.confirm-text').textContent = `⚠️ Ayla quiere: ${message}`;
        const [no, yes] = card.querySelectorAll('button');
        const finish = (ok) => {
          card.querySelector('.confirm-actions').remove();
          card.classList.add(ok ? 'accepted' : 'rejected');
          card.querySelector('.confirm-text').textContent += ok ? ' — permitido' : ' — cancelado';
          abortController?.signal.removeEventListener('abort', onAbort);
          resolve(ok);
        };
        const onAbort = () => finish(false);
        abortController?.signal.addEventListener('abort', onAbort);
        no.addEventListener('click', () => finish(false));
        yes.addEventListener('click', () => finish(true));
        body.appendChild(card);
        stepsGroup = null;
        scrollToBottom();
      });
    },
    note(text, cls = '') {
      clearThinking();
      const n = document.createElement('div');
      n.className = `turn-note ${cls}`;
      n.textContent = text;
      body.appendChild(n);
      scrollToBottom();
    },
    showChanges(changes) {
      clearThinking();
      if (!changes.length) return;
      const card = document.createElement('div');
      card.className = 'changes-card';
      const undoable = changes.filter((c) => c.undo);
      card.innerHTML = `<div class="changes-header"><span>✎ ${changes.length} ${changes.length === 1 ? 'cambio' : 'cambios'} en el libro</span></div><ul class="changes-list"></ul>`;
      const list = card.querySelector('.changes-list');
      for (const c of changes) {
        const li = document.createElement('li');
        li.textContent = c.label + (c.address ? ' · ' : '');
        if (c.address) {
          const ref = document.createElement('span');
          ref.className = 'cell-ref';
          ref.dataset.ref = c.address;
          ref.textContent = c.address;
          li.appendChild(ref);
        }
        if (!c.undo) li.classList.add('no-undo');
        list.appendChild(li);
      }
      if (undoable.length) {
        const btn = document.createElement('button');
        btn.className = 'btn btn-secondary undo-btn';
        btn.textContent = '↶ Deshacer';
        btn.title = undoable.length < changes.length ? 'Algunos cambios (formato, estructura) no se pueden deshacer desde Ayla' : 'Revertir estos cambios';
        btn.addEventListener('click', async () => {
          btn.disabled = true;
          btn.textContent = 'Deshaciendo…';
          const failures = [];
          for (const c of [...undoable].reverse()) {
            try { await c.undo(); } catch (err) { failures.push(`${c.label}: ${err.message}`); }
          }
          card.classList.add('reverted');
          btn.textContent = failures.length ? '⚠️ Deshecho parcialmente' : '✓ Cambios deshechos';
          if (failures.length) btn.title = failures.join('\n');
          pendingNotes.push(`El usuario deshizo los cambios de una respuesta anterior (${undoable.map((c) => c.address).filter(Boolean).join(', ')}).`);
        });
        card.querySelector('.changes-header').appendChild(btn);
      }
      body.appendChild(card);
      scrollToBottom();
    },
  };
}

function setBusy(active) {
  busy = active;
  sendBtn.classList.toggle('stop', active);
  sendBtn.textContent = active ? '■' : '↑';
  sendBtn.title = active ? 'Detener' : 'Enviar';
  sendBtn.disabled = !active && !userInput.value.trim();
}

/* ===================== Bucle principal ===================== */

function genId() {
  return `call_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
}

function truncateResult(result) {
  const s = JSON.stringify(result ?? null);
  return s.length > MAX_TOOL_RESULT_CHARS ? s.slice(0, MAX_TOOL_RESULT_CHARS) + '…(resultado recortado; pide rangos más pequeños)' : s;
}

async function runToolWithConfirmation(turn, name, args) {
  const question = SETTINGS.confirmDestructive !== false ? toolConfirmation(name, args) : null;
  if (question && !(await turn.confirm(question))) {
    return { error: 'El usuario canceló esta acción. No la reintentes salvo que te lo pida.' };
  }
  return callExcelTool(name, args);
}

async function sendMessage(textOverride) {
  if (busy) return;
  const text = (textOverride ?? userInput.value).trim();
  if (!text) return;

  if (!SETTINGS.apiKey || !SETTINGS.endpoint) {
    window.openSettings?.();
    return;
  }

  userInput.value = '';
  adjustInputHeight();
  addUserMessage(text);
  const turn = createTurn();
  abortController = new AbortController();
  const { signal } = abortController;
  setBusy(true);

  const parts = [];
  const autoContext = await buildAutoContext();
  if (autoContext) parts.push(`<contexto_excel>\n${autoContext}\n</contexto_excel>`);
  if (attachedContext) {
    parts.push(`<datos_adjuntos descripcion="${attachedContext.label}">\n${truncateResult(attachedContext.data)}\n</datos_adjuntos>`);
    clearAttachment();
  }
  if (pendingNotes.length) {
    parts.push(`<nota>${pendingNotes.join(' ')}</nota>`);
    pendingNotes = [];
  }
  parts.push(text);
  conversation.push({ role: 'user', content: parts.join('\n\n') });

  ChangeTracker.begin();
  let segment = null;
  let partial = '';
  try {
    let finished = false;
    for (let round = 0; round < MAX_ROUNDS && !finished; round++) {
      segment = turn.textSegment();
      partial = '';
      const { text: reply, toolCalls } = await LLM.chat({
        settings: SETTINGS,
        system: getSystemPrompt(),
        messages: conversation,
        tools: EXCEL_TOOLS_DEFINITIONS,
        signal,
        onText: (t) => { partial = t; segment.update(t); },
      });
      segment.finish();

      if (!toolCalls.length) {
        conversation.push({ role: 'assistant', content: reply || '' });
        if (!reply) turn.note('(sin respuesta del modelo)');
        finished = true;
        break;
      }

      const calls = toolCalls.map((tc) => {
        const call = { id: tc.id || genId(), type: 'function', function: { name: tc.name, arguments: tc.arguments || '{}' } };
        if (tc.extra_content) call.extra_content = tc.extra_content;
        return call;
      });
      conversation.push({ role: 'assistant', content: reply || null, tool_calls: calls });

      for (const call of calls) {
        let result;
        let args = {};
        try {
          args = JSON.parse(call.function.arguments || '{}') || {};
        } catch {
          result = { error: 'Argumentos JSON inválidos.' };
        }
        if (signal.aborted) {
          result = { error: 'Cancelado por el usuario.' };
        } else if (!result) {
          const step = turn.addStep(call.function.name, args);
          result = await runToolWithConfirmation(turn, call.function.name, args);
          step.done(result);
        }
        conversation.push({ role: 'tool', tool_call_id: call.id, content: truncateResult(result) });
      }
      if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
    }
    if (!finished) turn.note('Se alcanzó el límite de pasos. Escribe "continúa" para seguir.', 'warn');
  } catch (err) {
    segment?.finish();
    const last = conversation[conversation.length - 1];
    if (last.role !== 'tool' && last.role !== 'assistant') {
      conversation.push({ role: 'assistant', content: partial || '(respuesta interrumpida)' });
    } else if (last.role === 'tool') {
      conversation.push({ role: 'assistant', content: '(respuesta interrumpida)' });
    }
    if (err.name === 'AbortError') turn.note('Detenido.', 'warn');
    else turn.note(`⚠️ ${err.message}`, 'error');
  } finally {
    turn.showChanges(ChangeTracker.end());
    abortController = null;
    setBusy(false);
    userInput.focus();
  }
}

/* ===================== Adjuntos ===================== */

function clearAttachment() {
  attachedContext = null;
  $('context-badge').classList.add('hidden');
}

async function attachContext(mode) {
  try {
    let data;
    let label;
    if (mode === 'selection') {
      data = await getTool('get_selected_range').run({});
      label = `Selección ${data.address} (${data.rows}×${data.cols})`;
    } else {
      data = await getTool('get_worksheet_data').run({});
      label = `Hoja ${data.sheet} (${data.rows || 0}×${data.cols || 0})`;
    }
    attachedContext = { label, data };
    $('context-label').textContent = `📎 ${label}`;
    $('context-badge').classList.remove('hidden');
  } catch (err) {
    $('context-label').textContent = `⚠️ No se pudo adjuntar: ${err.message}`;
    $('context-badge').classList.remove('hidden');
  }
}

/* ===================== Init ===================== */

function adjustInputHeight() {
  userInput.style.height = 'auto';
  userInput.style.height = Math.min(userInput.scrollHeight, 140) + 'px';
}

function newConversation() {
  if (busy) abortController?.abort();
  conversation = [];
  pendingNotes = [];
  clearAttachment();
  msgContainer.querySelectorAll('.msg').forEach((m) => m.remove());
  $('welcome').style.display = '';
}

function updateWelcomeConfig() {
  $('welcome-config').classList.toggle('hidden', !!(SETTINGS.apiKey && SETTINGS.endpoint));
}

document.addEventListener('DOMContentLoaded', () => {
  userInput.addEventListener('input', () => {
    adjustInputHeight();
    if (!busy) sendBtn.disabled = !userInput.value.trim();
  });
  userInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
      e.preventDefault();
      sendMessage();
    }
  });
  sendBtn.addEventListener('click', () => (busy ? abortController?.abort() : sendMessage()));
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && busy) abortController?.abort();
  });

  document.querySelectorAll('.suggestion').forEach((b) => b.addEventListener('click', () => sendMessage(b.textContent)));

  const attachMenu = $('attach-menu');
  $('btn-attach').addEventListener('click', (e) => {
    e.stopPropagation();
    attachMenu.classList.toggle('hidden');
  });
  document.addEventListener('click', () => attachMenu.classList.add('hidden'));
  document.querySelectorAll('.attach-option').forEach((btn) => btn.addEventListener('click', () => {
    attachMenu.classList.add('hidden');
    attachContext(btn.dataset.mode);
  }));
  $('btn-remove-context').addEventListener('click', clearAttachment);

  $('btn-clear').addEventListener('click', newConversation);
  $('config-link').addEventListener('click', (e) => { e.preventDefault(); window.openSettings?.(); });
  $('btn-save-settings').addEventListener('click', () => setTimeout(updateWelcomeConfig, 0));
  updateWelcomeConfig();
});

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    inExcel = info && info.host === Office.HostType.Excel;
    $('host-banner').classList.toggle('hidden', inExcel);
    if (!localStorage.getItem('ayla_settings')) {
      loadSettings();
      applySettingsToUI();
      updateConnectionBar();
      updateWelcomeConfig();
    }
  });
} else {
  document.addEventListener('DOMContentLoaded', () => $('host-banner').classList.remove('hidden'));
}
