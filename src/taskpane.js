/* === Ayla Add-Exc — Chat Principal === */

let messages = [];
let isStreaming = false;
let currentContext = null; // { mode: string, data: any }

/* --- DOM refs --- */
const $ = id => document.getElementById(id);
const msgContainer = $('messages');
const userInput = $('user-input');
const sendBtn = $('btn-send');
const attachBtn = $('btn-attach');
const attachMenu = $('attach-menu');
const contextBadge = $('context-badge');
const contextLabel = $('context-label');

/* --- Init --- */
document.addEventListener('DOMContentLoaded', () => {
  // Adjust textarea height
  userInput.addEventListener('input', adjustInputHeight);
  userInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
    // Enable/disable send
    sendBtn.disabled = !e.target.value.trim();
  });
  userInput.addEventListener('input', () => {
    sendBtn.disabled = !userInput.value.trim();
  });

  sendBtn.addEventListener('click', sendMessage);

  // Attach menu
  attachBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    attachMenu.classList.toggle('hidden');
  });
  document.addEventListener('click', () => attachMenu.classList.add('hidden'));

  document.querySelectorAll('.attach-option').forEach(btn => {
    btn.addEventListener('click', async () => {
      attachMenu.classList.add('hidden');
      const mode = btn.dataset.mode;
      await attachContext(mode);
    });
  });

  // Clear chat
  $('btn-clear').addEventListener('click', () => {
    if (messages.length === 0) return;
    if (confirm('¿Limpiar toda la conversación?')) {
      messages = [];
      msgContainer.innerHTML = '';
      showWelcome();
    }
  });

  // Remove context
  $('btn-remove-context').addEventListener('click', () => {
    currentContext = null;
    contextBadge.classList.add('hidden');
  });
});

function adjustInputHeight() {
  userInput.style.height = 'auto';
  userInput.style.height = Math.min(userInput.scrollHeight, 120) + 'px';
}

/* --- Welcome --- */
function showWelcome() {
  const welcome = document.querySelector('.welcome');
  if (welcome) {
    welcome.style.display = '';
    return;
  }
  // Recreate if it was cleared
  // (not needed if we never remove it)
}

/* --- Context Attachment --- */
async function attachContext(mode) {
  if (mode === 'none') {
    currentContext = null;
    contextBadge.classList.add('hidden');
    return;
  }

  try {
    let data;
    switch (mode) {
      case 'selection':
        data = await ExcelTools.get_selected_range({});
        contextLabel.textContent = `📌 Rango seleccionado: ${data.address} (${data.rows}x${data.cols})`;
        break;
      case 'sheet':
        data = await ExcelTools.get_worksheet_data({});
        contextLabel.textContent = `📄 Hoja: ${data.sheet} — ${data.rows} filas x ${data.cols} columnas`;
        break;
      case 'all-sheets':
        const names = await ExcelTools.get_worksheet_names();
        contextLabel.textContent = `📚 ${names.length} hojas: ${names.map(n => n.name).join(', ')}`;
        data = { sheets: names };
        break;
    }
    currentContext = { mode, data };
    contextBadge.classList.remove('hidden');
  } catch (err) {
    addMessage('tool', `⚠️ Error al leer contexto: ${err.message}`);
  }
}

/* --- Add Message to UI --- */
function addMessage(role, content, extra = {}) {
  // Remove welcome if present
  const welcome = document.querySelector('.welcome');
  if (welcome) welcome.style.display = 'none';

  // Check if last message is streaming — replace it
  let msgEl;
  if (extra.streaming) {
    const last = msgContainer.lastElementChild;
    if (last && last.dataset.role === role && last.classList.contains('streaming')) {
      msgEl = last;
      msgEl.querySelector('.msg-text').innerHTML = '';
    }
  }

  if (!msgEl) {
    msgEl = document.createElement('div');
    msgEl.className = `msg ${role}`;
    if (extra.streaming) msgEl.classList.add('streaming');
    msgEl.dataset.role = role;

    if (role === 'assistant') {
      const icon = document.createElement('div');
      icon.className = 'msg-icon';
      icon.textContent = '✦';
      msgEl.appendChild(icon);
    }

    const textDiv = document.createElement('div');
    textDiv.className = 'msg-text';
    msgEl.appendChild(textDiv);
    msgContainer.appendChild(msgEl);
  }

  const textDiv = msgEl.querySelector('.msg-text');

  if (role === 'assistant' && !extra.noMarkdown) {
    // If streaming, append text
    if (extra.append) {
      textDiv.innerHTML = marked.parse(textDiv.textContent + content);
    } else {
      textDiv.innerHTML = marked.parse(content);
    }
    // Highlight code blocks
    textDiv.querySelectorAll('pre code').forEach(block => {
      hljs.highlightElement(block);
    });
  } else {
    if (extra.append) {
      textDiv.textContent += content;
    } else {
      textDiv.textContent = content;
    }
  }

  // Scroll to bottom
  msgContainer.scrollTop = msgContainer.scrollHeight;

  return msgEl;
}

function setStreaming(active) {
  isStreaming = active;
  sendBtn.disabled = active || !userInput.value.trim();
  userInput.disabled = active;
  if (!active) {
    document.querySelectorAll('.msg.streaming').forEach(el => el.classList.remove('streaming'));
  }
}

function showToolCall(name, args) {
  const argsStr = typeof args === 'object' ? JSON.stringify(args).slice(0, 200) : String(args).slice(0, 200);
  addMessage('tool', `🔧 Ejecutando: ${name}(${argsStr})`);
}

function showToolResult(name, result) {
  const lastTool = msgContainer.querySelector('.msg.tool:last-child');
  if (lastTool) {
    const text = lastTool.querySelector('.msg-text').textContent;
    const status = result.error ? '⚠️' : '✅';
    lastTool.querySelector('.msg-text').textContent = `${status} ${name} → ${result.error ? result.error : 'OK'}`;
  }
}

/* --- Send Message --- */
async function sendMessage() {
  const text = userInput.value.trim();
  if (!text || isStreaming) return;

  userInput.value = '';
  userInput.style.height = 'auto';
  sendBtn.disabled = true;

  // Check if configured
  if (!SETTINGS.apiKey || !SETTINGS.endpoint) {
    addMessage('assistant', '⚠️ Primero configura la conexión en ⚙️. Necesitas un endpoint y API key.');
    return;
  }

  // Add user message
  addMessage('user', text);

  // Build context string
  let contextStr = '';
  if (currentContext && currentContext.data) {
    const d = currentContext.data;
    if (d.address) contextStr += `\n[Rango seleccionado: ${d.address}]\n`;
    if (d.values) {
      const preview = d.values.slice(0, 20).map(r =>
        r.map(v => v === null || v === undefined ? '' : String(v).slice(0, 50)).join('\t')
      ).join('\n');
      contextStr += `\`\`\`\n${preview}\n\`\`\`\n`;
      if (d.values.length > 20) contextStr += `... y ${d.values.length - 20} filas más\n`;
    }
    if (d.sheets) {
      contextStr += `\nHojas: ${d.sheets.map(s => s.name).join(', ')}\n`;
    }
    if (d.sheet) contextStr += `\nHoja activa: ${d.sheet}\n`;
  }

  // Build user message for API
  const userMsgContent = contextStr
    ? `Contexto de Excel:\n${contextStr}\n\nInstrucción del usuario: ${text}`
    : text;

  messages.push({ role: 'user', content: userMsgContent });

  // Prepare API call
  const url = getEndpointUrl();
  const apiKey = SETTINGS.apiKey;

  try {
    setStreaming(true);
    const msgEl = addMessage('assistant', '**...**', { streaming: true });

    let fullResponse = '';
    let assistantMessage = { role: 'assistant', content: '' };

    // Main interaction loop (handle multiple tool rounds)
    let rounds = 0;
    const MAX_ROUNDS = 8;

    while (rounds < MAX_ROUNDS) {
      rounds++;

      const requestBody = {
        model: SETTINGS.model,
        messages: [
          { role: 'system', content: SETTINGS.systemPrompt || undefined },
          ...messages,
          ...(assistantMessage.content ? [assistantMessage] : []),
        ].filter(m => m.content),
        tools: EXCEL_TOOLS_DEFINITIONS,
        tool_choice: 'auto',
        stream: true,
        max_tokens: 4096,
      };

      let responseContent = '';
      let toolCalls = [];

      const resp = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
        },
        body: JSON.stringify(requestBody),
      });

      if (!resp.ok) {
        const errBody = await resp.text().catch(() => '');
        let errMsg = `Error ${resp.status}: ${resp.statusText}`;
        try {
          const errJson = JSON.parse(errBody);
          errMsg = errJson.error?.message || errMsg;
        } catch {}
        throw new Error(errMsg);
      }

      // Read streaming response
      const reader = resp.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || trimmed === 'data: [DONE]') continue;
          if (!trimmed.startsWith('data: ')) continue;

          try {
            const json = JSON.parse(trimmed.slice(6));

            if (json.choices?.[0]?.delta?.content) {
              const delta = json.choices[0].delta.content;
              responseContent += delta;
              // Update streaming UI
              msgEl.querySelector('.msg-text').innerHTML = marked.parse(responseContent + '▊');
              msgEl.querySelectorAll('pre code').forEach(b => hljs.highlightElement(b));
              msgContainer.scrollTop = msgContainer.scrollHeight;
            }

            if (json.choices?.[0]?.delta?.tool_calls) {
              for (const tc of json.choices[0].delta.tool_calls) {
                const idx = tc.index || 0;
                if (!toolCalls[idx]) {
                  toolCalls[idx] = { id: tc.id || '', function: { name: '', arguments: '' } };
                }
                if (tc.id) toolCalls[idx].id = tc.id;
                if (tc.function?.name) toolCalls[idx].function.name += tc.function.name;
                if (tc.function?.arguments) toolCalls[idx].function.arguments += tc.function.arguments;
              }
            }
          } catch {}
        }
      }

      // Final flush of buffer
      if (buffer.trim() && !buffer.includes('[DONE]')) {
        try {
          const line = buffer.trim();
          if (line.startsWith('data: ')) {
            const json = JSON.parse(line.slice(6));
            if (json.choices?.[0]?.delta?.content) {
              responseContent += json.choices[0].delta.content;
            }
          }
        } catch {}
      }

      // Update the assistant message display
      if (responseContent) {
        fullResponse += responseContent;
        msgEl.querySelector('.msg-text').innerHTML = marked.parse(fullResponse);
        msgEl.querySelectorAll('pre code').forEach(b => hljs.highlightElement(b));
        msgEl.classList.remove('streaming');
        msgContainer.scrollTop = msgContainer.scrollHeight;
      }

      // Process tool calls
      toolCalls = toolCalls.filter(tc => tc.function?.name);
      if (toolCalls.length === 0) {
        // No more tool calls — this is the final response
        assistantMessage.content = responseContent;
        assistantMessage.tool_calls = undefined;
        break;
      }

      // Add assistant message with tool calls to history
      const asstMsg = {
        role: 'assistant',
        content: responseContent || null,
        tool_calls: toolCalls.map(tc => ({
          id: tc.id || `call_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
          type: 'function',
          function: { name: tc.function.name, arguments: tc.function.arguments },
        })),
      };
      messages.push(asstMsg);
      assistantMessage = { role: 'assistant', content: '' };

      // Execute each tool call
      for (const tc of asstMsg.tool_calls) {
        let args = {};
        try {
          args = JSON.parse(tc.function.arguments || '{}');
        } catch {}
        showToolCall(tc.function.name, args);
        const result = await callExcelTool(tc.function.name, args);
        showToolResult(tc.function.name, result);
        messages.push({
          role: 'tool',
          tool_call_id: tc.id,
          content: JSON.stringify(result),
        });
      }
    }

    if (rounds >= MAX_ROUNDS) {
      addMessage('assistant', '⚠️ Se alcanzó el límite de operaciones. Puedes continuar preguntando.');
    }

    // Update final response
    if (fullResponse) {
      // Already rendered via streaming
    } else {
      msgEl.querySelector('.msg-text').innerHTML = '(sin respuesta)';
    }
    msgEl.classList.remove('streaming');

  } catch (err) {
    // Remove streaming indicator
    const streaming = msgContainer.querySelector('.msg.streaming');
    if (streaming) {
      streaming.querySelector('.msg-text').textContent = `⚠️ Error: ${err.message}`;
      streaming.classList.remove('streaming');
    } else {
      addMessage('assistant', `⚠️ Error: ${err.message}`);
    }
  } finally {
    setStreaming(false);
    messages.push(assistantMessage);
  }
}

/* --- Keyboard shortcut: Ctrl+Shift+A to focus --- */
document.addEventListener('keydown', (e) => {
  if (e.ctrlKey && e.shiftKey && e.key === 'A') {
    e.preventDefault();
    userInput.focus();
  }
});
