/* === Ayla Add-Exc — Proveedores de IA === */
/* Historial interno en formato OpenAI; se convierte a Anthropic cuando hace falta. */

const LLM = (() => {
  function detectProvider(endpoint, provider) {
    if (provider === 'anthropic' || provider === 'openai') return provider;
    return /anthropic\.com/i.test(endpoint || '') ? 'anthropic' : 'openai';
  }

  function isFreeLLMAPI(endpoint) {
    return /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\]):(3001|3939|31415)(\/|$)/i.test(String(endpoint || '').trim());
  }

  function isLocal(endpoint) {
    return /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(:\d+)?(\/|$)/i.test(String(endpoint || '').trim());
  }

  /** fetch() solo lanza TypeError ante fallos de red/CORS: añade una pista útil. */
  async function doFetch(url, init) {
    try {
      return await fetch(url, init);
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      let hint = 'Revisa el endpoint y tu conexión a internet.';
      if (isLocal(url)) {
        const origin = typeof location !== 'undefined' ? location.origin : 'el origen del complemento';
        hint = `¿Está corriendo el servidor local (${new URL(url).host})? Si es FreeLLMAPI, debe permitir el origen del complemento: añade DASHBOARD_ORIGINS=${origin} a su archivo .env y reinícialo.`;
      }
      throw new Error(`No se pudo conectar con ${new URL(url).host}. ${hint}`);
    }
  }

  function joinUrl(base, suffix) {
    const url = String(base || '').trim().replace(/\/+$/, '');
    return url.endsWith(suffix) ? url : url + suffix;
  }

  function safeParse(s) {
    if (s && typeof s === 'object') return s;
    try { return JSON.parse(s || '{}') || {}; } catch { return {}; }
  }

  async function errorFromResponse(resp) {
    const body = await resp.text().catch(() => '');
    let msg = `Error ${resp.status}${resp.statusText ? ': ' + resp.statusText : ''}`;
    try {
      const j = JSON.parse(body);
      const e = Array.isArray(j) ? j[0]?.error : j.error;
      if (e?.message) msg += ` — ${e.message}`;
    } catch {
      if (body) msg += ` — ${body.slice(0, 200)}`;
    }
    const err = new Error(msg);
    err.status = resp.status;
    return err;
  }

  async function fetchWithRetry(url, init, retries = 2) {
    for (let attempt = 0; ; attempt++) {
      const resp = await doFetch(url, init);
      if (resp.ok) return resp;
      if (attempt < retries && [429, 500, 502, 503, 529].includes(resp.status)) {
        await new Promise((r) => setTimeout(r, 1500 * (attempt + 1)));
        if (init.signal?.aborted) throw new DOMException('Aborted', 'AbortError');
        continue;
      }
      throw await errorFromResponse(resp);
    }
  }

  /** Lee un stream SSE y produce cada payload "data:" como objeto JSON. */
  async function* readSSE(resp) {
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    const parse = function* (line) {
      const t = line.trim();
      if (!t.startsWith('data:')) return;
      const data = t.slice(5).trim();
      if (!data || data === '[DONE]') return;
      try { yield JSON.parse(data); } catch { /* fragmento inválido */ }
    };
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';
      for (const line of lines) yield* parse(line);
    }
    if (buffer) yield* parse(buffer);
  }

  /* ---------------- OpenAI-compatible (OpenAI, Gemini, OpenRouter, Ollama…) ---------------- */

  async function chatOpenAI({ settings, system, messages, tools, signal, onText }) {
    const resp = await fetchWithRetry(joinUrl(settings.endpoint, '/chat/completions'), {
      method: 'POST',
      signal,
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${settings.apiKey}` },
      body: JSON.stringify({
        model: settings.model,
        messages: [{ role: 'system', content: system }, ...messages],
        tools: tools && tools.length ? tools : undefined,
        tool_choice: tools && tools.length ? 'auto' : undefined,
        stream: true,
      }),
    });

    let text = '';
    const calls = [];
    for await (const json of readSSE(resp)) {
      if (json.error) throw new Error(json.error.message || JSON.stringify(json.error));
      const delta = json.choices?.[0]?.delta;
      if (!delta) continue;
      if (delta.content) {
        text += delta.content;
        onText?.(text);
      }
      for (const tc of delta.tool_calls || []) {
        // Algunos proveedores (p. ej. Gemini) no envían "index": se agrupa por id.
        let call;
        if (typeof tc.index === 'number') {
          call = calls[tc.index] || (calls[tc.index] = { id: '', name: '', arguments: '' });
        } else if (tc.id) {
          call = calls.find((c) => c && c.id === tc.id);
          if (!call) calls.push((call = { id: tc.id, name: '', arguments: '' }));
        } else {
          call = calls[calls.length - 1] || (calls[0] = { id: '', name: '', arguments: '' });
        }
        if (tc.id) call.id = tc.id;
        if (tc.function?.name) call.name += tc.function.name;
        if (tc.function?.arguments) {
          call.arguments += typeof tc.function.arguments === 'string' ? tc.function.arguments : JSON.stringify(tc.function.arguments);
        }
        if (tc.extra_content) call.extra_content = tc.extra_content;
      }
    }
    return { text, toolCalls: calls.filter((c) => c && c.name) };
  }

  /* ---------------- Anthropic (Claude) ---------------- */

  function toAnthropicMessages(messages) {
    const out = [];
    const push = (role, blocks) => {
      if (!blocks.length) return;
      const last = out[out.length - 1];
      if (last && last.role === role) last.content.push(...blocks);
      else out.push({ role, content: [...blocks] });
    };
    for (const m of messages) {
      if (m.role === 'user') {
        push('user', [{ type: 'text', text: m.content || '…' }]);
      } else if (m.role === 'assistant') {
        const blocks = [];
        if (m.content) blocks.push({ type: 'text', text: m.content });
        for (const tc of m.tool_calls || []) {
          blocks.push({ type: 'tool_use', id: tc.id, name: tc.function.name, input: safeParse(tc.function.arguments) });
        }
        push('assistant', blocks);
      } else if (m.role === 'tool') {
        push('user', [{ type: 'tool_result', tool_use_id: m.tool_call_id, content: m.content }]);
      }
    }
    return out;
  }

  async function chatAnthropic({ settings, system, messages, tools, signal, onText, maxTokens }) {
    const resp = await fetchWithRetry(joinUrl(settings.endpoint, '/messages'), {
      method: 'POST',
      signal,
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': settings.apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: settings.model,
        max_tokens: maxTokens,
        system,
        messages: toAnthropicMessages(messages),
        tools: (tools || []).map((t) => ({ name: t.function.name, description: t.function.description, input_schema: t.function.parameters })),
        stream: true,
      }),
    });

    let text = '';
    const blocks = {};
    for await (const ev of readSSE(resp)) {
      if (ev.type === 'error') throw new Error(ev.error?.message || 'Error de Anthropic');
      if (ev.type === 'content_block_start' && ev.content_block?.type === 'tool_use') {
        blocks[ev.index] = { id: ev.content_block.id, name: ev.content_block.name, arguments: '' };
      } else if (ev.type === 'content_block_delta') {
        if (ev.delta?.type === 'text_delta') {
          text += ev.delta.text;
          onText?.(text);
        } else if (ev.delta?.type === 'input_json_delta' && blocks[ev.index]) {
          blocks[ev.index].arguments += ev.delta.partial_json || '';
        }
      }
    }
    const toolCalls = Object.keys(blocks).sort((a, b) => a - b).map((k) => blocks[k]);
    toolCalls.forEach((c) => { if (!c.arguments) c.arguments = '{}'; });
    return { text, toolCalls };
  }

  /**
   * Envía la conversación al modelo.
   * @returns {Promise<{text: string, toolCalls: {id, name, arguments, extra_content?}[]}>}
   */
  async function chat(opts) {
    const provider = detectProvider(opts.settings.endpoint, opts.settings.provider);
    const args = { maxTokens: 8192, ...opts };
    return provider === 'anthropic' ? chatAnthropic(args) : chatOpenAI(args);
  }

  /** Petición mínima sin streaming para "Probar conexión". */
  async function ping(settings) {
    const provider = detectProvider(settings.endpoint, settings.provider);
    const prompt = 'Responde solo "ok".';
    if (provider === 'anthropic') {
      const resp = await fetchWithRetry(joinUrl(settings.endpoint, '/messages'), {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': settings.apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({ model: settings.model, max_tokens: 16, messages: [{ role: 'user', content: prompt }] }),
      }, 0);
      const data = await resp.json();
      return { provider, reply: data.content?.find((b) => b.type === 'text')?.text || '(vacío)' };
    }
    const resp = await fetchWithRetry(joinUrl(settings.endpoint, '/chat/completions'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${settings.apiKey}` },
      body: JSON.stringify({ model: settings.model, messages: [{ role: 'user', content: prompt }] }),
    }, 0);
    const data = await resp.json();
    return { provider, reply: data.choices?.[0]?.message?.content || '(vacío)', servedBy: data.model && data.model !== settings.model ? data.model : '' };
  }

  /** Lista de modelos (GET /models), usada para sugerir nombres en la configuración. */
  async function listModels(settings) {
    if (detectProvider(settings.endpoint, settings.provider) === 'anthropic') return [];
    const resp = await doFetch(joinUrl(settings.endpoint, '/models'), { headers: { Authorization: `Bearer ${settings.apiKey}` } });
    if (!resp.ok) return [];
    const data = await resp.json();
    return (data.data || []).map((m) => m.id).filter(Boolean);
  }

  return { chat, ping, listModels, detectProvider, isFreeLLMAPI, toAnthropicMessages };
})();
