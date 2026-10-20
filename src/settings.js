/* === Ayla Add-Exc — Configuración === */

const BASE_SYSTEM_PROMPT = `Eres Ayla, un asistente de IA integrado en Microsoft Excel. Trabajas directamente sobre el libro abierto del usuario mediante herramientas.

CÓMO TRABAJAS
- Cada mensaje del usuario llega con un bloque <contexto_excel> que resume el libro (hojas, rangos usados, encabezados, tablas) y la selección actual. Úsalo para orientarte; para ver datos concretos llama a read_range u otras herramientas. Nunca inventes valores ni digas que no tienes acceso a la hoja.
- Si el usuario dice "esto", "aquí" o "estos datos", se refiere a la selección actual.
- Para explicar o auditar una fórmula usa get_cell_dependencies (sigue la cadena de precedentes si hace falta). Para depurar errores (#REF!, #DIV/0!, #N/A…) usa find_errors y luego analiza las causas.
- Al construir o modificar modelos, prefiere fórmulas en lugar de valores fijos para que el libro siga recalculando. Conserva la estructura y el formato existentes.
- Las fórmulas que escribas deben ir en inglés y con coma como separador (SUM, AVERAGE, IF, VLOOKUP, XLOOKUP…): Excel las mostrará en el idioma del usuario.
- No sobrescribas datos existentes sin necesidad; si tienes que hacerlo, dilo. Las acciones destructivas piden confirmación al usuario automáticamente.

CÓMO RESPONDES
- Cita SIEMPRE las celdas en las que te basas o que modificaste usando referencias de Excel, por ejemplo B5, C2:C20, Ventas!D10 o 'Flujo de caja'!B3. Se convierten en enlaces que llevan al usuario a esa celda.
- Al terminar cambios, resume brevemente qué hiciste y dónde.
- Sé claro y conciso. Responde en el idioma del usuario (español por defecto). Usa Markdown (listas, tablas cortas) cuando ayude.`;

const PRESETS = {
  freellmapi: { endpoint: 'http://localhost:3001/v1', model: 'auto', provider: 'openai' },
  claude: { endpoint: 'https://api.anthropic.com/v1', model: 'claude-sonnet-4-5', provider: 'anthropic' },
  gemini: { endpoint: 'https://generativelanguage.googleapis.com/v1beta/openai/', model: 'gemini-2.5-flash', provider: 'openai' },
  openai: { endpoint: 'https://api.openai.com/v1', model: 'gpt-4o-mini', provider: 'openai' },
  openrouter: { endpoint: 'https://openrouter.ai/api/v1', model: 'anthropic/claude-sonnet-4.5', provider: 'openai' },
};

const SETTINGS = {
  endpoint: PRESETS.freellmapi.endpoint,
  model: PRESETS.freellmapi.model,
  provider: 'openai',
  apiKey: '',
  systemPrompt: '',
  confirmDestructive: true,
};

function getSystemPrompt() {
  const extra = (SETTINGS.systemPrompt || '').trim();
  return extra ? `${BASE_SYSTEM_PROMPT}\n\nINSTRUCCIONES ADICIONALES DEL USUARIO\n${extra}` : BASE_SYSTEM_PROMPT;
}

/* --- persistencia --- */

function readSettingsForm() {
  return {
    endpoint: document.getElementById('endpoint-url').value.trim(),
    model: document.getElementById('model-name').value.trim(),
    provider: document.getElementById('provider').value,
    apiKey: document.getElementById('api-key').value.trim(),
    systemPrompt: document.getElementById('system-prompt').value.trim(),
    confirmDestructive: document.getElementById('confirm-destructive').checked,
  };
}

function saveSettings() {
  try {
    const data = readSettingsForm();
    localStorage.setItem('ayla_settings', JSON.stringify(data));
    if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
      Office.context.roamingSettings.set('ayla_settings', JSON.stringify(data));
      Office.context.roamingSettings.saveAsync();
    }
    Object.assign(SETTINGS, data);
    updateConnectionBar();
    return true;
  } catch (e) {
    console.error('Error guardando configuración:', e);
    return false;
  }
}

function loadSettings() {
  try {
    let raw = localStorage.getItem('ayla_settings');
    if (!raw && typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
      raw = Office.context.roamingSettings.get('ayla_settings');
    }
    if (raw) {
      const data = JSON.parse(raw);
      // v1 guardaba el prompt base completo; ahora solo se guardan instrucciones adicionales.
      if (data.systemPrompt && data.systemPrompt.startsWith('Eres un asistente experto en Excel. Tu función')) {
        data.systemPrompt = '';
      }
      Object.assign(SETTINGS, data);
    }
  } catch (e) {
    console.warn('No se pudo cargar la configuración:', e);
  }
}

function applySettingsToUI() {
  document.getElementById('endpoint-url').value = SETTINGS.endpoint || '';
  document.getElementById('model-name').value = SETTINGS.model || '';
  document.getElementById('provider').value = SETTINGS.provider || 'auto';
  document.getElementById('api-key').value = SETTINGS.apiKey || '';
  document.getElementById('system-prompt').value = SETTINGS.systemPrompt || '';
  document.getElementById('confirm-destructive').checked = SETTINGS.confirmDestructive !== false;
}

function updateConnectionBar() {
  const bar = document.getElementById('connection-bar');
  const text = bar.querySelector('.status-text');
  if (SETTINGS.apiKey && SETTINGS.endpoint) {
    const provider = LLM.isFreeLLMAPI(SETTINGS.endpoint) ? 'FreeLLMAPI'
      : LLM.detectProvider(SETTINGS.endpoint, SETTINGS.provider) === 'anthropic' ? 'Claude' : 'OpenAI-compatible';
    bar.className = 'status-bar connected';
    text.textContent = `${SETTINGS.model} · ${provider}`;
  } else {
    bar.className = 'status-bar disconnected';
    text.textContent = 'Sin configurar — abre ⚙️';
  }
}

async function testConnection() {
  const result = document.getElementById('test-result');
  const form = readSettingsForm();
  if (!form.endpoint || !form.model || !form.apiKey) {
    result.className = 'test-result error';
    result.textContent = '❌ Endpoint, modelo y API key son requeridos.';
    return;
  }
  result.className = 'test-result';
  result.textContent = '🔄 Probando conexión...';
  try {
    const { provider, reply, servedBy } = await LLM.ping(form);
    const name = LLM.isFreeLLMAPI(form.endpoint) ? 'FreeLLMAPI' : provider === 'anthropic' ? 'Anthropic' : 'OpenAI-compatible';
    result.className = 'test-result success';
    result.textContent = `✅ Conexión exitosa (${name}${servedBy ? ' → ' + servedBy : ''}). Respuesta: "${String(reply).slice(0, 80)}"`;
    loadModelList(form);
  } catch (err) {
    result.className = 'test-result error';
    result.textContent = `❌ ${err.message}`;
  }
}

/** Rellena la lista de sugerencias del campo Modelo con GET /models (si el servidor lo permite). */
async function loadModelList(form) {
  try {
    const models = await LLM.listModels(form || SETTINGS);
    const list = document.getElementById('model-list');
    if (!list || !models.length) return;
    list.innerHTML = '';
    const opts = LLM.isFreeLLMAPI((form || SETTINGS).endpoint) ? ['auto', 'auto:smartest', 'auto:fastest', ...models] : models;
    [...new Set(opts)].slice(0, 300).forEach((id) => {
      const o = document.createElement('option');
      o.value = id;
      list.appendChild(o);
    });
  } catch { /* opcional */ }
}

/* --- Init --- */
document.addEventListener('DOMContentLoaded', () => {
  loadSettings();
  applySettingsToUI();
  updateConnectionBar();

  const settingsModal = document.getElementById('settings-modal');

  function openSettings() {
    applySettingsToUI();
    settingsModal.classList.remove('hidden');
  }
  function closeSettings() {
    settingsModal.classList.add('hidden');
    document.getElementById('test-result').className = 'test-result hidden';
  }
  window.openSettings = openSettings;

  document.getElementById('btn-settings').addEventListener('click', openSettings);
  settingsModal.querySelector('.modal-close').addEventListener('click', closeSettings);
  settingsModal.querySelector('.modal-backdrop').addEventListener('click', closeSettings);

  document.getElementById('btn-toggle-key').addEventListener('click', () => {
    const input = document.getElementById('api-key');
    input.type = input.type === 'password' ? 'text' : 'password';
  });

  document.getElementById('btn-save-settings').addEventListener('click', () => {
    if (saveSettings()) {
      const result = document.getElementById('test-result');
      result.className = 'test-result success';
      result.textContent = '✅ Configuración guardada.';
      setTimeout(closeSettings, 800);
    }
  });

  document.getElementById('btn-test-connection').addEventListener('click', testConnection);

  document.querySelectorAll('.preset-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      const preset = PRESETS[btn.dataset.preset];
      if (!preset) return;
      document.getElementById('endpoint-url').value = preset.endpoint;
      document.getElementById('model-name').value = preset.model;
      document.getElementById('provider').value = preset.provider;
    });
  });
});
