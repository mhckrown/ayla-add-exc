/* === Ayla Add-Exc — Settings Manager === */

const SETTINGS = {
  endpoint: 'https://generativelanguage.googleapis.com/v1beta/openai/',
  model: 'gemini-2.5-flash',
  apiKey: '',
  systemPrompt: `Eres un asistente experto en Excel. Tu función es ayudar al usuario a trabajar con sus hojas de cálculo.

Cuando el usuario pida algo, SIGUE ESTOS PASOS:
1. Analiza el contexto (datos de celdas, estructura de la hoja).
2. Usa las herramientas de Excel disponibles para ejecutar lo que pide.
3. Explica siempre lo que hiciste y el resultado.

Herramientas disponibles:
- Leer datos de celdas, rangos, hojas completas
- Escribir valores, fórmulas, formatos
- Crear y modificar gráficos, tablas, filtros
- Aplicar formato condicional
- Ordenar y filtrar datos
- Insertar/eliminar filas, columnas, hojas
- Calcular promedios, sumas, conteos y estadísticas

IMPORTANTE:
- Cuando uses una herramienta, EXPLICA al usuario qué estás haciendo.
- Si necesitas leer datos primero para responder, usa get_selected_range o get_worksheet_data.
- Siempre devuelve la respuesta en español.
- Se claro y conciso.`,
};

/* --- persistence --- */
function saveSettings() {
  try {
    const data = {
      endpoint: document.getElementById('endpoint-url').value.trim(),
      model: document.getElementById('model-name').value.trim(),
      apiKey: document.getElementById('api-key').value.trim(),
      systemPrompt: document.getElementById('system-prompt').value.trim(),
    };
    localStorage.setItem('ayla_settings', JSON.stringify(data));
    // Office roaming settings (portable entre equipos si usan la misma cuenta)
    if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
      Office.context.roamingSettings.set('ayla_settings', JSON.stringify(data));
      Office.context.roamingSettings.saveAsync();
    }
    Object.assign(SETTINGS, data);
    updateConnectionBar();
    return true;
  } catch (e) {
    console.error('Error saving settings:', e);
    return false;
  }
}

function loadSettings() {
  try {
    // Try localStorage first, then roaming settings
    let raw = localStorage.getItem('ayla_settings');
    if (!raw && typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
      raw = Office.context.roamingSettings.get('ayla_settings');
    }
    if (raw) {
      const data = JSON.parse(raw);
      Object.assign(SETTINGS, data);
    }
  } catch (e) {
    console.warn('Could not load settings:', e);
  }
}

function applySettingsToUI() {
  document.getElementById('endpoint-url').value = SETTINGS.endpoint || '';
  document.getElementById('model-name').value = SETTINGS.model || '';
  document.getElementById('api-key').value = SETTINGS.apiKey || '';
  document.getElementById('system-prompt').value = SETTINGS.systemPrompt || '';
}

function getEndpointUrl() {
  let url = SETTINGS.endpoint || '';
  url = url.replace(/\/+$/, ''); // trailing slash
  // Ensure it ends with /chat/completions
  if (!url.endsWith('/chat/completions')) {
    url += '/chat/completions';
  }
  return url;
}

function updateConnectionBar() {
  const bar = document.getElementById('connection-bar');
  const dot = bar.querySelector('.status-dot');
  const text = bar.querySelector('.status-text');
  if (SETTINGS.apiKey && SETTINGS.endpoint) {
    bar.className = 'status-bar connected';
    dot.style.background = 'var(--success)';
    text.textContent = `Conectado → ${SETTINGS.model}`;
  } else {
    bar.className = 'status-bar disconnected';
    dot.style.background = 'var(--danger)';
    text.textContent = 'Sin configurar — abre ⚙️';
  }
}

async function testConnection() {
  const result = document.getElementById('test-result');
  result.className = 'test-result hidden';
  result.textContent = '';

  const endpoint = document.getElementById('endpoint-url').value.trim();
  const model = document.getElementById('model-name').value.trim();
  const key = document.getElementById('api-key').value.trim();

  if (!endpoint || !model) {
    result.className = 'test-result error';
    result.textContent = '❌ Endpoint y modelo son requeridos.';
    return;
  }
  if (!key) {
    result.className = 'test-result error';
    result.textContent = '❌ API Key requerida.';
    return;
  }

  result.className = 'test-result';
  result.textContent = '🔄 Probando conexión...';

  try {
    let url = endpoint.replace(/\/+$/, '');
    url += '/chat/completions';

    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${key}`,
      },
      body: JSON.stringify({
        model: model,
        messages: [{ role: 'user', content: 'Responde solo "ok" si me entiendes.' }],
        max_tokens: 10,
      }),
    });

    if (!resp.ok) {
      const errText = await resp.text().catch(() => '');
      result.className = 'test-result error';
      result.textContent = `❌ Error ${resp.status}: ${resp.statusText}${errText ? ' — ' + errText.slice(0, 200) : ''}`;
      return;
    }

    const data = await resp.json();
    const reply = data.choices?.[0]?.message?.content || '(vacio)';
    result.className = 'test-result success';
    result.textContent = `✅ Conexión exitosa! Modelo respondió: "${reply.slice(0, 100)}"`;
  } catch (err) {
    result.className = 'test-result error';
    result.textContent = `❌ Error de conexión: ${err.message}`;
  }
}

/* --- Init --- */
document.addEventListener('DOMContentLoaded', () => {
  loadSettings();
  applySettingsToUI();
  updateConnectionBar();

  // Settings modal
  const settingsBtn = document.getElementById('btn-settings');
  const settingsModal = document.getElementById('settings-modal');
  const modalClose = settingsModal.querySelector('.modal-close');
  const modalBackdrop = settingsModal.querySelector('.modal-backdrop');

  function openSettings() {
    applySettingsToUI();
    settingsModal.classList.remove('hidden');
  }

  function closeSettings() {
    settingsModal.classList.add('hidden');
    document.getElementById('test-result').className = 'test-result hidden';
  }

  settingsBtn.addEventListener('click', openSettings);
  modalClose.addEventListener('click', closeSettings);
  modalBackdrop.addEventListener('click', closeSettings);

  // Config link in welcome
  const configLink = document.getElementById('config-link');
  if (configLink) configLink.addEventListener('click', (e) => { e.preventDefault(); openSettings(); });

  // Toggle API key visibility
  document.getElementById('btn-toggle-key').addEventListener('click', () => {
    const input = document.getElementById('api-key');
    input.type = input.type === 'password' ? 'text' : 'password';
  });

  // Save
  document.getElementById('btn-save-settings').addEventListener('click', () => {
    if (saveSettings()) {
      const result = document.getElementById('test-result');
      result.className = 'test-result success';
      result.textContent = '✅ Configuración guardada.';
      setTimeout(closeSettings, 1000);
    }
  });

  // Test
  document.getElementById('btn-test-connection').addEventListener('click', testConnection);

  // Preset providers
  document.querySelectorAll('.preset-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const ep = btn.dataset.endpoint;
      const model = btn.dataset.model;
      if (ep === 'custom') return;
      document.getElementById('endpoint-url').value = ep;
      document.getElementById('model-name').value = model;
    });
  });
});
