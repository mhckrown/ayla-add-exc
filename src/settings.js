/* === Ayla Add-Exc — Settings Manager === */

const SETTINGS = {
  endpoint: 'https://generativelanguage.googleapis.com/v1beta/openai/',
  model: 'gemini-2.5-flash',
  apiKey: '',
  systemPrompt: `Eres un asistente experto en Excel. Tu función es ayudar al usuario a trabajar con sus hojas de cálculo.

REGLAS DE ORO:
1. Cuando el usuario pida analizar, calcular, modificar o consultar datos, SIEMPRE llama primero a las herramientas de lectura de Excel (get_selected_range, get_worksheet_data, get_worksheet_names) para obtener la información actual. No respondas basándote en suposiciones.
2. Después de leer los datos, analízalos y ejecuta las herramientas que necesites.
3. Explica SIEMPRE qué estás haciendo y el resultado obtenido.

FLUJO DE TRABAJO TÍPICO:
- Usuario: "analiza estos datos" → Tú: get_selected_range() → analizas → respondes
- Usuario: "calcula el promedio de ventas" → Tú: get_worksheet_data() → calculas → set_cell_value() → respondes
- Usuario: "crea un gráfico" → Tú: get_worksheet_data() → create_chart() → respondes
- Usuario: "dame información de la hoja" → Tú: get_worksheet_names() + get_active_worksheet() → respondes

NUNCA digas "no tengo acceso a tus datos" o "no puedo ver tu hoja". SIMPLEMENTE USA LAS HERRAMIENTAS para leer los datos.

IMPORTANTE:
- Siempre responde en español.
- Se claro, conciso y profesional.
- Si el usuario no especifica un rango, usa get_worksheet_data sin parámetros para leer la hoja activa completa.
- Después de escribir datos o fórmulas, explica qué pusiste y por qué.`,
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
