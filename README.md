# ✦ Ayla Add-Exc

**Asistente IA para Excel** — chat inteligente que automatiza tareas en tus hojas de cálculo.

> Complemento que se integra directamente en Excel (on-premises y Microsoft 365).  
> Conéctalo a cualquier modelo de IA vía API compatible con OpenAI (Gemini, ChatGPT, Claude, etc.) y controla Excel desde lenguaje natural.

---

## 📦 Contenido del paquete

```
ayla-add-exc/
├── manifest.xml              # Identidad del add-in para Excel
├── server.js                 # Servidor HTTPS/HTTP para servir el add-in
├── package.json              # Dependencias Node.js
├── sideload.ps1              # Instalador automático para Windows
├── generate-icons.py         # Generador de iconos
├── .cert/                    # Certificado SSL (se genera solo)
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-80.png
│   └── icon-128.png
└── src/
    ├── taskpane.html         # Interfaz del chat
    ├── taskpane.css          # Estilos oscuros modernos
    ├── taskpane.js           # Lógica del chat y conexión API
    ├── excel-bridge.js       # 25 herramientas de Excel para el modelo IA
    ├── settings.js           # Configuración de conexión
    └── vendor/
        ├── marked.min.js     # Renderizado Markdown (local)
        ├── highlight.min.js  # Resaltado de código (local)
        └── highlight.min.css # Estilos de código (local)
```

---

## ⚙️ ¿Qué puede hacer?

El modelo de IA recibe un conjunto de **25 herramientas de Excel** y las ejecuta automáticamente según lo que le pidas:

| Categoría | Acciones disponibles |
|---|---|
| **Lectura** | Leer selección, hoja activa, todas las hojas, info del libro |
| **Escritura** | Escribir valores, fórmulas, arrays 2D, autocompletar |
| **Formato** | Color, negrita, tamaño, alineación, bordes, formato numérico, columnas, filas |
| **Formato condicional** | Escalas de color, barras de datos, reglas personalizadas |
| **Tablas** | Crear tablas con estilo |
| **Gráficos** | Barras, columnas, líneas, pastel, dona, área, dispersión |
| **Ordenar/filtrar** | Auto-filtro, ordenar por columna |
| **Hojas** | Insertar, eliminar, renombrar, activar, proteger |
| **Estructura** | Insertar/eliminar filas/columnas, combinar celdas |
| **Navegación** | Ir a celda, activar hoja |
| **Utilidades** | Nombres definidos, autoajuste, merge |

Ejemplos de lo que puedes decir:
- *"Calcula el promedio de este rango y ponlo en C10"*
- *"Crea una tabla con formato en A1:D50"*
- *"Haz un gráfico de barras de ventas por mes"*
- *"Aplica escala de colores verde-rojo a D2:D100"*
- *"Ordena estos datos por columna B descendente"*
- *"Filtra filas donde el total > 1000"*
- *"Inserta una columna de totales con fórmula SUM"*
- *"Protege la hoja con contraseña admin123"*

---

## 🚀 Instalación

### Requisitos mínimos

| Componente | Requisito |
|---|---|
| **Excel** | Office 2019, Office 2021, Office LTSC, o Microsoft 365 |
| **Windows** | Windows 10 (build 1903+) o Windows 11 |
| **WebView2 Runtime** | Incluido en Win11 / [Descargar para Win10](https://developer.microsoft.com/microsoft-edge/webview2/) |
| **Node.js** | 18+ (solo para el servidor) |
| **Conexión a internet** | Solo para llamar a la API del modelo |

### Paso 1: Obtener el manifest

El panel ya está publicado en GitHub Pages (`https://mhckrown.github.io/ayla-add-exc/`, rama `main`), así que no necesitas servidor propio. Solo descarga [`manifest.xml`](https://raw.githubusercontent.com/mhckrown/ayla-add-exc/main/manifest.xml) (clic derecho → *Guardar como*).

### Paso 2: Instalar en Excel

#### Windows (Excel de escritorio) — carpeta compartida de confianza

1. Crea una carpeta, p. ej. `C:\AylaAddin`, y copia ahí `manifest.xml`.
2. Clic derecho en la carpeta → **Propiedades → Compartir → Compartir…** → **Compartir**. Anota la ruta de red, p. ej. `\\MI-PC\AylaAddin`.
3. En Excel: **Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza → Catálogos de complementos de confianza**.
4. En **Dirección URL del catálogo** pega `\\MI-PC\AylaAddin`, pulsa **Agregar catálogo**, marca **Mostrar en menú** y acepta.
5. Cierra Excel por completo y vuelve a abrirlo.
6. **Inicio → Complementos → Más complementos → CARPETA COMPARTIDA** (o **Insertar → Mis complementos**), elige **Ayla** y pulsa **Agregar**.

> También puedes ejecutar `setup.bat`, que copia el manifest en `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`.

#### Mac

1. Copia `manifest.xml` en `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (crea la carpeta `wef` si no existe).
2. Reinicia Excel → **Insertar → Complementos → Mis complementos** → **Ayla**.

#### Excel en la web

**Inicio → Complementos → Más complementos → Cargar mi complemento** → selecciona `manifest.xml`.

#### Actualizar después de un cambio

Office guarda en caché el panel. Si no ves la versión nueva, cierra Excel y borra el contenido de `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (Windows) o `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/` (Mac).

#### Desarrollo sin Excel

`npm run dev` abre un simulador de Excel en `http://localhost:3000/dev/` con el panel real y un Office.js simulado.

### Paso 3: Configurar la conexión al modelo

1. En Excel, abre **Ayla Add-Exc → Abrir Chat**
2. Haz clic en **⚙️ Configuración**
3. Elige tu proveedor:

#### 🆓 FreeLLMAPI (local, en tu equipo — recomendado)

[FreeLLMAPI](https://github.com/tashfeenahmed/freellmapi) agrupa muchos proveedores gratuitos detrás de un único endpoint compatible con OpenAI (con streaming y tool calling).

1. Inicia FreeLLMAPI en tu equipo (por defecto `http://localhost:3001`; la app de escritorio usa `31415`) y añade tus claves de proveedores en su dashboard.
2. Copia la **clave unificada** desde la página *Keys* del dashboard.
3. Como el panel se sirve desde `https://mhckrown.github.io`, el navegador de Excel necesita permiso CORS. Elige **una** opción:
   - **Windows, un clic:** descarga [`iniciar-ayla-freellmapi.bat`](https://raw.githubusercontent.com/mhckrown/ayla-add-exc/main/iniciar-ayla-freellmapi.bat) (clic derecho → *Guardar como*) y ábrelo. Instala Node.js si hace falta, descarga el puente y lo deja corriendo. Endpoint: `http://127.0.0.1:3939/v1`.
   - **Puente local (sin tocar FreeLLMAPI):** en este repo ejecuta `npm run freellmapi` y usa el endpoint `http://localhost:3939/v1`.
   - **Directo:** añade `DASHBOARD_ORIGINS=https://mhckrown.github.io` al `.env` de FreeLLMAPI, reinícialo y usa `http://localhost:3001/v1`.
4. En Ayla pulsa el preset **🆓 FreeLLMAPI (local)**, pega la clave unificada y deja el modelo en `auto` (o `auto:smartest` / `auto:fastest`).

| Campo | Valor |
|---|---|
| Endpoint | `http://localhost:3939/v1` (puente) o `http://localhost:3001/v1` (directo) |
| Modelo | `auto` |
| API Key | Clave unificada de FreeLLMAPI |

Variables del puente: `FREELLMAPI_URL` (destino, por defecto detecta `127.0.0.1:3001` y `:31415`), `PROXY_PORT` (por defecto `3939`), `ALLOWED_ORIGINS` (orígenes permitidos, separados por coma).

#### ✨ Gemini Flash (gratis — recomendado)

| Campo | Valor |
|---|---|
| Endpoint | `https://generativelanguage.googleapis.com/v1beta/openai/` |
| Modelo | `gemini-2.5-flash` |
| API Key | Obtén una gratis en https://aistudio.google.com/apikey |

*(La cuenta gratuita de Gemini Flash permite 1,500 requests/día)*

#### 🔵 OpenAI

| Campo | Valor |
|---|---|
| Endpoint | `https://api.openai.com/v1` |
| Modelo | `gpt-4o-mini` o `gpt-4o` |
| API Key | De https://platform.openai.com/api-keys |

#### 🟠 OpenRouter

| Campo | Valor |
|---|---|
| Endpoint | `https://openrouter.ai/api/v1` |
| Modelo | `google/gemini-2.5-flash`, `anthropic/claude-sonnet-4`, `openai/gpt-4o` |
| API Key | De https://openrouter.ai/keys |

4. Haz clic en **Probar conexión** para verificar
5. Haz clic en **Guardar**

---

## 🔒 Consideraciones de seguridad

- **API Keys**: Se almacenan en `localStorage` de Office. No se envían a ningún lado excepto al endpoint que configures.
- **HTTPS**: Office requiere HTTPS para add-ins en producción. El servidor incluido genera un certificado autofirmado.
- **Certificado autofirmado**: En tu primera visita a `https://localhost:3001`, el navegador mostrará una advertencia. Haz clic en **Avanzado → Continuar**.
- **Datos de Excel**: Solo se envían al modelo de IA los datos de las celdas que tú adjuntes explícitamente vía el botón 📎.

---

## 🌐 Multiplataforma

| Plataforma | Sideload | Notas |
|---|---|---|
| **Excel en Windows (escritorio)** | ✅ | Vía sideload.ps1 o carpeta WEF |
| **Excel en Mac** | ✅ | Vía sideload con manifest local |
| **Excel para Web** | ✅ | Vía SharePoint Catalog o AppSource |

Para Mac:
```bash
# Copiar manifest a la carpeta de contenedores
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

---

## 🛠️ Solución de problemas

| Problema | Causa | Solución |
|---|---|---|
| "No se puede cargar el complemento" | HTTPS incorrecto | Usa localhost:3000 (HTTP) para pruebas |
| Pantalla en blanco | WebView2 no instalado | Descarga desde Microsoft |
| Error 401 en conexión | API Key inválida | Verifica en settings ⚙️ |
| No responde | Modelo no soporta tool calling | Usa Gemini Flash o GPT-4o-mini |
| "Failed to fetch" | Firewall / proxy | Verifica conectividad al endpoint |
| Sideload no aparece | Excel no reiniciado | Cierra y abre Excel |

---

## 📄 Licencia

Uso interno. Desarrollado por Ayla.
