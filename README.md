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

### Paso 1: Preparar el servidor

El add-in necesita un servidor web que sirva los archivos. Opciones:

#### Opción A: Servidor local con Node.js (recomendado para desarrollo)

```powershell
cd ayla-add-exc
npm install
node server.js
```

Esto inicia:
- HTTP en `http://localhost:3000`
- HTTPS en `https://localhost:3001`

#### Opción B: Servidor IIS en Windows Server (para producción)

1. Abre **Administrador de IIS**
2. Crea un sitio web nuevo o aplicación
3. Copia la carpeta `ayla-add-exc` al directorio del sitio
4. Asegúrate de tener HTTPS configurado (certificado válido)
5. Actualiza `manifest.xml` con tu URL pública

#### Opción C: nginx en Linux

```nginx
server {
    listen 443 ssl;
    server_name ayla.midominio.com;
    root /ruta/a/ayla-add-exc;
    ssl_certificate /ruta/cert.pem;
    ssl_certificate_key /ruta/key.pem;
}
```

### Paso 2: Instalar en Excel

#### Método A: Instalador PowerShell (Windows)

```powershell
# Asegúrate de que el servidor esté corriendo
.\sideload.ps1 -ServerUrl "https://localhost:3001"

# O apuntando a un servidor de producción:
.\sideload.ps1 -ServerUrl "https://ayla.midominio.com"
```

#### Método B: Sideload manual

1. Abre Excel
2. Ve a **Insertar → Complementos → Mis complementos**
3. Haz clic en **Administrar mis complementos → Cargar complemento**
4. Navega hasta `manifest.xml` y selecciónalo
5. Aparecerá la pestaña **Ayla Add-Exc** en la cinta

#### Método C: Catálogo compartido (para múltiples equipos)

En un entorno corporativo:

1. **Opción fácil**: Copia el manifest modificado a cada equipo en:
   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
   ```
   (esto es lo que hace `sideload.ps1`)

2. **Opción centralizada**: Usa el Catálogo de SharePoint:
   - Carga el add-in en la biblioteca de Catálogo de aplicaciones de SharePoint
   - Los usuarios lo instalan desde **Insertar → Mis complementos → Compartidos conmigo**

3. **Opción GPO**: Distribuye el manifest via Política de Grupo:
   ```
   Computer Configuration\Administrative Templates\Microsoft Office 2016\Add-ins
   ```

### Paso 3: Configurar la conexión al modelo

1. En Excel, abre **Ayla Add-Exc → Abrir Chat**
2. Haz clic en **⚙️ Configuración**
3. Elige tu proveedor:

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
