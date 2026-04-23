# Flujo de Caja — Modelo Automático

Aplicación web para generar el modelo de Flujo de Caja en el marco de postulación de proyectos de investigación aplicada a fuentes de financiamiento público.

## ¿Qué hace?

El usuario completa un formulario con:
- Inversiones (maquinaria, equipos, infraestructura)
- Costos fijos mensuales
- Costos de producción por unidad
- Cantidades de venta proyectadas por año (1–5)
- Canvas de Modelo de Negocios (opcional)

Y descarga un archivo `.xlsx` con las 6 pestañas del modelo original, con formato y fórmulas idénticas al documento oficial:
1. Inversión y Depreciación
2. Costos
3. Flujo de Caja (VAN, TIR, Payback, C/B)
4. Payback (cálculo automático)
5. Costo/Beneficio
6. Diagrama Modelo de Negocios

---

## Estructura del proyecto

```
flujo-caja-app/
├── frontend/
│   └── index.html          # App completa (un solo archivo HTML)
├── backend/
│   ├── main.py             # API FastAPI
│   ├── excel_generator.py  # Generador de Excel con openpyxl
│   └── requirements.txt
└── README.md
```

---

## Deploy en 3 pasos

### 1. Backend — Railway (gratis)

1. Ve a [railway.app](https://railway.app) y crea una cuenta con GitHub
2. Nuevo proyecto → Deploy from GitHub repo → selecciona este repo
3. Configura el directorio raíz: `backend`
4. Railway detecta Python automáticamente. Agrega este comando de inicio:
   ```
   uvicorn main:app --host 0.0.0.0 --port $PORT
   ```
5. Copia la URL que Railway te asigna (ej: `https://tu-app.up.railway.app`)

**Alternativa: Render.com**
1. New Web Service → conecta tu repo
2. Root directory: `backend`
3. Build command: `pip install -r requirements.txt`
4. Start command: `uvicorn main:app --host 0.0.0.0 --port $PORT`

### 2. Frontend — conectar backend

Edita `frontend/index.html`, busca esta línea (cerca del final del `<script>`):

```js
const API_URL = window.BACKEND_URL || 'http://localhost:8000';
```

Reemplázala con tu URL de Railway/Render:

```js
const API_URL = 'https://tu-app.up.railway.app';
```

### 3. Frontend — GitHub Pages

1. Ve a Settings → Pages en tu repositorio
2. Source: `Deploy from a branch`
3. Branch: `main`, carpeta: `/frontend` (o `/docs` si renombras la carpeta)
4. Guarda — en ~2 minutos tu app estará en `https://tu-usuario.github.io/nombre-repo/`

---

## Uso local (desarrollo)

### Backend
```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload
# API disponible en http://localhost:8000
# Docs en http://localhost:8000/docs
```

### Frontend
```bash
# Simplemente abre el archivo en tu navegador
open frontend/index.html

# O usa un servidor local
cd frontend
python3 -m http.server 3000
# Visita http://localhost:3000
```

---

## Acceso privado (solo tu equipo)

Para que solo tu equipo pueda usar la app, tienes dos opciones sin costo:

**Opción A — GitHub Pages privado** (requiere plan GitHub Team ~$4/usuario/mes)
- Settings → Pages → solo visibles para miembros del repositorio

**Opción B — Contraseña simple en el frontend** (gratis)
Agrega esto al inicio del `<body>` en `index.html`:
```html
<script>
  const PASS = 'tu-clave-aqui';
  if (localStorage.getItem('auth') !== PASS) {
    const input = prompt('Ingresa la clave de acceso:');
    if (input !== PASS) { document.body.innerHTML = '<p>Acceso denegado</p>'; }
    else { localStorage.setItem('auth', PASS); }
  }
</script>
```

**Opción C — Netlify con autenticación de identidad** (gratis hasta 5 usuarios)
Netlify Identity permite proteger el sitio con email/contraseña sin costo adicional.

---

## Personalización

### Cambiar valores por defecto
En `frontend/index.html`, busca el bloque `let inversiones = [...]` y modifica los ejemplos que aparecen al cargar la app.

### Agregar campos al formulario
1. Agrega el campo HTML en la sección correspondiente
2. Inclúyelo en el objeto `payload` dentro de `handleDownload()`
3. En `backend/main.py`, agrégalo al modelo Pydantic correspondiente
4. En `backend/excel_generator.py`, úsalo en la función `build_sheet*` que corresponda

### Cambiar la tasa de impuesto (actualmente 27%)
En `excel_generator.py`, busca `0.27` y reemplaza por el valor que corresponda.

---

## Tecnologías

| Capa | Tecnología | Por qué |
|------|-----------|---------|
| Frontend | HTML + CSS + JS vanilla | Sin dependencias, carga instantánea, fácil de mantener |
| Backend | Python + FastAPI | Rápido, tipado, documentación automática en /docs |
| Excel | openpyxl | Única librería que replica colores, merges y fórmulas con fidelidad |
| Deploy | Railway + GitHub Pages | Ambos gratuitos para uso de equipo pequeño |

---

## Notas

- El Excel generado contiene **fórmulas reales** (no valores hardcodeados), igual que el original
- Los colores institucionales del documento (azul marino `#1F3864` y verde oscuro `#385623`) están replicados exactamente
- Las pestañas 4 (Payback) y 5 (C/B) se generan automáticamente con sus fórmulas enlazadas
- Compatible con Excel 2016+, LibreOffice Calc y Google Sheets
