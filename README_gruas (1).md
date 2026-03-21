# Dashboard Grúas Leasing — AEROSAN 🏗️

Dashboard interactivo con datos de grúas en leasing. Se actualiza automáticamente todos los días desde tu Excel en OneDrive.

## ¿Cómo configurarlo? (una sola vez)

### Paso 1 — Subir los archivos

Sube estos 4 archivos a tu repositorio GitHub:
- `generar_dashboard.py`
- `template.html`
- `.github/workflows/actualizar.yml`
- `README.md`

### Paso 2 — Preparar tu Excel en OneDrive

Tu Excel debe tener una hoja con estas columnas (los nombres son flexibles):

| GRUA | PATENTE | KM_MES |
|------|---------|--------|
| 1    | AB-1234 | 42     |
| 2    | CD-5678 | 98     |
| ...  | ...     | ...    |

Luego:
1. Abre tu Excel en OneDrive
2. Clic en **Compartir → Copiar vínculo**
3. Asegúrate que diga **"Cualquier persona con el vínculo puede ver"**

### Paso 3 — Guardar la URL como Secret en GitHub

1. En tu repositorio ve a **Settings → Secrets and variables → Actions**
2. Clic en **New repository secret**
3. **Name:** `ONEDRIVE_URL`
4. **Secret:** pega la URL de tu Excel
5. Clic en **Add secret**

### Paso 4 — Activar GitHub Pages

1. Ve a **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: **main → / (root)**
4. Clic en **Save**

### Paso 5 — Correr por primera vez

1. Ve a la pestaña **Actions**
2. Clic en **Actualizar Dashboard Grúas**
3. Clic en **Run workflow → Run workflow**
4. Espera 2 minutos

### Paso 6 — Ver tu dashboard

```
https://yversondelus-spec.github.io/dashboard-gruas/
```

## Actualización automática

El dashboard se actualiza solo todos los días a las **08:00 AM hora Chile**.  
También puedes actualizarlo manualmente desde **Actions → Run workflow**.

## Estructura de archivos

```
dashboard-gruas/
├── .github/
│   └── workflows/
│       └── actualizar.yml    ← automatización
├── generar_dashboard.py      ← script principal
├── template.html             ← plantilla del dashboard
├── index.html                ← dashboard generado (no editar)
└── README.md
```

## Columnas del Excel reconocidas

El script detecta automáticamente columnas que contengan estas palabras:
- **Grúa:** GRUA, GRÚA, ID, NUMERO, NÚMERO
- **Patente:** PATENTE, PLACA, PLATE
- **Kilómetros:** KM, KILOMETRO, KILÓMETRO, RECORRIDO
