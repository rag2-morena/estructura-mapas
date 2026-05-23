# MORENA · Tenant Estructura · Mapas de Tabletas

Actualización automática cada 2 horas desde Samsung Knox Manage API.

## URLs de acceso

| Entidad | URL |
|---|---|
| Ciudad de México | `/docs/09_CDMX.html` |
| Estado de México | `/docs/15_Estado_de_Mexico.html` |
| Jalisco | `/docs/14_Jalisco.html` |
| *(ver index completo)* | `/docs/index.html` |

## Configuración inicial (una sola vez)

### 1. Agregar secretos en GitHub
Ve a `Settings → Secrets → Actions` y agrega:
- `KNOX_BASE_URL` = `https://us02.manage.samsungknox.com`
- `KNOX_CLIENT_ID` = `knox-estructura@morena.si`
- `KNOX_SECRET` = tu client secret de Knox

### 2. Activar GitHub Pages
Ve a `Settings → Pages`:
- Source: `Deploy from a branch`
- Branch: `main` / carpeta `docs`

### 3. Copiar logo
Coloca el archivo `logo.b64` en la carpeta `scripts/`

### 4. Primera ejecución manual
Ve a `Actions → Generar Mapas Estructura → Run workflow`

## Estructura del repositorio
```
estructura-mapas/
├── .github/
│   └── workflows/
│       └── generar_mapas.yml   ← cron cada 2 horas
├── scripts/
│   ├── generar_mapas.py        ← script principal
│   └── logo.b64                ← logo MORENA (agregar manualmente)
├── docs/                       ← HTML generados (GitHub Pages)
│   ├── index.html
│   ├── 09_CDMX.html
│   └── ...
└── README.md
```
