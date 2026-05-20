# FIP — Fondos Financieros Vantrust

Sistema de generación mensual automática de folletos comerciales.

## 🌐 Acceso
Abre `index.html` en tu navegador → contraseña: `1234`

## 📁 Estructura
```
FIP/
├── index.html                    ← Login
├── menu_principal.html           ← Menú (Fondos / Admin)
├── MenuFF.html                   ← Lista de fondos + descargas
├── admin.html                    ← Generar folletos del mes
├── O1FipVantrustLiquidezActiva.html   ← Página de cada fondo
├── O2...O24...html
│
├── .github/workflows/
│   └── generar_folletos.yml      ← Workflow mensual automático
│
├── scripts/
│   ├── main.py                   ← Orquestador
│   ├── config.py                 ← Configuración fondos
│   ├── requirements.txt
│   ├── etl/
│   │   ├── actualizar_templates.py   ← ICP + SQL + CMF → Excel
│   │   ├── template_reader.py        ← Lee Excel calculado
│   │   ├── excel_reader.py           ← Lee cartera.xlsx
│   │   ├── icp_bcch.py               ← ICP desde BCCh/mindicador
│   │   ├── sql_extractor.py          ← Valores cuota ODS SQL
│   │   └── cmf_scraper.py            ← Competencia CMF
│   ├── calculos/
│   │   └── rentabilidades.py
│   ├── generador/
│   │   ├── pptx_builder.py
│   │   └── pdf_exporter.py
│   └── recalc_skill/             ← LibreOffice recalculator
│       └── recalc.py
│
├── inputs/
│   ├── cartera.xlsx              ← Subido desde admin.html
│   ├── comentarios.json          ← Subido desde admin.html
│   └── templates/                ← 24 templates Excel (fuente de verdad)
│       ├── TEMPLATE_FONDO_LIQUIDEZ_ACTIVA.xlsx
│       └── ...
│
├── folletos/
│   ├── latest.zip                ← Último ZIP generado
│   ├── 2026-04/                  ← Folletos Abril 2026
│   │   ├── FIP_Liquidez_Uno_2026-04.html
│   │   └── ...
│   └── folletos_2026-04.zip
│
└── regl.int/                     ← Reglamentos internos (PDFs)
    └── *.pdf
```

## 🚀 Uso mensual
1. Ir a `admin.html`
2. Subir `cartera.xlsx` del mes
3. Pegar comentarios PM (CLP y USD)
4. Ingresar tu GitHub Token
5. Click **Generar folletos** → esperar ~10 min
6. Descargar ZIP

## ⚙️ Setup inicial (una vez)
- Agregar secret `SQL_CONN_STR` en GitHub → Settings → Secrets → Actions
- Formato: `DRIVER={ODBC Driver 17 for SQL Server};SERVER=claudeods.vantrustcapital.cl;DATABASE=ODS;UID=xxx;PWD=xxx;`

## 📊 Fuentes de datos (automáticas)
| Dato | Fuente |
|------|--------|
| ICP | mindicador.cl / BCCh API |
| Valor cuota fondos | ODS SQL Server (`ODS.VALORES_CUOTA_GPI`) |
| Competencia CLP | CMF — Santander Money Market (rut 8057, serie UNIVE) |
| Competencia USD | CMF — BanChile Corporate Dollar (rut 8248, serie A) |
| Cartera/Composición | `cartera.xlsx` subido manualmente |
| Comentario PM | Ingresado en `admin.html` |
