# Fondos Financieros — Generador Automático de Folletos

Sistema de generación mensual automática de folletos comerciales para los fondos Vantrust Capital.

## Arquitectura

```
┌─────────────────────────────────────────────────────────────────┐
│  admin.html  (abrir en cualquier navegador, una vez al mes)     │
│                                                                  │
│  1. Subes cartera.xlsx     (composición del mes)                │
│  2. Escribes comentarios PM (CLP y USD)                         │
│  3. Haces clic en ▶ Generar folletos del mes                    │
└────────────────────┬────────────────────────────────────────────┘
                     │ GitHub API (PUT archivos + dispatch)
                     ▼
┌─────────────────────────────────────────────────────────────────┐
│  GitHub Actions Workflow                                         │
│                                                                  │
│  ① SQL Server (ODS.VALORES_CUOTA_GPI)  ← valores cuota fondos  │
│  ② mindicador.cl/api/tpm               ← ICP del BCCh          │
│  ③ CMF scraping (Playwright)           ← competencia           │
│  ④ inputs/cartera.xlsx                 ← composición           │
│  ⑤ inputs/comentarios.json             ← comentarios PM        │
│                                                                  │
│  → Calcula rentabilidades (mensuales, acumuladas, históricas)   │
│  → Genera 23 HTMLs (1 por fondo activo)                         │
│  → ZIP + commit a folletos/YYYY-MM/                             │
└─────────────────────────────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────────┐
│  folletos/2026-05/                                              │
│    FIP_Alto_Aporte_2026-05.html                                 │
│    FIP_Alto_Capital_2026-05.html                                │
│    ...  (23 folletos)                                           │
│  folletos/folletos_2026-05.zip                                  │
└─────────────────────────────────────────────────────────────────┘
```

## Setup inicial (una sola vez)

### 1. Clonar el repo
```bash
git clone https://github.com/mivelascor/fondos-financieros.git
cd fondos-financieros
```

### 2. Configurar el Secret SQL en GitHub
En **Settings → Secrets → Actions**, crear:
- `SQL_CONN_STR` → cadena de conexión ODBC al servidor ODS

Formato:
```
DRIVER={ODBC Driver 17 for SQL Server};SERVER=claudeods.vantrustcapital.cl;DATABASE=ODS;UID=tu_usuario;PWD=tu_contraseña;
```

### 3. Subir la estructura al repo
```
.github/workflows/generar_folletos.yml
scripts/
  main.py
  config.py
  requirements.txt
  etl/
    icp_bcch.py
    sql_extractor.py
    cmf_scraper.py
    cartera_reader.py
  calculos/
    rentabilidades.py
  generador/
    html_builder.py
admin.html
inputs/    (carpeta vacía — se llena via admin.html)
folletos/  (carpeta vacía — el workflow la llena)
```

---

## Uso mensual

1. Abrir `admin.html` en tu navegador
2. Arrastrar el `cartera.xlsx` del mes (exportado del sistema)
3. Pegar los comentarios del Portfolio Manager (CLP y USD)
4. Hacer clic en **▶ Generar folletos del mes**
5. Esperar ~5–10 minutos
6. Descargar el ZIP con los 23 folletos
7. Abrir cada HTML en Chrome → Ctrl+P → **Guardar como PDF** (A4, sin márgenes)

---

## Fondos activos (23)

| Fondo | SQL Name | Moneda |
|---|---|---|
| Alto Aporte | FIP VANTRUST LIQUIDEZ ALTO APORTE | CLP |
| Alto Capital | FIP VANTRUST LIQUIDEZ ALTO CAPITAL | CLP |
| Liquidez Caja | FIP VANTRUST LIQUIDEZ CAJA | CLP |
| Liquidez Uno | FIP VANTRUST LIQUIDEZ I | CLP |
| Liquidez Sencillo | FIP VANTRUST LIQUIDEZ SENCILLO | CLP |
| Liquidez Rendimiento | FIP VANTRUST LIQUIDEZ RENDIMIENTO | CLP |
| Liquidez Local | FIP VANTRUST LIQUIDEZ LOCAL | CLP |
| Liquidez Plus | FIP VANTRUST LIQUIDEZ PLUS | CLP |
| Liquidez Permanente | FIP VANTRUST LIQUIDEZ PERMANENTE | CLP |
| Liquidez Continua | FIP VANTRUST LIQUIDEZ CONTINUA | CLP |
| Liquidez Corto Plazo | FIP VANTRUST LIQUIDEZ CORTO PLAZO | CLP |
| Liquidez Monetario I | FIP VANTRUST LIQUIDEZ MONETARIO I | CLP |
| Liquidez Disponible I | FIP VANTRUST LIQUIDEZ DISPONIBLE I | CLP |
| Liquidez Presente | FIP VANTRUST LIQUIDEZ PRESENTE | CLP |
| Liquidez Corriente | FIP VANTRUST LIQUIDEZ CORRIENTE | CLP |
| Liquidez Alto Monto | FIP VANTRUST LIQUIDEZ ALTO MONTO | CLP |
| Liquidez Activa | FIP VANTRUST LIQUIDEZ ACTIVA | CLP |
| Liquidez Efectivo | FIP VANTRUST LIQUIDEZ EFECTIVO | CLP |
| Liquidez Flexible | FIP VANTRUST LIQUIDEZ FLEXIBLE | CLP |
| Liquidez Dólar | FIP VANTRUST LIQUIDEZ DOLAR | USD |
| Liquidez Dólar Caja | FIP VANTRUST LIQUIDEZ DOLAR CAJA | USD |
| Liquidez Reserva Dólar | FIP VANTRUST LIQUIDEZ RESERVA DOLAR | USD |
| Factura Dólar | FIP VANTRUST LIQUIDEZ FACTURA DOLAR | USD |

---

## Competencia (CMF scraping)

- **CLP**: Santander Money Market, serie UNIVE (RUT 8057)
- **USD**: BanChile Corporate Dollar, serie A (RUT 8248)

Si el scraping falla (sitio CMF caído), el sistema usa el histórico hardcodeado en `cmf_scraper.py` y extrapola el último mes disponible. Al terminar exitosamente, el workflow actualiza automáticamente el histórico en el archivo y hace commit.

---

## Generar PDF desde HTML

En Chrome: Ctrl+P → Impresora: "Guardar como PDF" → Tamaño: A4 → Márgenes: Ninguno → Habilitar gráficos de fondo → Guardar.

O con Playwright (automático):
```bash
playwright pdf folletos/2026-05/FIP_Alto_Aporte_2026-05.html output.pdf
```
