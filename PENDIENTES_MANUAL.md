# Pendientes manuales — Automatización folletos FIP

Resumen del estado tras la automatización. El **flujo mensual** es: subir 3 archivos en
`admin.html` (cartera, planilla VC, comentarios) y disparar el workflow. El sistema obtiene
ICP, Santander (CLP) y Banchile (USD) automáticamente.

## ✅ Ya funciona (validado en GitHub Actions, rama `claude/automatizacion-icp-fondo`)
- Fondos CLP: retornos con dividendo correcto (Alto Aporte clava la referencia 0,62 / 1,84 / 3,66 / 7,26 / 5,82).
- Fondos USD: retornos anualizados (Dólar Caja 5,07%); competencia Banchile scrapeada de CMF (1.467,7431).
- ICP y competencia Santander exactos.
- Scraper CMF arreglado (Santander serie UNIVE y Banchile serie A): toma el valor cuota correcto del popup.
- `admin.html` ya NO pide ICP ni Santander a mano.
- Fetch automático (ICP→BCCh, competencia→CMF) corre ANTES de actualizar templates y escribe los JSON.

## 🔧 LO QUE DEBES HACER TÚ (manual)

### 1. Configurar los secrets del Banco Central (CRÍTICO para ICP exacto)
En el repo: **Settings → Secrets and variables → Actions → New repository secret**. Crear dos:
- `BCCH_USER` = tu usuario de la API REST del BCCh (correo)
- `BCCH_PASS` = tu contraseña de la API REST del BCCh

Sin estos secrets el ICP del mes nuevo se calcula con la TPM de mindicador.cl (aproximado,
~0,01–0,02 pp/mes de error) en vez del valor exacto. CON los secrets, el ICP es exacto
(validado: reproduce 26.535,18 al centésimo).

### 2. Mergear la rama a `main`
Todo el trabajo está en la rama `claude/automatizacion-icp-fondo`. Revisa el diff y haz merge a
`main` cuando estés conforme. El workflow de producción corre sobre `main`.

### 3. Seguridad (URGENTE)
- **Rota el token de GitHub** `ghp_...` (quedó expuesto en el chat y en archivos del proyecto).
- **Pon los repos en privado** (`mivelascor/FIP`).
- Considera rotar también la contraseña del BCCh por la misma razón.

## ⚠️ Notas / límites
- El workflow commitea solo `folletos/`. Los JSON (`icp_clicp.json`, `comp_clp.json`,
  `comp_usd.json`) se actualizan en memoria durante la corrida; si quieres persistir el caché,
  habría que agregar `git add inputs/` al workflow (opcional).
- El parser de CMF toma el valor cuota como el número en rango [100, 20.000] tras la fecha
  (4ª columna del popup). Si CMF cambia el formato del popup, revisar `cmf_scraper.py`.
- Hay un guardrail: si la competencia scrapeada queda fuera de rango (Santander 3.000–20.000,
  Banchile 500–5.000) se ignora y el fondo USD/CLP de ese mes se omite (sale "—") en vez de
  mostrar un número erróneo.
