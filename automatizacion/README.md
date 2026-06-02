# Automatización ICP / Fondo / Competencia — estado

## RESUELTO y validado
- **ICP (icp_bcch.py)**: API del BCCh, serie `F022.TIB.TIP.D001.NO.Z.D` + composición
  diaria. Reproduce el ICP real con 0.0000% de error. Requiere `BCCH_USER`/`BCCH_PASS`
  (Secrets) y el web service ACTIVADO en la cuenta del BDE.
- **Retorno del fondo (build_dataset.py)**: ajuste por dividendo `(VC+d)/(VC_prev+d)-1`,
  con `d` recuperado automáticamente de la historia de cada template. 19 fondos CLP
  validados; Alto Aporte = M 0.62 / T 1.84 / S 3.66 / A 7.26 / Acum 5.82 (exacto).
- **Fórmulas**: M/T/S/A = nivel_fin/nivel_{1,3,6,12}-1; Acum = (fin/enero-1)/n*12.
- La historia se LEE del template (intocable); solo se calcula el mes nuevo.

## PENDIENTE
- **Competencia (competencia_cmf.py)**: CMF sin captcha (confirmado), pero falta
  confirmar los parámetros del formulario en vivo. Esqueleto listo.
- **3 fondos USD** (Reserva Dólar, Dólar, Dólar Caja): su retorno oficial NO sale del
  VC en dólares (0.4%) ni del VC×dólar observado. Metodología por aclarar.

## SEGURIDAD
- Rotar el token de GitHub (quedó expuesto). Poner el repo en privado.
- Credenciales BCCh solo como Secrets, nunca en el código.
