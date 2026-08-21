# -*- coding: utf-8 -*-
# id, nombre_menu, titulo_l1, titulo_l2, rut, moneda, estado, reglamento, prorroga, folleto_pdf
#
# CRITERIO DE VIGENCIA (acordado con la usuaria, ago-2026):
#   vencido = su reglamento vino en Vencidos.zip  O  su valor cuota no aparece en
#             inputs/planilla_vc.xlsx desde mayo de 2026.
#   En la duda se deja "vigente".
#   Un fondo vencido IGUAL genera folleto si tiene valor cuota en la planilla:
#   el estado solo controla el filtro del menu (MenuFF.html).
R = [
("O1FipVantrustLiquidezActiva","FIP Liquidez Activa","Liquidez","ACTIVA","76.637.334-8","CLP","vigente","76637334-8 FIP Vantrust Liquidez Activa.pdf",None,"FIP_VANTRUST_LIQUIDEZ_ACTIVA"),
("O2FipVantrustLiquidezAltoAporte","FIP Liquidez Alto Aporte","Liquidez","ALTO APORTE","77.270.966-8","CLP","vigente",None,"Prorroga FIP Vantrust Liquidez Alto Aporte (2026-01).pdf","FIP_VANTRUST_LIQUIDEZ_ALTO_APORTE"),
("O3FipVantrustLiquidezAltoCapital","FIP Liquidez Alto Capital","Liquidez","ALTO CAPITAL","76.933.995-7","CLP","vigente","76933995-7 FIP Vantrust Liquidez Alto Capital.pdf","Prorroga FIP Vantrust Liquidez Alto Capital (2026-01).pdf","FIP_VANTRUST_LIQUIDEZ_ALTO_CAPITAL"),
("O4FipVantrustLiquidezAltoMonto","FIP Liquidez Alto Monto","Liquidez","ALTO MONTO","77.414.857-4","CLP","vigente","77414857-4 FIP Vantrust Liquidez Alto Monto.pdf",None,"FIP_VANTRUST_LIQUIDEZ_ALTO_MONTO"),
("O5FipVantrustLiquidezCaja","FIP Liquidez Caja","Liquidez","CAJA","76.933.989-2","CLP","vigente",None,"Prorroga FIP Vantrust Liquidez Caja (2026-01).pdf","FIP_VANTRUST_LIQUIDEZ_CAJA"),
("O6FipVantrustLiquidezContinua","FIP Liquidez Continua","Liquidez","CONTINUA","77.806.944-K","CLP","vigente",None,None,"FIP_VANTRUST_LIQUIDEZ_CONTINUA"),
("O7FipVantrustLiquidezCorriente","FIP Liquidez Corriente","Liquidez","CORRIENTE","77.428.236-K","CLP","vigente","77428236-K FIP Vantrust Liquidez Corriente.pdf",None,"FIP_VANTRUST_LIQUIDEZ_CORRIENTE"),
("O8FipVantrustLiquidezCortoPlazo","FIP Liquidez Corto Plazo","Liquidez","CORTO PLAZO","77.806.943-1","CLP","vigente",None,None,"FIP_VANTRUST_LIQUIDEZ_CORTO_PLAZO"),
("O9FipVantrustLiquidezDisponibleI","FIP Liquidez Disponible I","Liquidez","DISPONIBLE I","76.623.064-4","CLP","vigente","76623064-4 FIP Vantrust Liquidez Disponible I.pdf",None,"FIP_VANTRUST_LIQUIDEZ_DISPONIBLE_I"),
("O10FipVantrustLiquidezDolar","FIP Liquidez Dólar","Liquidez","DÓLAR","77.270.965-K","USD","vigente","77270965-K FIP Vantrust Liquidez Dolar.pdf","Prorroga FIP Vantrust Liquidez Dolar (2026-01).pdf","FIP_VANTRUST_LIQUIDEZ_DOLAR"),
("O11FipVantrustLiquidezDolarCaja","FIP Liquidez Dólar Caja","Liquidez","DÓLAR CAJA","77.697.015-8","USD","vigente","77697015-8 FIP Vantrust Liquidez Dolar Caja.pdf",None,"FIP_VANTRUST_LIQUIDEZ_DOLAR_CAJA"),
("O12FipVantrustLiquidezEfectivo","FIP Liquidez Efectivo","Liquidez","EFECTIVO","77.270.964-1","CLP","vigente","77270964-1 FIP Vantrust Liquidez Efectivo.pdf","Prorroga FIP Vantrust Liquidez Efectivo (2026-01).pdf","FIP_VANTRUST_LIQUIDEZ_EFECTIVO"),
("O13FipVantrustLiquidezFlexible","FIP Liquidez Flexible","Liquidez","FLEXIBLE","76.637.336-4","CLP","vigente","76637336-4 FIP Vantrust Liquidez Flexible.pdf",None,"FIP_VANTRUST_LIQUIDEZ_FLEXIBLE"),
("O14FipVantrustLiquidezI","FIP Liquidez I (Uno)","Liquidez","I","77.155.267-6","CLP","vigente",None,"Prorroga FIP Vantrust Liquidez I (2026-01).pdf","FIP_VANTRUST_LIQUIDEZ_I"),
("O15FipVantrustLiquidezLocal","FIP Liquidez Local","Liquidez","LOCAL","77.414.856-6","CLP","vigente","77414856-6 FIP Vantrust Liquidez Local.pdf",None,"FIP_VANTRUST_LIQUIDEZ_LOCAL"),
("O16FipVantrustLiquidezMonetarioI","FIP Liquidez Monetario I","Liquidez","MONETARIO I","76.623.036-9","CLP","vigente","76623036-9 FIP Vantrust Liquidez Monetario I.pdf",None,"FIP_VANTRUST_LIQUIDEZ_MONETARIO_I"),
("O17FipVantrustLiquidezPermanente","FIP Liquidez Permanente","Liquidez","PERMANENTE","77.806.942-3","CLP","vigente",None,None,"FIP_VANTRUST_LIQUIDEZ_PERMANENTE"),
("O18FipVantrustLiquidezPlus","FIP Liquidez Plus","Liquidez","PLUS","77.414.858-2","CLP","vigente","77414858-2 FIP Vantrust Liquidez Plus.pdf",None,"FIP_VANTRUST_LIQUIDEZ_PLUS"),
("O19FipVantrustLiquidezPresente","FIP Liquidez Presente","Liquidez","PRESENTE","77.414.855-8","CLP","vigente","77414855-8 FIP Vantrust Liquidez Presente.pdf",None,"FIP_VANTRUST_LIQUIDEZ_PRESENTE"),
("O20FipVantrustLiquidezRecurrente","FIP Liquidez Recurrente","Liquidez","RECURRENTE","76.639.712-3","CLP","vigente","76639712-3 FIP Vantrust Liquidez Recurrente.pdf",None,"FIP_VANTRUST_LIQUIDEZ_RECURRENTE"),
("O21FipVantrustLiquidezRendimiento","FIP Liquidez Rendimiento","Liquidez","RENDIMIENTO","77.270.963-3","CLP","vigente","77270963-3 FIP Vantrust Liquidez Rendimiento.pdf","Prorroga FIP Vantrust Liquidez Rendimiento (2026-01).pdf","FIP_VANTRUST_LIQUIDEZ_RENDIMIENTO"),
("O22FipVantrustLiquidezReservaDolar","FIP Liquidez Reserva Dólar","Liquidez","RESERVA DÓLAR","76.637.335-6","USD","vigente","76637335-6 FIP Vantrust Liquidez Reserva Dolar.pdf",None,"FIP_VANTRUST_LIQUIDEZ_RESERVA_DOLAR"),
("O23FipVantrustLiquidezSencillo","FIP Liquidez Sencillo","Liquidez","SENCILLO","76.933.993-0","CLP","vigente",None,"Prorroga FIP Vantrust Liquidez Sencillo (2026-01).pdf","FIP_VANTRUST_LIQUIDEZ_SENCILLO"),
("O24FipVantrustLiquidezTemporal","FIP Liquidez Temporal","Liquidez","TEMPORAL","76.639.716-6","CLP","vigente",None,None,"FIP_VANTRUST_LIQUIDEZ_TEMPORAL"),
("O25FipVantrustLiquidezFlexibleDolar","FIP Liquidez Flexible Dólar","Liquidez","FLEXIBLE DÓLAR","76.637.326-7","USD","vigente","76637326-7 FIP Vantrust Liquidez Flexible Dolar.pdf",None,"FIP_VANTRUST_LIQUIDEZ_FLEXIBLE_DOLAR"),
("O26FipVantrustLiquidezHorizonte","FIP Liquidez Horizonte","Liquidez","HORIZONTE","76.639.719-0","CLP","vigente","76639719-0 FIP Vantrust Liquidez Horizonte.pdf",None,"FIP_VANTRUST_LIQUIDEZ_HORIZONTE"),
# ── Fondos nuevos (reglamento vigente entregado; folleto aún no se genera) ──
("O27FipVantrustLiquidezEstrategico","FIP Liquidez Estratégico","Liquidez","ESTRATÉGICO","76.639.718-2","CLP","vigente","76639718-2 FIP Vantrust Liquidez Estrategico.pdf",None,None),
("O28FipVantrustLiquidezReserva","FIP Liquidez Reserva","Liquidez","RESERVA","76.639.715-8","CLP","vigente","76639715-8 FIP Vantrust Liquidez Reserva.pdf",None,None),
("O29FipVantrustLiquidezInmediata","FIP Liquidez Inmediata","Liquidez","INMEDIATA","76.650.244-K","CLP","vigente","76650244-K FIP Vantrust Liquidez Inmediata.pdf",None,None),
("O30FipVantrustLiquidezMercadoMonetario","FIP Liquidez Mercado Monetario","Liquidez","MERCADO MONETARIO","76.650.253-9","CLP","vigente","76650253-9 FIP Vantrust Liquidez Mercado Monetario.pdf",None,None),
("O31FipVantrustLiquidezRemanente","FIP Liquidez Remanente","Liquidez","REMANENTE","76.650.254-7","CLP","vigente","76650254-7 FIP Vantrust Liquidez Remanente.pdf",None,None),
("O32FipVantrustLiquidezFlotante","FIP Liquidez Flotante","Liquidez","FLOTANTE","76.650.256-3","CLP","vigente","76650256-3 FIP Vantrust Liquidez Flotante.pdf",None,None),
("O33FipVantrustLiquidezTransitorio","FIP Liquidez Transitorio","Liquidez","TRANSITORIO","76.650.237-7","CLP","vigente","76650237-7 FIP Vantrust Liquidez Transitorio.pdf",None,None),
("O34FipVantrustLiquidezAdicional","FIP Liquidez Adicional","Liquidez","ADICIONAL","76.650.243-1","CLP","vigente","76650243-1 FIP Vantrust Liquidez Adicional.pdf",None,None),
("O35FipVantrustLiquidezIncremental","FIP Liquidez Incremental","Liquidez","INCREMENTAL","76.650.242-3","CLP","vigente","76650242-3 FIP Vantrust Liquidez Incremental.pdf",None,None),
("O36FipVantrustLiquidezGestionCaja","FIP Liquidez Gestión Caja","Liquidez","GESTIÓN CAJA","76.650.252-0","CLP","vigente","76650252-0 FIP Vantrust Liquidez Gestion Caja.pdf",None,None),
("O37FipVantrustDeudaPrivada","FIP Deuda Privada","Deuda","PRIVADA","","CLP","vigente","FIP Vantrust Deuda Privada.pdf",None,None),
("O38FipVantrustTesoreria","FIP Tesorería","","TESORERÍA","","CLP","vigente","FIP Vantrust Tesoreria.pdf",None,None),
("O39FipVantrustUsdMoneyMarket","FIP USD Money Market","USD","MONEY MARKET","","USD","vencido","FIP Vantrust USD Money Market.pdf",None,None),
("O40FipVantrustLiquidezGranPatrimonio","FIP Liquidez Gran Patrimonio","Liquidez","GRAN PATRIMONIO","76.623.035-0","CLP","vencido","76623035-0 FIP Vantrust Liquidez Gran Patrimonio.pdf",None,None),
("O41FipVantrustLiquidezII","FIP Liquidez II","Liquidez","II","","CLP","vigente","FIP Vantrust Liquidez II.pdf",None,None),
("O42FipVantrustLiquidezDisponible","FIP Liquidez Disponible","Liquidez","DISPONIBLE","","CLP","vencido","FIP Vantrust Liquidez Disponible.pdf",None,None),
("O43FipVantrustLiquidezMonetario","FIP Liquidez Monetario","Liquidez","MONETARIO","","CLP","vencido","FIP Vantrust Liquidez Monetario.pdf",None,None),
("O44FipVantrustLiquidez","FIP Vantrust Liquidez","","LIQUIDEZ","","CLP","vencido","FIP Vantrust Liquidez.pdf",None,None),
("O45FipVantrustExtra","FIP Extra","","EXTRA","","CLP","vencido","FIP Vantrust Extra.pdf",None,None),
]
