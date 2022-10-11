SELECT a.tipo_comp, b.cta_codigo, 
sum(b.monto_bolivianos) montoBs, 
sum(b.monto_dolares) montoUs
FROM pagos a, pago_detalle b
WHERE a.org_codigo = b.org_codigo
 AND a.ges_gestion = b.ges_gestion 
 AND a.codigo_pago = b.codigo_pago
 AND a.estado_pagado = 'S'
GROUP BY a.tipo_comp, b.cta_codigo

SELECT a.cta_codigo, 
sum(a.monto_bolivianos) montoBs, 
sum(a.monto_dolares) montoUs
FROM fo_ingresos a
WHERE a.estado_aprobacion = 'S'
GROUP BY a.cta_codigo

SELECT a.d_cta_larga, 
sum(a.d_montoBs) montoBs, 
sum(a.d_montoDl) montoUs
FROM co_diario a, co_comprobante_m b
WHERE a.cod_comp = b.cod_comp
  and b.tipo_comp = 'TRP'
  and b.status	= 'S'
GROUP BY a.d_cta_larga

SELECT a.h_cta_larga, 
sum(a.h_montoBs) montoBs, 
sum(a.h_montoDl) montoUs
FROM co_diario a, co_comprobante_m b
WHERE a.cod_comp = b.cod_comp
  and b.tipo_comp = 'TRP'
  and b.status	= 'S'
GROUP BY a.h_cta_larga


SELECT a.d_cta_larga, 
sum(a.d_montoBs) montoBs, 
sum(a.d_montoDl) montoUs
FROM co_diario a, co_comprobante_m b
WHERE a.cod_comp = b.cod_comp
  and b.tipo_comp = 'PCO'
  and b.status	= 'S'
  and a.d_cuenta = '1111' 
  and a.d_subcta1 = '02' 
GROUP BY a.d_cta_larga

SELECT a.h_cta_larga, 
sum(isnull(a.h_montoBs,0)) montoBs, 
sum(isnull(a.h_montoDl,0)) montoUs
FROM co_diario a, co_comprobante_m b
WHERE a.cod_comp = b.cod_comp
  and b.tipo_comp = 'PCO'
  and b.status	= 'S'
  and a.h_cuenta = '1111' 
  and a.h_subcta1 = '02' 
GROUP BY a.h_cta_larga

-------CAM
SELECT a.d_cta_larga, 
sum(a.d_montoBs) montoBs, 
sum(a.d_montoDl) montoUs
FROM co_diario a, co_comprobante_m b
WHERE a.cod_comp = b.cod_comp
  and b.tipo_comp = 'CAM'
  and b.status	= 'S'
  and a.d_cuenta = '1111' 
  and a.d_subcta1 = '02' 
GROUP BY a.d_cta_larga

---PCE
SELECT a.d_cta_larga, 
sum(a.d_montoBs) montoBs, 
sum(a.d_montoDl) montoUs
FROM co_diario a, co_comprobante_m b
WHERE a.cod_comp = b.cod_comp
  and b.tipo_comp = 'PCE'
  and b.status	= 'S'
  and a.d_cuenta = '1111' 
  and a.d_subcta1 = '02' 
GROUP BY a.d_cta_larga


SELECT a.h_cta_larga, 
sum(a.h_montoBs) montoBs, 
sum(a.h_montoDl) montoUs
FROM co_diario a, co_comprobante_m b
WHERE a.cod_comp = b.cod_comp
  and b.tipo_comp = 'PCE'
  and b.status	= 'S'
  and a.h_cuenta = '1111' 
  and a.h_subcta1 = '02' 
GROUP BY a.h_cta_larga


SELECT a.d_cta_larga, 
sum(a.d_montoBs) montoBs, 
sum(a.d_montoDl) montoUs
FROM co_diario a, co_comprobante_m b
WHERE a.cod_comp = b.cod_comp
  and b.tipo_comp = 'ANC'
  and b.status	= 'S'
  and a.d_cuenta = '1111' 
  and a.d_subcta1 = '02' 
GROUP BY a.d_cta_larga

EXEC ts_mf_ActualizaCtaBancaria
  SELECT a.cta_codigo, 
    sum(ISNULL(a.monto_bolivianos,0)) montoBs, 
    sum(ISNULL(a.monto_dolares,0)) montoUs
  FROM fo_ingresos a
  WHERE a.estado_aprobacion = 'S'
     AND a.estado_recaudado = 'S'
     AND a.estado_anulado is null
  GROUP BY a.cta_codigo

exec po_generacion_cod_poa_dev '01/01/2000', '01/01/2000', 'GEN_POA_APROX'

select *
from tmp_poa_devengado_real
where nivel_error = 0

select *
from tmp_poa_devengado_real
where nivel_error > 2

  SELECT
count(*)
  FROM pagos a, detalle_soes b,  soes c, soes_cab d
   WHERE b.dso_rechazado = 'No' 
    AND a.codigo_pago 	= b.codigo_pago 
    AND a.org_codigo	= b.org_codigo 
    AND a.ges_gestion	= b.ges_gestion
    AND b.soc_nro_sol	= c.soc_nro_sol
    AND b.soe_cod_convenio = c.soe_cod_convenio
    AND b.soe_nro_sec	= c.soe_nro_sec
    AND c.soc_nro_sol	= d.soc_nro_sol
    AND c.soe_cod_convenio = d.soc_codigo_convenio

SELECT *
from detalle_soes