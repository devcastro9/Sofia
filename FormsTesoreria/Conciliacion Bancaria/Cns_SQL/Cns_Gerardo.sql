
exec unionbancogtz '05/05/2000', '31/05/2000'


DECLARE @FechIni DateTime
DECLARE @FechFin DateTime

select * from 
fc_datosgtz


INSERT INTO fc_datosgtz(Nro_Cmpte, Organismo, Fecha_Pago, Monto, 
			Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Estado_Conciliacion)
SELECT pago_detalle.codigo_pago, pago_detalle.org_codigo, 
    pago_detalle.fecha_pago, pago_detalle.monto_Bolivianos, 
    pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, 
    pago_detalle.numero_cheque_trf, pago_detalle.cheque_o_trf, 
    pago_detalle.cta_codigo, fc_cuenta_bancaria.Bco_codigo, pago_detalle.Estado_Conciliacion
FROM pago_detalle INNER JOIN
    fc_cuenta_bancaria ON 
    pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo
WHERE pago_detalle.fecha_pago BETWEEN @FechIni AND @FechFin


INSERT INTO fc_datosgtz(Nro_Cmpte, Organismo, Fecha_Pago, Monto, 
			Cambio, Beneficiario, Nro_Doc, Transf_Cheq, Cta_Codigo, Bco_Codigo, Estado_Conciliacion)

SELECT org_codigo, codigo_beneficiario, cheque_o_trf, 
    Bco_codigo, numero_documento, cta_codigo, 
    monto_bolivianos, estado_conciliacion, tipo_cambio
FROM fo_ingresos
WHERE pago_detalle.fecha_pago BETWEEN @FechIni AND @FechFin