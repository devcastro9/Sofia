select p.codigo_unidad, p.codigo_solicitud, p.org_codigo, p.codigo_pago, d.codigo_beneficiario, P.JUSTIFICACION  
from pagos p, pago_detalle d
where 	p.ges_gestion	= d.ges_gestion and
	p.org_codigo	= d.org_codigo	and
	p.codigo_pago	= d.codigo_pago	and
	p.tipo_formulario= 'COM'	and
	p.codigo_solicitud= '185'	and
	p.codigo_unidad='UCPE'
order by p.codigo_unidad, p.codigo_solicitud, p.org_codigo DESC, p.codigo_pago, d.codigo_beneficiario  

--select * from pago_detalle where org_codigo='111' and codigo_pago=1115

edRelacionadorAutomatico 'DEV'
edRelacionadorAutomatico 'CYD'

select count(*) from ac_ben_comprdeven where usr_usuario='EMA'

delete ac_ben_comprdeven where usr_usuario='EMA'

SELECT * FROM FC_BENEFICIARIO WHERE CODIGO_BENEFICIARIO='2340367'