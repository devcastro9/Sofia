select p.ges_Gestion, p.codigo_unidad, p.codigo_solicitud, p.org_codigo, p.codigo_pago,d.codigo_beneficiario 
from pagos p, pago_detalle d 
where 	p.ges_Gestion	= d.ges_gestion	and
	p.org_codigo	= d.org_codigo	and
	p.codigo_pago	= d.codigo_pago	and
	p.tipo_formulario  ='COM'	and
	p.estado_compromiso='S'		and
	d.codigo_beneficiario in (select ci from rc_personal)

order by d.codigo_beneficiario

select p.ges_Gestion, p.codigo_unidad, p.codigo_solicitud, p.org_codigo, p.codigo_pago, d.codigo_beneficiario
from pagos p, pago_detalle d 
where 	p.ges_Gestion	= d.ges_gestion	and
	p.org_codigo	= d.org_codigo	and
	p.codigo_pago	= d.codigo_pago	and
	p.codigo_solicitud='0649'