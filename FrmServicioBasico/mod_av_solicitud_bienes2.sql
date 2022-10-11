CREATE VIEW dbo.av_solicitud_bienes2    
AS    
SELECT     dbo.ao_solicitud_bienes.ges_gestion, dbo.ao_solicitud_bienes.unidad_codigo, dbo.ao_solicitud_bienes.solicitud_codigo, dbo.ao_solicitud_bienes.bien_codigo,     
                      dbo.ao_solicitud_bienes.grupo_codigo, dbo.ao_solicitud_bienes.subgrupo_codigo, dbo.ao_solicitud_bienes.par_codigo, dbo.ao_solicitud_bienes.marca_codigo,     
                      dbo.ao_solicitud_bienes.modelo_codigo, dbo.ao_solicitud_bienes.bien_cantidad, dbo.ao_solicitud_bienes.bien_precio_compra,     
                      dbo.ao_solicitud_bienes.bien_total_compra, dbo.ao_solicitud_bienes.bien_precio_venta_base, dbo.ao_solicitud_bienes.bien_total_venta,     
                      dbo.ao_solicitud_bienes.tipo_moneda, dbo.ao_solicitud_bienes.unimed_codigo, dbo.ao_solicitud_bienes.unimed_codigo_empaque,     
                      dbo.ao_solicitud_bienes.bien_cantidad_por_empaque, dbo.ao_solicitud_bienes.venta_o_compra, dbo.ao_solicitud_bienes.fosa_dimension_frente,     
                      dbo.ao_solicitud_bienes.fosa_dimension_fondo, dbo.ao_solicitud_bienes.estado_codigo, dbo.ao_solicitud_bienes.usr_codigo,     
                      dbo.ao_solicitud_bienes.fecha_registro, dbo.ao_solicitud_bienes.hora_registro, dbo.ac_bienes.bien_descripcion, dbo.ac_bienes.bien_codigo_anterior,     
                      dbo.ac_bienes.bien_codigo_universal, dbo.ac_bienes.bien_descripcion_anterior, dbo.ac_bienes.pais_codigo, dbo.ac_bienes.edif_codigo,     
                      dbo.ac_bienes.estado_codigo AS estado_codigo_bien    
FROM         dbo.ac_bienes INNER JOIN    
                      dbo.ao_solicitud_bienes ON dbo.ac_bienes.bien_codigo = dbo.ao_solicitud_bienes.bien_codigo AND     
                      dbo.ac_bienes.grupo_codigo = dbo.ao_solicitud_bienes.grupo_codigo    
--WHERE     (dbo.ao_solicitud_bienes.grupo_codigo = '30000') AND (dbo.ac_bienes.estado_codigo = 'APR') OR    
--                      (dbo.ao_solicitud_bienes.grupo_codigo = '20000') AND (dbo.ac_bienes.estado_codigo = 'APR') 