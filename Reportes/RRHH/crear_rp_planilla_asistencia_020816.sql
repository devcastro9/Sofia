-- =============================================  
-- Autor: Jose Luis Rodriguez Ch.  
-- Modificado por:   
-- Sistema: ADMIN.EMPRESA  
-- Objeto: PROCEDIMIENTO ALMACENADO  
-- Nombre del objeto:rp_planilla_asistencia  
-- Descripcion: PROCEDIMIENTO ALMACENADO PARA INSERTAR TABLA  
-- Tabla:   
-- Tablas relacionadas:  
-- Modificado por:  
-- Fecha de creacion: 02/08/2016  
-- Fecha de modificacion:  
-- Version: 1.0  
-- ================================================  
--exec rp_planilla_asistencia '2016', 'P01','7'
CREATE PROCEDURE [dbo].[rp_planilla_asistencia]    
@ges_gestion varchar(4),  
@planilla_codigo varchar(3),  
@mes_grupo int  
AS  
BEGIN  
SELECT DISTINCT      
                     asis.Fecha_control,  
                     gc_beneficiario.beneficiario_codigo As ci ,  
                     (gc_beneficiario.beneficiario_primer_apellido + ' ' + gc_beneficiario.beneficiario_segundo_apellido + ' ' + gc_beneficiario.beneficiario_nombres) As nombre,  
                     asis.TipoHorario,  
                     asis.HoraUno As horaentrada,  
                     asis.HoraDos As horasalida,  
                     asis.HoraTres As marcaentrada,  
                     asis.HoraCuatro As marcasalida,  
                     asis.Tardanza As retraso,  
                     asis.TiemAsist As tiempotrabajo,  
                     ISNULL(CASE asis.EsFalta WHEN 1 THEN 'SI' ELSE 'NO' END,'NO') As esfalta  
FROM         dbo.gc_beneficiario INNER JOIN  
                      dbo.ro_personal_contratado ON dbo.gc_beneficiario.beneficiario_codigo = dbo.ro_personal_contratado.beneficiario_codigo INNER JOIN  
                      dbo.ro_pagos_cronograma_Detalle ON dbo.gc_beneficiario.beneficiario_codigo = dbo.ro_pagos_cronograma_Detalle.beneficiario_codigo INNER JOIN  
                      dbo.rc_cargos ON dbo.ro_personal_contratado.cargo_codigo = dbo.rc_cargos.cargo_codigo INNER JOIN  
                      dbo.rc_puestos ON dbo.ro_personal_contratado.puesto_codigo = dbo.rc_puestos.puesto_codigo INNER JOIN  
                      dbo.ro_pagos_cronograma ON dbo.ro_pagos_cronograma_Detalle.ges_gestion = dbo.ro_pagos_cronograma.ges_gestion INNER JOIN  
                      dbo.ro_pagos_cronograma AS ro_pagos_cronograma_2 ON dbo.ro_pagos_cronograma_Detalle.ges_gestion = ro_pagos_cronograma_2.ges_gestion INNER JOIN  
                      dbo.gc_unidad_ejecutora ON dbo.ro_personal_contratado.unidad_codigo = dbo.gc_unidad_ejecutora.unidad_codigo INNER JOIN  
                      dbo.rc_planilla_grupo ON dbo.ro_pagos_cronograma_Detalle.planilla_codigo = dbo.rc_planilla_grupo.planilla_codigo                        
                      -- Control de asistencia  
                     INNER JOIN ro_ControlAsistencia As asis ON asis.beneficiario_codigo = gc_beneficiario.beneficiario_codigo  
               WHERE ( YEAR(asis.fecha_control) = @ges_gestion 
				AND MONTH(asis.fecha_control) = @mes_grupo
                 and ro_pagos_cronograma_Detalle.planilla_codigo like @planilla_codigo )                                                  
                --( ro_pagos_cronograma_Detalle.ges_gestion like @ges_gestion  and ro_pagos_cronograma_Detalle.planilla_codigo like @planilla_codigo and ro_pagos_cronograma_Detalle.mes_grupo like '%' )                                                  
     
 END             
 
 
 
 --select * from ro_ControlAsistencia