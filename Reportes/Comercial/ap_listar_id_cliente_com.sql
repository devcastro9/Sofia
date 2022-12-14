USE [ADMIN_EMPRESA]
GO
/****** Object:  StoredProcedure [dbo].[ap_listar_id_cliente_com]    Script Date: 04/03/2017 11:42:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Autor: DANIEL GUSTAVO LAURA PAREDES
-- Modificado por:
-- Sistema: ADMIN.EMPRESA
-- Objeto: PROCEDIMIENTO ALMACENADO
-- Nombre del objeto: ap_listar_id_cliente_com
-- Descripcion:	PROCEDIMIENTO ALMACENADO PARA LISTAR
-- Tabla: ao_solicitud
-- Tablas relacionadas: 
-- Fecha de creacion: 09/09/2016
-- Fecha de modificacion; 
-- Version: 1.0
-- =============================================
ALTER PROCEDURE [dbo].[ap_listar_id_cliente_com]
(@unidad_codigo	varchar(15), @FFInicio varchar(10),   @FFFinal varchar(10)
)
AS
DECLARE @FInicio DateTime
SET @FInicio = CONVERT(DateTime, @FFInicio, 103)
DECLARE @FFinal DateTime
SET @FFinal = CONVERT(DateTime, @FFFinal, 103)

SELECT     dbo.gc_edificaciones.beneficiario_codigo AS beneficiario_codigo_edif, dbo.gc_edificaciones.edif_referencia, dbo.gc_edificaciones.edif_nro, 
                      dbo.gc_edificaciones.calle_codigo, dbo.gc_edificaciones.zona_codigo, dbo.gc_edificaciones.munic_codigo, dbo.gc_edificaciones.prov_codigo, 
                      dbo.gc_edificaciones.depto_codigo, dbo.ao_solicitud.ges_gestion, dbo.ao_solicitud.unidad_codigo, dbo.ao_solicitud.solicitud_codigo, 
                      dbo.ao_solicitud.solicitud_fecha_solicitud, dbo.ao_solicitud.solicitud_fecha_recepción, dbo.ao_solicitud.solicitud_tipo, dbo.ao_solicitud.edif_codigo, 
                      dbo.ao_solicitud.beneficiario_codigo, dbo.ao_solicitud.beneficiario_codigo_resp, dbo.ao_solicitud.beneficiario_codigo_resp2, dbo.ao_solicitud.unidad_codigo_sol, 
                      dbo.ao_solicitud.solicitud_justificacion, dbo.ao_solicitud.solicitud_observaciones, dbo.ao_solicitud.proceso_codigo, dbo.ao_solicitud.subproceso_codigo, 
                      dbo.ao_solicitud.etapa_codigo, dbo.ao_solicitud.etapa_codigo2, dbo.ao_solicitud.clasif_codigo, dbo.ao_solicitud.doc_codigo, dbo.ao_solicitud.doc_codigo2, 
                      dbo.ao_solicitud.doc_numero, dbo.ao_solicitud.doc_numero2, dbo.ao_solicitud.poa_codigo, dbo.ao_solicitud.ges_gestion_ant, dbo.ao_solicitud.unidad_codigo_ant, 
                      dbo.ao_solicitud.solicitud_codigo_ant, dbo.ao_solicitud.correl_detalle, dbo.ao_solicitud.correl_edificacion, dbo.ao_solicitud.correl_calculo, 
                      dbo.ao_solicitud.correl_persona, dbo.ao_solicitud.correl_cotiza, dbo.ao_solicitud.correl_bitacora, dbo.ao_solicitud.archivo_respaldo, 
                      dbo.ao_solicitud.archivo_respaldo_cargado, dbo.ao_solicitud.estado_codigo, dbo.ao_solicitud.estado_etapa2, dbo.ao_solicitud.estado_cotiza, 
                      dbo.ao_solicitud.fecha_registro, dbo.ao_solicitud.hora_registro, dbo.ao_solicitud.usr_codigo, dbo.ao_solicitud.usr_codigo_aprueba, dbo.ao_solicitud.fecha_aprueba, 
                      dbo.ao_solicitud.hora_aprueba, dbo.ao_solicitud.fecha_registro2, dbo.ao_solicitud.usr_codigo2, dbo.gc_provincia.prov_descripcion, 
                      dbo.gc_zonas.zona_denominacion, dbo.gc_beneficiario.beneficiario_denominacion, dbo.gc_calles.calle_denominacion, dbo.gc_municipio.munic_descripcion, 
                      dbo.gc_departamento.depto_descripcion, dbo.gc_unidad_ejecutora.unidad_descripcion, dbo.gc_tipo_solicitud.solicitud_tipo_descripcion, 
                      gc_beneficiario_1.beneficiario_denominacion AS beneficiario_denominacion_edif, dbo.gc_calles.calle_tipo, dbo.gc_edificaciones.edif_descripcion, 
                      dbo.gc_edificaciones.edif_tipo
FROM         dbo.ao_solicitud INNER JOIN
                      dbo.gc_unidad_ejecutora ON dbo.ao_solicitud.unidad_codigo = dbo.gc_unidad_ejecutora.unidad_codigo LEFT OUTER JOIN
                      dbo.gc_tipo_solicitud ON dbo.ao_solicitud.solicitud_tipo = dbo.gc_tipo_solicitud.solicitud_tipo LEFT OUTER JOIN
                      dbo.gc_edificaciones ON dbo.ao_solicitud.edif_codigo = dbo.gc_edificaciones.edif_codigo LEFT OUTER JOIN
                      dbo.gc_departamento ON dbo.gc_edificaciones.depto_codigo = dbo.gc_departamento.depto_codigo LEFT OUTER JOIN
                      dbo.gc_beneficiario ON dbo.ao_solicitud.beneficiario_codigo = dbo.gc_beneficiario.beneficiario_codigo LEFT OUTER JOIN
                      dbo.gc_municipio ON dbo.gc_edificaciones.munic_codigo = dbo.gc_municipio.munic_codigo LEFT OUTER JOIN
                      dbo.gc_provincia ON dbo.gc_edificaciones.prov_codigo = dbo.gc_provincia.prov_codigo LEFT OUTER JOIN
                      dbo.gc_calles ON dbo.gc_edificaciones.zona_codigo = dbo.gc_calles.zona_codigo AND 
                      dbo.gc_edificaciones.calle_codigo = dbo.gc_calles.calle_codigo LEFT OUTER JOIN
                      dbo.gc_zonas ON dbo.gc_edificaciones.zona_codigo = dbo.gc_zonas.zona_codigo LEFT OUTER JOIN
                      dbo.gc_beneficiario AS gc_beneficiario_1 ON dbo.gc_edificaciones.beneficiario_codigo = gc_beneficiario_1.beneficiario_codigo
WHERE     (dbo.ao_solicitud.unidad_codigo = 'DVTA') and  (dbo.ao_solicitud.solicitud_fecha_solicitud BETWEEN CONVERT(DATETIME, @FInicio, 103) AND CONVERT(DATETIME, @FFinal, 103)) 