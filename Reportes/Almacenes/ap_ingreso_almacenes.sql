USE [ADMIN_EMPRESA]
GO
/****** Object:  StoredProcedure [dbo].[ap_ingreso_almacenes]    Script Date: 03/30/2017 17:55:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Autor: Jorge Quintanilla Arancibia
-- Modificado por:
-- Sistema: ADMIN.EMPRESA
-- Objeto: PROCEDIMIENTO ALMACENADO
-- Nombre del objeto: ap_salida_almacen
-- Descripcion:	REPORTE - SALIDA ALMACEN 
-- Tabla: ao_ventas_cabecera
-- Tablas relacionadas: 
-- Fecha de creacion: 11/01/2017
-- Fecha de modificacion;
-- Version: 1.0
-- =============================================
ALTER PROCEDURE [dbo].[ap_ingreso_almacenes]
@compra_codigo INTEGER,
--@dia_correl INT,
@ges_gestion varchar(4)
AS
BEGIN

------------------////////////////////////////////////////////////////////-------------------------
--delete from ac_almacen_aux
---------------------------------Copia añ Auxiliar
--insert into ac_almacen_aux (fmes_plan, bien_codigo, cantidad, edif_descripcion, fecha_almi, doc_codigo ,doc_numero_m, observaciones2 )
--(SELECT venta_codigo, bien_codigo, bien_cantidad_por_empaque AS bien_cantidad, '' as edif_descripcion, fecha_verif, doc_codigo_alm, doc_numero_alm, observaciones
--FROM av_ventas_y_detalle WHERE venta_codigo = @venta_codigo and ges_gestion = @ges_gestion and par_codigo <> '43340')

----Actualiza unimed_codigo y bien_descripcion
--update ac_almacen_aux set ac_almacen_aux.unimed_codigo =ac_bienes.unimed_codigo, ac_almacen_aux.bien_descripcion =ac_bienes.bien_descripcion
--from ac_almacen_aux INNER JOIN ac_bienes 
--ON ac_almacen_aux.bien_codigo =ac_bienes.bien_codigo

-----------------------------------------////////////////////////////////////////////////

SELECT     dbo.ao_compra_cabecera.ges_gestion, dbo.ao_compra_cabecera.compra_codigo, dbo.ao_compra_cabecera.unidad_codigo, 
                      dbo.ao_compra_cabecera.solicitud_codigo, dbo.ao_compra_cabecera.edif_codigo, dbo.gc_edificaciones.edif_descripcion, dbo.ao_compra_cabecera.fecha_registro, 
                      gc_beneficiario_1.beneficiario_denominacion AS solicitante, dbo.gc_beneficiario.beneficiario_denominacion AS responsable, 
                      dbo.ao_compra_cabecera.doc_numero_alm, dbo.ao_compra_cabecera.doc_numero, dbo.ao_compra_cabecera.doc_codigo
FROM         dbo.ao_compra_cabecera INNER JOIN
                      dbo.gc_edificaciones ON dbo.ao_compra_cabecera.edif_codigo = dbo.gc_edificaciones.edif_codigo INNER JOIN
                      dbo.gc_beneficiario AS gc_beneficiario_1 ON dbo.ao_compra_cabecera.beneficiario_codigo_alm = gc_beneficiario_1.beneficiario_codigo INNER JOIN
                      dbo.gc_beneficiario ON dbo.ao_compra_cabecera.beneficiario_codigo_resp = dbo.gc_beneficiario.beneficiario_codigo
                      
WHERE     (dbo.ao_compra_cabecera.ges_gestion = @ges_gestion) AND (dbo.ao_compra_cabecera.compra_codigo = @compra_codigo )

	--GROUP BY dbo.gc_beneficiario.beneficiario_denominacion, dbo.ac_almacen_aux.bien_codigo, dbo.ac_almacen_aux.cantidad, dbo.ac_almacen_aux.edif_descripcion, 
    -- -                 dbo.ac_almacen_aux.fecha_almi, dbo.ac_almacen_aux.doc_codigo, dbo.ac_almacen_aux.doc_numero_m, dbo.ac_almacen_aux.observaciones2, 
    --                  dbo.ac_almacen_aux.unimed_codigo, dbo.ac_almacen_aux.bien_descripcion, dbo.ac_almacen_aux.fmes_plan

END
