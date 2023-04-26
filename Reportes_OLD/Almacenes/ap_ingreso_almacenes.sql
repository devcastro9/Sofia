USE [ADMIN_EMPRESA]਍䜀伀ഀഀ
/****** Object:  StoredProcedure [dbo].[ap_ingreso_almacenes]    Script Date: 03/30/2017 17:55:06 ******/਍匀䔀吀 䄀一匀䤀开一唀䰀䰀匀 伀一ഀഀ
GO਍匀䔀吀 儀唀伀吀䔀䐀开䤀䐀䔀一吀䤀䘀䤀䔀刀 伀一ഀഀ
GO਍ⴀⴀ 㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀ഀഀ
-- Autor: Jorge Quintanilla Arancibia਍ⴀⴀ 䴀漀搀椀昀椀挀愀搀漀 瀀漀爀㨀ഀഀ
-- Sistema: ADMIN.EMPRESA਍ⴀⴀ 伀戀樀攀琀漀㨀 倀刀伀䌀䔀䐀䤀䴀䤀䔀一吀伀 䄀䰀䴀䄀䌀䔀一䄀䐀伀ഀഀ
-- Nombre del objeto: ap_salida_almacen਍ⴀⴀ 䐀攀猀挀爀椀瀀挀椀漀渀㨀ऀ刀䔀倀伀刀吀䔀 ⴀ 匀䄀䰀䤀䐀䄀 䄀䰀䴀䄀䌀䔀一 ഀഀ
-- Tabla: ao_ventas_cabecera਍ⴀⴀ 吀愀戀氀愀猀 爀攀氀愀挀椀漀渀愀搀愀猀㨀 ഀഀ
-- Fecha de creacion: 11/01/2017਍ⴀⴀ 䘀攀挀栀愀 搀攀 洀漀搀椀昀椀挀愀挀椀漀渀㬀ഀഀ
-- Version: 1.0਍ⴀⴀ 㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀ഀഀ
ALTER PROCEDURE [dbo].[ap_ingreso_almacenes]਍䀀挀漀洀瀀爀愀开挀漀搀椀最漀 䤀一吀䔀䜀䔀刀Ⰰഀഀ
--@dia_correl INT,਍䀀最攀猀开最攀猀琀椀漀渀 瘀愀爀挀栀愀爀⠀㐀⤀ഀഀ
AS਍䈀䔀䜀䤀一ഀഀ
਍ⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀ⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀ⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀഀഀ
--delete from ac_almacen_aux਍ⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀ䌀漀瀀椀愀 愀 䄀甀砀椀氀椀愀爀ഀഀ
--insert into ac_almacen_aux (fmes_plan, bien_codigo, cantidad, edif_descripcion, fecha_almi, doc_codigo ,doc_numero_m, observaciones2 )਍ⴀⴀ⠀匀䔀䰀䔀䌀吀 瘀攀渀琀愀开挀漀搀椀最漀Ⰰ 戀椀攀渀开挀漀搀椀最漀Ⰰ 戀椀攀渀开挀愀渀琀椀搀愀搀开瀀漀爀开攀洀瀀愀焀甀攀 䄀匀 戀椀攀渀开挀愀渀琀椀搀愀搀Ⰰ ✀✀ 愀猀 攀搀椀昀开搀攀猀挀爀椀瀀挀椀漀渀Ⰰ 昀攀挀栀愀开瘀攀爀椀昀Ⰰ 搀漀挀开挀漀搀椀最漀开愀氀洀Ⰰ 搀漀挀开渀甀洀攀爀漀开愀氀洀Ⰰ 漀戀猀攀爀瘀愀挀椀漀渀攀猀ഀഀ
--FROM av_ventas_y_detalle WHERE venta_codigo = @venta_codigo and ges_gestion = @ges_gestion and par_codigo <> '43340')਍ഀഀ
----Actualiza unimed_codigo y bien_descripcion਍ⴀⴀ甀瀀搀愀琀攀 愀挀开愀氀洀愀挀攀渀开愀甀砀 猀攀琀 愀挀开愀氀洀愀挀攀渀开愀甀砀⸀甀渀椀洀攀搀开挀漀搀椀最漀 㴀愀挀开戀椀攀渀攀猀⸀甀渀椀洀攀搀开挀漀搀椀最漀Ⰰ 愀挀开愀氀洀愀挀攀渀开愀甀砀⸀戀椀攀渀开搀攀猀挀爀椀瀀挀椀漀渀 㴀愀挀开戀椀攀渀攀猀⸀戀椀攀渀开搀攀猀挀爀椀瀀挀椀漀渀ഀഀ
--from ac_almacen_aux INNER JOIN ac_bienes ਍ⴀⴀ伀一 愀挀开愀氀洀愀挀攀渀开愀甀砀⸀戀椀攀渀开挀漀搀椀最漀 㴀愀挀开戀椀攀渀攀猀⸀戀椀攀渀开挀漀搀椀最漀ഀഀ
਍ⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀ⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀ഀഀ
਍匀䔀䰀䔀䌀吀     搀戀漀⸀愀漀开挀漀洀瀀爀愀开挀愀戀攀挀攀爀愀⸀最攀猀开最攀猀琀椀漀渀Ⰰ 搀戀漀⸀愀漀开挀漀洀瀀爀愀开挀愀戀攀挀攀爀愀⸀挀漀洀瀀爀愀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀漀开挀漀洀瀀爀愀开挀愀戀攀挀攀爀愀⸀甀渀椀搀愀搀开挀漀搀椀最漀Ⰰ ഀഀ
                      dbo.ao_compra_cabecera.solicitud_codigo, dbo.ao_compra_cabecera.edif_codigo, dbo.gc_edificaciones.edif_descripcion, dbo.ao_compra_cabecera.fecha_registro, ਍                      最挀开戀攀渀攀昀椀挀椀愀爀椀漀开㄀⸀戀攀渀攀昀椀挀椀愀爀椀漀开搀攀渀漀洀椀渀愀挀椀漀渀 䄀匀 猀漀氀椀挀椀琀愀渀琀攀Ⰰ 搀戀漀⸀最挀开戀攀渀攀昀椀挀椀愀爀椀漀⸀戀攀渀攀昀椀挀椀愀爀椀漀开搀攀渀漀洀椀渀愀挀椀漀渀 䄀匀 爀攀猀瀀漀渀猀愀戀氀攀Ⰰ ഀഀ
                      dbo.ao_compra_cabecera.doc_numero_alm, dbo.ao_compra_cabecera.doc_numero, dbo.ao_compra_cabecera.doc_codigo਍䘀刀伀䴀         搀戀漀⸀愀漀开挀漀洀瀀爀愀开挀愀戀攀挀攀爀愀 䤀一一䔀刀 䨀伀䤀一ഀഀ
                      dbo.gc_edificaciones ON dbo.ao_compra_cabecera.edif_codigo = dbo.gc_edificaciones.edif_codigo INNER JOIN਍                      搀戀漀⸀最挀开戀攀渀攀昀椀挀椀愀爀椀漀 䄀匀 最挀开戀攀渀攀昀椀挀椀愀爀椀漀开㄀ 伀一 搀戀漀⸀愀漀开挀漀洀瀀爀愀开挀愀戀攀挀攀爀愀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀开愀氀洀 㴀 最挀开戀攀渀攀昀椀挀椀愀爀椀漀开㄀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀 䤀一一䔀刀 䨀伀䤀一ഀഀ
                      dbo.gc_beneficiario ON dbo.ao_compra_cabecera.beneficiario_codigo_resp = dbo.gc_beneficiario.beneficiario_codigo਍                      ഀഀ
WHERE     (dbo.ao_compra_cabecera.ges_gestion = @ges_gestion) AND (dbo.ao_compra_cabecera.compra_codigo = @compra_codigo )਍ഀഀ
	--GROUP BY dbo.gc_beneficiario.beneficiario_denominacion, dbo.ac_almacen_aux.bien_codigo, dbo.ac_almacen_aux.cantidad, dbo.ac_almacen_aux.edif_descripcion, ਍    ⴀⴀ ⴀ                 搀戀漀⸀愀挀开愀氀洀愀挀攀渀开愀甀砀⸀昀攀挀栀愀开愀氀洀椀Ⰰ 搀戀漀⸀愀挀开愀氀洀愀挀攀渀开愀甀砀⸀搀漀挀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀挀开愀氀洀愀挀攀渀开愀甀砀⸀搀漀挀开渀甀洀攀爀漀开洀Ⰰ 搀戀漀⸀愀挀开愀氀洀愀挀攀渀开愀甀砀⸀漀戀猀攀爀瘀愀挀椀漀渀攀猀㈀Ⰰ ഀഀ
    --                  dbo.ac_almacen_aux.unimed_codigo, dbo.ac_almacen_aux.bien_descripcion, dbo.ac_almacen_aux.fmes_plan਍ഀഀ
END਍�