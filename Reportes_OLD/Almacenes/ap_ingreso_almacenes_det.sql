USE [ADMIN_EMPRESA]਍䜀伀ഀഀ
/****** Object:  StoredProcedure [dbo].[ap_ingreso_almacenes_det]    Script Date: 03/30/2017 17:55:19 ******/਍匀䔀吀 䄀一匀䤀开一唀䰀䰀匀 伀一ഀഀ
GO਍匀䔀吀 儀唀伀吀䔀䐀开䤀䐀䔀一吀䤀䘀䤀䔀刀 伀一ഀഀ
GO਍ⴀⴀ 㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀ഀഀ
-- Autor: Jorge Quintanilla Arancibia਍ⴀⴀ 䴀漀搀椀昀椀挀愀搀漀 瀀漀爀㨀ഀഀ
-- Sistema: ADMIN.EMPRESA਍ⴀⴀ 伀戀樀攀琀漀㨀 倀刀伀䌀䔀䐀䤀䴀䤀䔀一吀伀 䄀䰀䴀䄀䌀䔀一䄀䐀伀ഀഀ
-- Nombre del objeto: ap_salida_almacen਍ⴀⴀ 䐀攀猀挀爀椀瀀挀椀漀渀㨀ऀ刀䔀倀伀刀吀䔀 ⴀ 匀䄀䰀䤀䐀䄀 䄀䰀䴀䄀䌀䔀一 ഀഀ
-- Tabla: ao_ventas_cabecera਍ⴀⴀ 吀愀戀氀愀猀 爀攀氀愀挀椀漀渀愀搀愀猀㨀 ഀഀ
-- Fecha de creacion: 11/01/2017਍ⴀⴀ 䘀攀挀栀愀 搀攀 洀漀搀椀昀椀挀愀挀椀漀渀㬀ഀഀ
-- Version: 1.0਍ⴀⴀ 㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀ഀഀ
ALTER PROCEDURE [dbo].[ap_ingreso_almacenes_det]਍䀀挀漀洀瀀爀愀开挀漀搀椀最漀 䤀一吀䔀䜀䔀刀Ⰰഀഀ
--@dia_correl INT,਍䀀最攀猀开最攀猀琀椀漀渀 瘀愀爀挀栀愀爀⠀㐀⤀ഀഀ
AS਍䈀䔀䜀䤀一ഀഀ
਍ⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀ⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀ⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀഀഀ
--delete from ac_almacen_aux਍ⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀ䌀漀瀀椀愀 愀 䄀甀砀椀氀椀愀爀ഀഀ
--insert into ac_almacen_aux (fmes_plan, bien_codigo, cantidad, edif_descripcion, fecha_almi, doc_codigo ,doc_numero_m, observaciones2 )਍ⴀⴀ⠀匀䔀䰀䔀䌀吀 瘀攀渀琀愀开挀漀搀椀最漀Ⰰ 戀椀攀渀开挀漀搀椀最漀Ⰰ 戀椀攀渀开挀愀渀琀椀搀愀搀开瀀漀爀开攀洀瀀愀焀甀攀 䄀匀 戀椀攀渀开挀愀渀琀椀搀愀搀Ⰰ ✀✀ 愀猀 攀搀椀昀开搀攀猀挀爀椀瀀挀椀漀渀Ⰰ 昀攀挀栀愀开瘀攀爀椀昀Ⰰ 搀漀挀开挀漀搀椀最漀开愀氀洀Ⰰ 搀漀挀开渀甀洀攀爀漀开愀氀洀Ⰰ 漀戀猀攀爀瘀愀挀椀漀渀攀猀ഀഀ
--FROM av_ventas_y_detalle WHERE venta_codigo = @venta_codigo and ges_gestion = @ges_gestion and par_codigo <> '43340')਍ഀഀ
----Actualiza unimed_codigo y bien_descripcion਍ⴀⴀ甀瀀搀愀琀攀 愀挀开愀氀洀愀挀攀渀开愀甀砀 猀攀琀 愀挀开愀氀洀愀挀攀渀开愀甀砀⸀甀渀椀洀攀搀开挀漀搀椀最漀 㴀愀挀开戀椀攀渀攀猀⸀甀渀椀洀攀搀开挀漀搀椀最漀Ⰰ 愀挀开愀氀洀愀挀攀渀开愀甀砀⸀戀椀攀渀开搀攀猀挀爀椀瀀挀椀漀渀 㴀愀挀开戀椀攀渀攀猀⸀戀椀攀渀开搀攀猀挀爀椀瀀挀椀漀渀ഀഀ
--from ac_almacen_aux INNER JOIN ac_bienes ਍ⴀⴀ伀一 愀挀开愀氀洀愀挀攀渀开愀甀砀⸀戀椀攀渀开挀漀搀椀最漀 㴀愀挀开戀椀攀渀攀猀⸀戀椀攀渀开挀漀搀椀最漀ഀഀ
਍ⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀⴀ⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀⼀ഀഀ
਍匀䔀䰀䔀䌀吀     搀戀漀⸀愀漀开挀漀洀瀀爀愀开搀攀琀愀氀氀攀⸀最攀猀开最攀猀琀椀漀渀Ⰰ 搀戀漀⸀愀漀开挀漀洀瀀爀愀开搀攀琀愀氀氀攀⸀挀漀洀瀀爀愀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀挀开戀椀攀渀攀猀⸀戀椀攀渀开搀攀猀挀爀椀瀀挀椀漀渀Ⰰ 搀戀漀⸀愀漀开挀漀洀瀀爀愀开搀攀琀愀氀氀攀⸀挀漀洀瀀爀愀开挀愀渀琀椀搀愀搀Ⰰ ഀഀ
                      dbo.ao_compra_detalle.compra_precio_unitario_bs, dbo.ao_compra_detalle.compra_precio_total_bs, dbo.ac_bienes_unidad_medida.unimed_descripcion਍䘀刀伀䴀         搀戀漀⸀愀漀开挀漀洀瀀爀愀开搀攀琀愀氀氀攀 䤀一一䔀刀 䨀伀䤀一ഀഀ
                      dbo.ac_bienes ON dbo.ao_compra_detalle.bien_codigo = dbo.ac_bienes.bien_codigo INNER JOIN਍                      搀戀漀⸀愀挀开戀椀攀渀攀猀开甀渀椀搀愀搀开洀攀搀椀搀愀 伀一 搀戀漀⸀愀挀开戀椀攀渀攀猀⸀甀渀椀洀攀搀开挀漀搀椀最漀 㴀 搀戀漀⸀愀挀开戀椀攀渀攀猀开甀渀椀搀愀搀开洀攀搀椀搀愀⸀甀渀椀洀攀搀开挀漀搀椀最漀ഀഀ
਍ഀഀ
WHERE     (dbo.ao_compra_detalle.ges_gestion = @ges_gestion) AND (dbo.ao_compra_detalle.compra_codigo = @compra_codigo )਍ഀഀ
	--GROUP BY dbo.gc_beneficiario.beneficiario_denominacion, dbo.ac_almacen_aux.bien_codigo, dbo.ac_almacen_aux.cantidad, dbo.ac_almacen_aux.edif_descripcion, ਍    ⴀⴀ ⴀ                 搀戀漀⸀愀挀开愀氀洀愀挀攀渀开愀甀砀⸀昀攀挀栀愀开愀氀洀椀Ⰰ 搀戀漀⸀愀挀开愀氀洀愀挀攀渀开愀甀砀⸀搀漀挀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀挀开愀氀洀愀挀攀渀开愀甀砀⸀搀漀挀开渀甀洀攀爀漀开洀Ⰰ 搀戀漀⸀愀挀开愀氀洀愀挀攀渀开愀甀砀⸀漀戀猀攀爀瘀愀挀椀漀渀攀猀㈀Ⰰ ഀഀ
    --                  dbo.ac_almacen_aux.unimed_codigo, dbo.ac_almacen_aux.bien_descripcion, dbo.ac_almacen_aux.fmes_plan਍ഀഀ
END਍�