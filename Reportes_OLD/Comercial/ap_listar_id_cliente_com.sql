USE [ADMIN_EMPRESA]਍䜀伀ഀഀ
/****** Object:  StoredProcedure [dbo].[ap_listar_id_cliente_com]    Script Date: 04/03/2017 11:42:15 ******/਍匀䔀吀 䄀一匀䤀开一唀䰀䰀匀 伀一ഀഀ
GO਍匀䔀吀 儀唀伀吀䔀䐀开䤀䐀䔀一吀䤀䘀䤀䔀刀 伀一ഀഀ
GO਍ⴀⴀ 㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀ഀഀ
-- Autor: DANIEL GUSTAVO LAURA PAREDES਍ⴀⴀ 䴀漀搀椀昀椀挀愀搀漀 瀀漀爀㨀ഀഀ
-- Sistema: ADMIN.EMPRESA਍ⴀⴀ 伀戀樀攀琀漀㨀 倀刀伀䌀䔀䐀䤀䴀䤀䔀一吀伀 䄀䰀䴀䄀䌀䔀一䄀䐀伀ഀഀ
-- Nombre del objeto: ap_listar_id_cliente_com਍ⴀⴀ 䐀攀猀挀爀椀瀀挀椀漀渀㨀ऀ倀刀伀䌀䔀䐀䤀䴀䤀䔀一吀伀 䄀䰀䴀䄀䌀䔀一䄀䐀伀 倀䄀刀䄀 䰀䤀匀吀䄀刀ഀഀ
-- Tabla: ao_solicitud਍ⴀⴀ 吀愀戀氀愀猀 爀攀氀愀挀椀漀渀愀搀愀猀㨀 ഀഀ
-- Fecha de creacion: 09/09/2016਍ⴀⴀ 䘀攀挀栀愀 搀攀 洀漀搀椀昀椀挀愀挀椀漀渀㬀 ഀഀ
-- Version: 1.0਍ⴀⴀ 㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀㴀ഀഀ
ALTER PROCEDURE [dbo].[ap_listar_id_cliente_com]਍⠀䀀甀渀椀搀愀搀开挀漀搀椀最漀ऀ瘀愀爀挀栀愀爀⠀㄀㔀⤀Ⰰ 䀀䘀䘀䤀渀椀挀椀漀 瘀愀爀挀栀愀爀⠀㄀　⤀Ⰰ   䀀䘀䘀䘀椀渀愀氀 瘀愀爀挀栀愀爀⠀㄀　⤀ഀഀ
)਍䄀匀ഀഀ
DECLARE @FInicio DateTime਍匀䔀吀 䀀䘀䤀渀椀挀椀漀 㴀 䌀伀一嘀䔀刀吀⠀䐀愀琀攀吀椀洀攀Ⰰ 䀀䘀䘀䤀渀椀挀椀漀Ⰰ ㄀　㌀⤀ഀഀ
DECLARE @FFinal DateTime਍匀䔀吀 䀀䘀䘀椀渀愀氀 㴀 䌀伀一嘀䔀刀吀⠀䐀愀琀攀吀椀洀攀Ⰰ 䀀䘀䘀䘀椀渀愀氀Ⰰ ㄀　㌀⤀ഀഀ
਍匀䔀䰀䔀䌀吀     搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀 䄀匀 戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀开攀搀椀昀Ⰰ 搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀攀搀椀昀开爀攀昀攀爀攀渀挀椀愀Ⰰ 搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀攀搀椀昀开渀爀漀Ⰰ ഀഀ
                      dbo.gc_edificaciones.calle_codigo, dbo.gc_edificaciones.zona_codigo, dbo.gc_edificaciones.munic_codigo, dbo.gc_edificaciones.prov_codigo, ਍                      搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀搀攀瀀琀漀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀最攀猀开最攀猀琀椀漀渀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀甀渀椀搀愀搀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀猀漀氀椀挀椀琀甀搀开挀漀搀椀最漀Ⰰ ഀഀ
                      dbo.ao_solicitud.solicitud_fecha_solicitud, dbo.ao_solicitud.solicitud_fecha_recepción, dbo.ao_solicitud.solicitud_tipo, dbo.ao_solicitud.edif_codigo, ਍                      搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀开爀攀猀瀀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀开爀攀猀瀀㈀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀甀渀椀搀愀搀开挀漀搀椀最漀开猀漀氀Ⰰ ഀഀ
                      dbo.ao_solicitud.solicitud_justificacion, dbo.ao_solicitud.solicitud_observaciones, dbo.ao_solicitud.proceso_codigo, dbo.ao_solicitud.subproceso_codigo, ਍                      搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀攀琀愀瀀愀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀攀琀愀瀀愀开挀漀搀椀最漀㈀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀挀氀愀猀椀昀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀搀漀挀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀搀漀挀开挀漀搀椀最漀㈀Ⰰ ഀഀ
                      dbo.ao_solicitud.doc_numero, dbo.ao_solicitud.doc_numero2, dbo.ao_solicitud.poa_codigo, dbo.ao_solicitud.ges_gestion_ant, dbo.ao_solicitud.unidad_codigo_ant, ਍                      搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀猀漀氀椀挀椀琀甀搀开挀漀搀椀最漀开愀渀琀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀挀漀爀爀攀氀开搀攀琀愀氀氀攀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀挀漀爀爀攀氀开攀搀椀昀椀挀愀挀椀漀渀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀挀漀爀爀攀氀开挀愀氀挀甀氀漀Ⰰ ഀഀ
                      dbo.ao_solicitud.correl_persona, dbo.ao_solicitud.correl_cotiza, dbo.ao_solicitud.correl_bitacora, dbo.ao_solicitud.archivo_respaldo, ਍                      搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀愀爀挀栀椀瘀漀开爀攀猀瀀愀氀搀漀开挀愀爀最愀搀漀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀攀猀琀愀搀漀开挀漀搀椀最漀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀攀猀琀愀搀漀开攀琀愀瀀愀㈀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀攀猀琀愀搀漀开挀漀琀椀稀愀Ⰰ ഀഀ
                      dbo.ao_solicitud.fecha_registro, dbo.ao_solicitud.hora_registro, dbo.ao_solicitud.usr_codigo, dbo.ao_solicitud.usr_codigo_aprueba, dbo.ao_solicitud.fecha_aprueba, ਍                      搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀栀漀爀愀开愀瀀爀甀攀戀愀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀昀攀挀栀愀开爀攀最椀猀琀爀漀㈀Ⰰ 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀甀猀爀开挀漀搀椀最漀㈀Ⰰ 搀戀漀⸀最挀开瀀爀漀瘀椀渀挀椀愀⸀瀀爀漀瘀开搀攀猀挀爀椀瀀挀椀漀渀Ⰰ ഀഀ
                      dbo.gc_zonas.zona_denominacion, dbo.gc_beneficiario.beneficiario_denominacion, dbo.gc_calles.calle_denominacion, dbo.gc_municipio.munic_descripcion, ਍                      搀戀漀⸀最挀开搀攀瀀愀爀琀愀洀攀渀琀漀⸀搀攀瀀琀漀开搀攀猀挀爀椀瀀挀椀漀渀Ⰰ 搀戀漀⸀最挀开甀渀椀搀愀搀开攀樀攀挀甀琀漀爀愀⸀甀渀椀搀愀搀开搀攀猀挀爀椀瀀挀椀漀渀Ⰰ 搀戀漀⸀最挀开琀椀瀀漀开猀漀氀椀挀椀琀甀搀⸀猀漀氀椀挀椀琀甀搀开琀椀瀀漀开搀攀猀挀爀椀瀀挀椀漀渀Ⰰ ഀഀ
                      gc_beneficiario_1.beneficiario_denominacion AS beneficiario_denominacion_edif, dbo.gc_calles.calle_tipo, dbo.gc_edificaciones.edif_descripcion, ਍                      搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀攀搀椀昀开琀椀瀀漀ഀഀ
FROM         dbo.ao_solicitud INNER JOIN਍                      搀戀漀⸀最挀开甀渀椀搀愀搀开攀樀攀挀甀琀漀爀愀 伀一 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀甀渀椀搀愀搀开挀漀搀椀最漀 㴀 搀戀漀⸀最挀开甀渀椀搀愀搀开攀樀攀挀甀琀漀爀愀⸀甀渀椀搀愀搀开挀漀搀椀最漀 䰀䔀䘀吀 伀唀吀䔀刀 䨀伀䤀一ഀഀ
                      dbo.gc_tipo_solicitud ON dbo.ao_solicitud.solicitud_tipo = dbo.gc_tipo_solicitud.solicitud_tipo LEFT OUTER JOIN਍                      搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀 伀一 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀攀搀椀昀开挀漀搀椀最漀 㴀 搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀攀搀椀昀开挀漀搀椀最漀 䰀䔀䘀吀 伀唀吀䔀刀 䨀伀䤀一ഀഀ
                      dbo.gc_departamento ON dbo.gc_edificaciones.depto_codigo = dbo.gc_departamento.depto_codigo LEFT OUTER JOIN਍                      搀戀漀⸀最挀开戀攀渀攀昀椀挀椀愀爀椀漀 伀一 搀戀漀⸀愀漀开猀漀氀椀挀椀琀甀搀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀 㴀 搀戀漀⸀最挀开戀攀渀攀昀椀挀椀愀爀椀漀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀 䰀䔀䘀吀 伀唀吀䔀刀 䨀伀䤀一ഀഀ
                      dbo.gc_municipio ON dbo.gc_edificaciones.munic_codigo = dbo.gc_municipio.munic_codigo LEFT OUTER JOIN਍                      搀戀漀⸀最挀开瀀爀漀瘀椀渀挀椀愀 伀一 搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀瀀爀漀瘀开挀漀搀椀最漀 㴀 搀戀漀⸀最挀开瀀爀漀瘀椀渀挀椀愀⸀瀀爀漀瘀开挀漀搀椀最漀 䰀䔀䘀吀 伀唀吀䔀刀 䨀伀䤀一ഀഀ
                      dbo.gc_calles ON dbo.gc_edificaciones.zona_codigo = dbo.gc_calles.zona_codigo AND ਍                      搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀挀愀氀氀攀开挀漀搀椀最漀 㴀 搀戀漀⸀最挀开挀愀氀氀攀猀⸀挀愀氀氀攀开挀漀搀椀最漀 䰀䔀䘀吀 伀唀吀䔀刀 䨀伀䤀一ഀഀ
                      dbo.gc_zonas ON dbo.gc_edificaciones.zona_codigo = dbo.gc_zonas.zona_codigo LEFT OUTER JOIN਍                      搀戀漀⸀最挀开戀攀渀攀昀椀挀椀愀爀椀漀 䄀匀 最挀开戀攀渀攀昀椀挀椀愀爀椀漀开㄀ 伀一 搀戀漀⸀最挀开攀搀椀昀椀挀愀挀椀漀渀攀猀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀 㴀 最挀开戀攀渀攀昀椀挀椀愀爀椀漀开㄀⸀戀攀渀攀昀椀挀椀愀爀椀漀开挀漀搀椀最漀ഀഀ
WHERE     (dbo.ao_solicitud.unidad_codigo = 'DVTA') and  (dbo.ao_solicitud.solicitud_fecha_solicitud BETWEEN CONVERT(DATETIME, @FInicio, 103) AND CONVERT(DATETIME, @FFinal, 103)) 