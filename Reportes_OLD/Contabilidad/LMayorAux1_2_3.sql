USE [ADMIN_EMPRESA] ਍䜀伀ഀഀ
/****** Object:  StoredProcedure [dbo].[LMayorAux1_2_3]    Script Date: 04/10/2017 18:33:03 ******/਍匀䔀吀 䄀一匀䤀开一唀䰀䰀匀 伀䘀䘀ഀഀ
GO਍匀䔀吀 儀唀伀吀䔀䐀开䤀䐀䔀一吀䤀䘀䤀䔀刀 伀䘀䘀ഀഀ
GO਍䄀䰀吀䔀刀 瀀爀漀挀攀搀甀爀攀  嬀搀戀漀崀⸀嬀䰀䴀愀礀漀爀䄀甀砀㄀开㈀开㌀崀ഀഀ
(਍ऀ 䀀䘀䘀䤀渀椀挀椀漀 瘀愀爀挀栀愀爀⠀㄀　⤀Ⰰഀഀ
	 @FFFinal varchar(10) ,਍ऀ 䀀挀甀攀渀琀愀  瘀愀爀挀栀愀爀  ⠀㔀⤀ Ⰰഀഀ
	 @subcta1 varchar (3) ,਍ऀ 䀀猀甀戀挀琀愀㈀ 瘀愀爀挀栀愀爀 ⠀㌀⤀ Ⰰഀഀ
	 @busca1 varchar(40),਍ऀ 䀀戀甀猀挀愀㈀ 瘀愀爀挀栀愀爀⠀㐀　⤀Ⰰഀഀ
  	 @busca3 varchar(40),਍ऀ 䀀愀甀砀㄀ 瘀愀爀挀栀愀爀⠀㌀⤀Ⰰഀഀ
	 @aux2 varchar(3),਍ऀ 䀀愀甀砀㌀ 瘀愀爀挀栀愀爀⠀㌀⤀ഀഀ
)਍⼀⨀ഀഀ
declare	 @FFInicio varchar(10)਍搀攀挀氀愀爀攀ऀ 䀀䘀䘀䘀椀渀愀氀 瘀愀爀挀栀愀爀⠀㄀　⤀ ഀഀ
declare	 @cuenta  varchar  (5)਍搀攀挀氀愀爀攀ऀ 䀀猀甀戀挀琀愀㄀ 瘀愀爀挀栀愀爀 ⠀㌀⤀ ഀഀ
declare	 @subcta2 varchar (3) ਍搀攀挀氀愀爀攀ऀ 䀀戀甀猀挀愀㄀ 瘀愀爀挀栀愀爀⠀㄀㔀⤀ഀഀ
declare	 @busca2 varchar(15)਍搀攀挀氀愀爀攀ऀ 䀀戀甀猀挀愀㌀ 瘀愀爀挀栀愀爀⠀㄀㔀⤀ഀഀ
declare	 @aux1 varchar(3)਍搀攀挀氀愀爀攀ऀ 䀀愀甀砀㈀ 瘀愀爀挀栀愀爀⠀㌀⤀ഀഀ
declare	 @aux3 varchar(3)਍ഀഀ
set	 @FFInicio ='01/01/2016'਍猀攀琀ऀ 䀀䘀䘀䘀椀渀愀氀 㴀✀㌀㄀⼀㄀㈀⼀㈀　㄀㘀✀ഀഀ
set	 @cuenta  ='1121'਍猀攀琀ऀ 䀀猀甀戀挀琀愀㄀ 㴀✀　㈀✀ ഀഀ
set	 @subcta2 ='00' ਍猀攀琀ऀ 䀀戀甀猀挀愀㄀㴀✀㤀㤀㤀㤀㔀　✀ഀഀ
set	 @busca2 ='415'਍猀攀琀ऀ 䀀戀甀猀挀愀㈀ 㴀✀㐀㄀㔀✀ഀഀ
set	 @aux1 ='01'਍猀攀琀ऀ 䀀愀甀砀㈀ 㴀✀　㠀✀ഀഀ
set	 @aux3 ='00' ਍⨀⼀ഀഀ
---਍䄀匀ഀഀ
declare @SICtaBs money਍搀攀挀氀愀爀攀 䀀匀䤀䌀琀愀匀甀猀 洀漀渀攀礀ഀഀ
DECLARE @SIBs Money਍䐀䔀䌀䰀䄀刀䔀 䀀匀䤀匀甀猀 䴀漀渀攀礀ഀഀ
਍䐀䔀䌀䰀䄀刀䔀 䀀䘀䤀渀椀挀椀漀 䐀愀琀攀吀椀洀攀ഀഀ
SET @FInicio = CONVERT(DateTime, @FFInicio, 103)਍䐀䔀䌀䰀䄀刀䔀 䀀䘀䘀椀渀愀氀 䐀愀琀攀吀椀洀攀ഀഀ
SET @FFinal = CONVERT(DateTime, @FFFinal, 103)਍ഀഀ
if @busca1='' set @busca1= '%'਍椀昀 䀀戀甀猀挀愀㈀ 㴀✀✀ 猀攀琀 䀀戀甀猀挀愀㈀㴀 ✀─✀ഀഀ
if @busca3 ='' set @busca3= '%'਍ഀഀ
--਍⼀⨀⨀⨀⨀挀爀攀愀洀漀猀  琀愀戀氀愀 愀甀砀椀氀椀愀爀 瀀愀爀愀 攀氀 洀愀礀漀爀⨀⨀⨀⼀ഀഀ
create table #LMayorAux1_2_3਍⠀ഀഀ
	IDCta Int IDENTITY(1,1), ਍ऀ昀攀挀栀愀 搀愀琀攀琀椀洀攀Ⰰഀഀ
	TC money DEFAULT 0,਍ऀⴀⴀ挀漀洀瀀 瘀愀爀挀栀愀爀⠀㘀⤀Ⰰഀഀ
	comp INT,਍ऀ琀椀瀀漀 瘀愀爀挀栀愀爀⠀㌀⤀Ⰰഀഀ
	--cte varchar(6),਍ऀ挀琀攀 䤀一吀Ⰰഀഀ
	org varchar(15),਍ऀ最氀漀猀愀 瘀愀爀挀栀愀爀 ⠀㌀㔀㔀⤀Ⰰഀഀ
	debe money DEFAULT 0 ,਍ऀ栀愀戀攀爀 洀漀渀攀礀 䐀䔀䘀䄀唀䰀吀 　Ⰰഀഀ
	MovSus money DEFAULT 0,਍ऀ匀愀氀搀漀䈀猀 洀漀渀攀礀 䐀䔀䘀䄀唀䰀吀 　Ⰰഀഀ
	SaldoSus money DEFAULT 0,਍ऀ匀䤀䈀猀 洀漀渀攀礀 䐀䔀䘀䄀唀䰀吀 　Ⰰഀഀ
	SISus money DEFAULT 0਍ഀഀ
)਍ഀഀ
/*****movimientos de la cuenta en el debe****/਍䤀一匀䔀刀吀 䤀一吀伀 ⌀䰀䴀愀礀漀爀䄀甀砀㄀开㈀开㌀ ⠀昀攀挀栀愀Ⰰ吀䌀Ⰰ䌀漀洀瀀Ⰰ琀椀瀀漀Ⰰ挀琀攀Ⰰ漀爀最Ⰰ最氀漀猀愀Ⰰ搀攀戀攀Ⰰ䴀漀瘀匀甀猀⤀ഀഀ
(਍ऀ匀䔀䰀䔀䌀吀 䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀䘀攀挀栀愀开琀爀愀渀猀愀挀椀漀渀Ⰰ 椀猀渀甀氀氀⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䌀愀洀戀椀漀Ⰰ　⤀ 䄀匀 䐀开䌀愀洀戀椀漀Ⰰഀഀ
		 Co_Comprobante_M.Cod_Comp, Co_Comprobante_M.Tipo_Comp,  ਍ऀऀ 椀猀渀甀氀氀⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀猀漀氀椀挀椀琀甀搀开挀漀搀椀最漀Ⰰ✀　✀⤀ 䄀匀 䌀漀搀开琀爀愀渀猀Ⰰ 椀猀渀甀氀氀⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀甀渀椀搀愀搀开挀漀搀椀最漀Ⰰ✀一一✀⤀ 䄀匀 漀爀最开挀漀搀椀最漀Ⰰഀഀ
	              Co_Comprobante_M.Glosa, isnull(CO_Diario.D_MontoBs,0) AS D_MontoBs, ਍ऀऀ椀猀渀甀氀氀⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䴀漀渀琀漀䐀氀Ⰰ　⤀ 䄀匀 䐀开䴀漀渀琀漀䐀氀ഀഀ
	FROM Co_Comprobante_M INNER JOIN਍ऀ           䌀伀开䐀椀愀爀椀漀 伀一 ഀഀ
	           Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp਍ऀ圀䠀䔀刀䔀 ⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀攀猀琀愀搀漀开挀漀搀椀最漀 㴀 ✀䄀倀刀✀⤀ 䄀一䐀 ഀഀ
	          (CO_Diario.D_Cuenta = @cuenta) AND ਍ऀ          ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开匀甀戀挀琀愀㄀ 㴀 䀀猀甀戀挀琀愀㄀⤀ 䄀一䐀 ഀഀ
	          (CO_Diario.D_SubCta2 = @subcta2) AND਍ऀऀ  ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䄀甀砀㄀㴀 䀀愀甀砀㄀⤀ 䄀一䐀ഀഀ
		  (CO_Diario.D_Aux2= @aux2) AND਍ऀ ऀ  ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䌀琀愀开䄀甀砀㄀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㄀⤀䄀一䐀ഀഀ
	 	  (CO_Diario.D_Cta_Aux2  LIKE @busca2)AND਍ऀऀ  ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䌀琀愀开䄀甀砀㌀  䰀䤀䬀䔀 䀀戀甀猀挀愀㌀⤀ഀഀ
		  AND (Co_Comprobante_M.Fecha_transacion BETWEEN ਍ऀ         䌀伀一嘀䔀刀吀⠀䐀䄀吀䔀吀䤀䴀䔀Ⰰ 䀀䘀䤀渀椀挀椀漀Ⰰ ㄀　㌀⤀ 䄀一䐀 ഀഀ
	        CONVERT(DATETIME, @FFinal, 103)) ਍⤀ഀഀ
/********Movimiento de la cuenta en el Haber****/਍䤀一匀䔀刀吀 䤀一吀伀 ⌀䰀䴀愀礀漀爀䄀甀砀㄀开㈀开㌀ ⠀昀攀挀栀愀Ⰰ吀䌀Ⰰ䌀漀洀瀀Ⰰ琀椀瀀漀Ⰰ挀琀攀Ⰰ漀爀最Ⰰ最氀漀猀愀Ⰰ䠀愀戀攀爀Ⰰ䴀漀瘀匀甀猀⤀ഀഀ
(਍ऀ匀䔀䰀䔀䌀吀 䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀䘀攀挀栀愀开琀爀愀渀猀愀挀椀漀渀Ⰰ椀猀渀甀氀氀⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䌀愀洀戀椀漀Ⰰ　⤀ 䄀匀 䠀开䌀愀洀戀椀漀Ⰰഀഀ
		Co_Comprobante_M.Cod_Comp, Co_Comprobante_M.Tipo_Comp,  ਍ऀऀ椀猀渀甀氀氀⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀猀漀氀椀挀椀琀甀搀开挀漀搀椀最漀Ⰰ✀　✀⤀ 䄀匀 䌀漀搀开琀爀愀渀猀Ⰰ 椀猀渀甀氀氀⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀甀渀椀搀愀搀开挀漀搀椀最漀Ⰰ✀一一✀⤀ 䄀匀 漀爀最开挀漀搀椀最漀Ⰰഀഀ
	             Co_Comprobante_M.Glosa, isnull(CO_Diario.H_MontoBs,0) AS H_MontoBs, ਍ऀ             椀猀渀甀氀氀⠀⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䴀漀渀琀漀䐀氀 ⨀ ⴀ㄀⤀Ⰰ　⤀ 䄀匀 䠀开䴀漀渀琀漀䐀氀ഀഀ
	FROM Co_Comprobante_M INNER JOIN਍ऀ             䌀伀开䐀椀愀爀椀漀 伀一 ഀഀ
	             Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp਍ऀ圀䠀䔀刀䔀 ⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀攀猀琀愀搀漀开挀漀搀椀最漀 㴀 ✀䄀倀刀✀⤀ 䄀一䐀 ഀഀ
	    	   (CO_Diario.H_Cuenta = @cuenta) AND ਍ऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开匀甀戀挀琀愀㄀ 㴀 䀀猀甀戀挀琀愀㄀⤀ 䄀一䐀 ഀഀ
		    (CO_Diario.H_SubCta2 = @subcta2) AND ਍ऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䄀甀砀㄀ 㴀 䀀愀甀砀㄀⤀ 䄀一䐀ഀഀ
		    (CO_Diario.H_Aux2= @aux2) AND਍ऀ ऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䌀琀愀开䄀甀砀㄀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㄀⤀䄀一䐀ഀഀ
	 	    (CO_Diario.H_Cta_Aux2  LIKE @busca2)AND਍ऀ ऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䌀琀愀开䄀甀砀㌀  䰀䤀䬀䔀 䀀戀甀猀挀愀㌀⤀ 䄀一䐀ഀഀ
		    (Co_Comprobante_M.Fecha_transacion BETWEEN ਍ऀऀ    䌀伀一嘀䔀刀吀⠀䐀䄀吀䔀吀䤀䴀䔀Ⰰ 䀀䘀䤀渀椀挀椀漀Ⰰ ㄀　㌀⤀ 䄀一䐀 ഀഀ
		    CONVERT(DATETIME, @FFinal, 103)) ਍ഀഀ
)਍⼀⨀⨀⨀⨀⨀⨀⼀ഀഀ
/*****tabla de saldos */਍ഀഀ
/**balance de apertura*/਍猀攀琀 䀀匀䤀䌀琀愀䈀猀 㴀 ⠀匀䔀䰀䔀䌀吀 匀唀䴀⠀椀猀渀甀氀氀⠀䐀攀戀攀匀愀氀搀漀䤀䈀猀Ⰰ　⤀⤀ⴀ匀唀䴀⠀椀猀渀甀氀氀⠀䠀愀戀攀爀匀愀氀搀漀䤀䈀猀Ⰰ　⤀⤀ഀഀ
				    	     	FROM fo_balance_apertura ਍                 ऀऀऀऀ圀䠀䔀刀䔀 挀甀攀渀琀愀 㴀 䀀挀甀攀渀琀愀ഀഀ
						      AND subcta1 = @subcta1਍ऀऀऀऀऀऀ      䄀一䐀 猀甀戀挀琀愀㈀ 㴀 䀀猀甀戀挀琀愀㈀ ഀഀ
						      AND denominacion_aux1 LIKE @busca1 ਍ऀऀऀऀऀऀ      䄀一䐀 搀攀渀漀洀椀渀愀挀椀漀渀开愀甀砀㈀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㈀ഀഀ
						      AND denominacion_aux3 LIKE @busca3)਍猀攀琀 䀀匀䤀䌀琀愀匀甀猀 㴀 ⠀匀䔀䰀䔀䌀吀 匀唀䴀⠀椀猀渀甀氀氀⠀䐀攀戀攀匀愀氀搀漀䤀匀甀猀Ⰰ　⤀⤀ⴀ匀唀䴀⠀椀猀渀甀氀氀⠀䠀愀戀攀爀匀愀氀搀漀䤀匀甀猀Ⰰ　⤀⤀ഀഀ
				    	     	FROM fo_balance_apertura ਍                 ऀऀऀऀ圀䠀䔀刀䔀 挀甀攀渀琀愀 㴀 䀀挀甀攀渀琀愀ഀഀ
						      AND subcta1 = @subcta1਍ऀऀऀऀऀऀ      䄀一䐀 猀甀戀挀琀愀㈀ 㴀 䀀猀甀戀挀琀愀㈀ ഀഀ
						      AND denominacion_aux1 LIKE @busca1 ਍ऀऀऀऀऀऀ      䄀一䐀 搀攀渀漀洀椀渀愀挀椀漀渀开愀甀砀㈀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㈀ഀഀ
						      AND denominacion_aux3 LIKE @busca3)਍ⴀⴀഀഀ
set @SIBs = isnull(@SICtaBs,0)+ ISNULL((SELECT SUM(isnull(CO_Diario.D_MontoBs,0)) ਍ऀऀ䘀刀伀䴀 䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀 䤀一一䔀刀 䨀伀䤀一ഀഀ
		CO_Diario ON ਍ऀऀ䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀䌀漀搀开䌀漀洀瀀 㴀 䌀伀开䐀椀愀爀椀漀⸀䌀漀搀开䌀漀洀瀀ഀഀ
		WHERE  (Co_Comprobante_M.estado_codigo = 'APR') and਍ऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䌀甀攀渀琀愀 㴀 䀀挀甀攀渀琀愀⤀ 䄀一䐀 ഀഀ
		    (CO_Diario.D_Subcta1 = @subcta1) AND ਍ऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开匀甀戀䌀琀愀㈀ 㴀 䀀猀甀戀挀琀愀㈀⤀ 䄀一䐀 ഀഀ
		    (CO_Diario.D_Aux1 = @aux1) AND਍ऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䄀甀砀㈀㴀 䀀愀甀砀㈀⤀ 䄀一䐀ഀഀ
	 	    (CO_Diario.D_Cta_Aux1 LIKE @busca1)AND਍ऀ ऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䌀琀愀开䄀甀砀㈀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㈀⤀䄀一䐀ഀഀ
	 	    (CO_Diario.D_Cta_Aux3 LIKE @busca3) AND਍ऀऀ    ⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀䘀攀挀栀愀开琀爀愀渀猀愀挀椀漀渀 㰀 䌀伀一嘀䔀刀吀⠀䐀䄀吀䔀吀䤀䴀䔀Ⰰ ഀഀ
		     @FInicio, 102))),0) - ਍ऀ       䤀匀一唀䰀䰀⠀⠀匀䔀䰀䔀䌀吀 匀唀䴀⠀椀猀渀甀氀氀⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䴀漀渀琀漀䈀猀Ⰰ　⤀⤀ ഀഀ
			FROM Co_Comprobante_M INNER JOIN਍ऀऀऀ    䌀伀开䐀椀愀爀椀漀 伀一 ഀഀ
			    Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp਍ऀऀऀ圀䠀䔀刀䔀  ⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀攀猀琀愀搀漀开挀漀搀椀最漀 㴀 ✀䄀倀刀✀⤀ 䄀一䐀ഀഀ
			    (CO_Diario.H_Cuenta = @cuenta) AND ਍ऀऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开匀甀戀挀琀愀㄀ 㴀 䀀猀甀戀挀琀愀㄀⤀ 䄀一䐀 ഀഀ
			    (CO_Diario.H_SubCta2 = @subcta2) AND ਍ऀऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䄀甀砀㄀ 㴀 䀀愀甀砀㄀⤀ 䄀一䐀ഀഀ
			    (CO_Diario.H_Aux2= @aux2) AND਍ऀऀ ऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䌀琀愀开䄀甀砀㄀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㄀⤀䄀一䐀ഀഀ
		 	    (CO_Diario.H_Cta_Aux2 LIKE @busca2)AND਍ऀऀ ऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䌀琀愀开䄀甀砀㌀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㌀⤀ 䄀一䐀ഀഀ
			    (Co_Comprobante_M.Fecha_transacion < CONVERT(DATETIME, ਍ऀऀऀ    䀀䘀䤀渀椀挀椀漀Ⰰ ㄀　㈀⤀⤀⤀Ⰰ　⤀ഀഀ
਍猀攀琀 䀀匀䤀匀甀猀 㴀椀猀渀甀氀氀⠀䀀匀䤀䌀琀愀匀甀猀Ⰰ　⤀⬀ 䤀匀一唀䰀䰀⠀⠀匀䔀䰀䔀䌀吀 匀唀䴀⠀椀猀渀甀氀氀⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䴀漀渀琀漀䐀氀Ⰰ　⤀⤀ ഀഀ
		FROM Co_Comprobante_M INNER JOIN਍ऀऀ䌀伀开䐀椀愀爀椀漀 伀一 ഀഀ
		Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp਍ऀऀ圀䠀䔀刀䔀  ⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀攀猀琀愀搀漀开挀漀搀椀最漀 㴀 ✀䄀倀刀✀⤀  䄀一䐀ഀഀ
		    (CO_Diario.D_Cuenta = @cuenta) AND ਍ऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开匀甀戀挀琀愀㄀ 㴀 䀀猀甀戀挀琀愀㄀⤀ 䄀一䐀 ഀഀ
		    (CO_Diario.D_SubCta2 = @subcta2) AND ਍ऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䄀甀砀㄀ 㴀 䀀愀甀砀㄀⤀ 䄀一䐀ഀഀ
		    (CO_Diario.D_Aux2= @aux2) AND਍ऀ ऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䌀琀愀开䄀甀砀㄀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㄀⤀䄀一䐀ഀഀ
	 	    (CO_Diario.D_Cta_Aux2 LIKE @busca2)AND਍ऀ ऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䐀开䌀琀愀开䄀甀砀㌀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㌀⤀ 䄀一䐀ഀഀ
		    (Co_Comprobante_M.Fecha_transacion < CONVERT(DATETIME, ਍ऀऀ    䀀䘀䤀渀椀挀椀漀Ⰰ ㄀　㈀⤀⤀⤀Ⰰ　⤀ ⴀ ഀഀ
		 ISNULL((SELECT SUM(isnull(CO_Diario.H_MontoDl,0)) ਍ऀऀऀ䘀刀伀䴀 䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀 䤀一一䔀刀 䨀伀䤀一ഀഀ
			    CO_Diario ON ਍ऀऀऀ    䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀䌀漀搀开䌀漀洀瀀 㴀 䌀伀开䐀椀愀爀椀漀⸀䌀漀搀开䌀漀洀瀀ഀഀ
			WHERE  (Co_Comprobante_M.estado_codigo = 'APR') AND਍ऀऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䌀甀攀渀琀愀 㴀 䀀挀甀攀渀琀愀⤀ 䄀一䐀 ഀഀ
			    (CO_Diario.H_Subcta1 = @subcta1) AND ਍ऀऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开匀甀戀䌀琀愀㈀ 㴀 䀀猀甀戀挀琀愀㈀⤀ 䄀一䐀 ഀഀ
			    (CO_Diario.H_Aux1 = @aux1) AND਍ऀऀऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䄀甀砀㈀㴀 䀀愀甀砀㈀⤀ 䄀一䐀ഀഀ
		 	    (CO_Diario.H_Cta_Aux1 LIKE @busca1)AND਍ऀऀ ऀ    ⠀䌀伀开䐀椀愀爀椀漀⸀䠀开䌀琀愀开䄀甀砀㈀ 䰀䤀䬀䔀 䀀戀甀猀挀愀㈀⤀䄀一䐀ഀഀ
		 	    (CO_Diario.H_Cta_Aux3 LIKE @busca3) AND਍ऀऀऀ    ⠀䌀漀开䌀漀洀瀀爀漀戀愀渀琀攀开䴀⸀䘀攀挀栀愀开琀爀愀渀猀愀挀椀漀渀 㰀 䌀伀一嘀䔀刀吀⠀䐀䄀吀䔀吀䤀䴀䔀Ⰰ ഀഀ
			    @FInicio, 102))),0)਍ഀഀ
create table #Saldos1_2_3਍⠀  猀愀氀搀漀戀猀 洀漀渀攀礀Ⰰഀഀ
   saldosus money਍⤀ഀഀ
਍䤀一匀䔀刀吀 䤀一吀伀 ⌀匀愀氀搀漀猀㄀开㈀开㌀⠀匀愀氀搀漀䈀猀Ⰰ 匀愀氀搀漀匀甀猀⤀ 匀䔀䰀䔀䌀吀 䀀匀䤀䈀猀Ⰰ 䀀匀䤀匀甀猀ഀഀ
਍唀倀䐀䄀吀䔀 ⌀䰀䴀愀礀漀爀䄀甀砀㄀开㈀开㌀ 猀攀琀ഀഀ
	SIBs =@SIBs,਍ऀ匀䤀匀甀猀㴀䀀匀䤀匀甀猀ഀഀ
਍䐀䔀䌀䰀䄀刀䔀 䀀䤀䐀䌀琀愀 䤀渀琀ഀഀ
਍䐀䔀䌀䰀䄀刀䔀 焀䰀䴀愀礀漀爀㠀 匀䌀刀伀䰀䰀 䌀唀刀匀伀刀ഀഀ
	FOR SELECT IDCta FROM #LMayorAux1_2_3 ORDER BY Fecha਍伀倀䔀一 焀䰀䴀愀礀漀爀㠀ഀഀ
FETCH FIRST FROM qLMayor8 INTO @IDCta਍圀䠀䤀䰀䔀 䀀䀀䘀䔀吀䌀䠀开匀吀䄀吀唀匀  㴀 　ഀഀ
 BEGIN਍ऀ匀䔀吀 䀀匀䤀䈀猀 㴀 䀀匀䤀䈀猀 ⬀ ⠀匀䔀䰀䔀䌀吀 䐀攀戀攀 䘀刀伀䴀 ⌀䰀䴀愀礀漀爀䄀甀砀㄀开㈀开㌀ 圀䠀䔀刀䔀 䤀䐀䌀琀愀 㴀 䀀䤀䐀䌀琀愀⤀ⴀ ⠀匀䔀䰀䔀䌀吀 䠀愀戀攀爀 䘀刀伀䴀 ⌀䰀䴀愀礀漀爀䄀甀砀㄀开㈀开㌀ 圀䠀䔀刀䔀 䤀䐀䌀琀愀 㴀 䀀䤀䐀䌀琀愀⤀ ഀഀ
	SET @SISus = @SISus + (SELECT MovSus FROM #LMayorAux1_2_3 WHERE IDCta = @IDCta) ਍ऀ唀倀䐀䄀吀䔀 ⌀䰀䴀愀礀漀爀䄀甀砀㄀开㈀开㌀ 匀䔀吀 匀愀氀搀漀䈀猀 㴀 䀀匀䤀䈀猀 圀䠀䔀刀䔀 䤀䐀䌀琀愀 㴀 䀀䤀䐀䌀琀愀ഀഀ
	UPDATE #LMayorAux1_2_3 SET SaldoSus = @SISus WHERE IDCta = @IDCta਍ऀ䘀䔀吀䌀䠀 一䔀堀吀 䘀刀伀䴀 焀䰀䴀愀礀漀爀㠀 䤀一吀伀 䀀䤀䐀䌀琀愀ഀഀ
 END਍䌀䰀伀匀䔀 焀䰀䴀愀礀漀爀㠀ഀഀ
DEALLOCATE qLMayor8਍ഀഀ
select * from #LMayorAux1_2_3 order by Fecha,comp 