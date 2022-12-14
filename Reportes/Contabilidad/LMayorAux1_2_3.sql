USE [ADMIN_EMPRESA] 
GO
/****** Object:  StoredProcedure [dbo].[LMayorAux1_2_3]    Script Date: 04/10/2017 18:33:03 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
ALTER procedure  [dbo].[LMayorAux1_2_3]
(
	 @FFInicio varchar(10),
	 @FFFinal varchar(10) ,
	 @cuenta  varchar  (5) ,
	 @subcta1 varchar (3) ,
	 @subcta2 varchar (3) ,
	 @busca1 varchar(40),
	 @busca2 varchar(40),
  	 @busca3 varchar(40),
	 @aux1 varchar(3),
	 @aux2 varchar(3),
	 @aux3 varchar(3)
)
/*
declare	 @FFInicio varchar(10)
declare	 @FFFinal varchar(10) 
declare	 @cuenta  varchar  (5)
declare	 @subcta1 varchar (3) 
declare	 @subcta2 varchar (3) 
declare	 @busca1 varchar(15)
declare	 @busca2 varchar(15)
declare	 @busca3 varchar(15)
declare	 @aux1 varchar(3)
declare	 @aux2 varchar(3)
declare	 @aux3 varchar(3)

set	 @FFInicio ='01/01/2016'
set	 @FFFinal ='31/12/2016'
set	 @cuenta  ='1121'
set	 @subcta1 ='02' 
set	 @subcta2 ='00' 
set	 @busca1='999950'
set	 @busca2 ='415'
set	 @busca2 ='415'
set	 @aux1 ='01'
set	 @aux2 ='08'
set	 @aux3 ='00' 
*/
---
AS
declare @SICtaBs money
declare @SICtaSus money
DECLARE @SIBs Money
DECLARE @SISus Money

DECLARE @FInicio DateTime
SET @FInicio = CONVERT(DateTime, @FFInicio, 103)
DECLARE @FFinal DateTime
SET @FFinal = CONVERT(DateTime, @FFFinal, 103)

if @busca1='' set @busca1= '%'
if @busca2 ='' set @busca2= '%'
if @busca3 ='' set @busca3= '%'

--
/****creamos  tabla auxiliar para el mayor***/
create table #LMayorAux1_2_3
(
	IDCta Int IDENTITY(1,1), 
	fecha datetime,
	TC money DEFAULT 0,
	--comp varchar(6),
	comp INT,
	tipo varchar(3),
	--cte varchar(6),
	cte INT,
	org varchar(15),
	glosa varchar (355),
	debe money DEFAULT 0 ,
	haber money DEFAULT 0,
	MovSus money DEFAULT 0,
	SaldoBs money DEFAULT 0,
	SaldoSus money DEFAULT 0,
	SIBs money DEFAULT 0,
	SISus money DEFAULT 0

)

/*****movimientos de la cuenta en el debe****/
INSERT INTO #LMayorAux1_2_3 (fecha,TC,Comp,tipo,cte,org,glosa,debe,MovSus)
(
	SELECT Co_Comprobante_M.Fecha_transacion, isnull(CO_Diario.D_Cambio,0) AS D_Cambio,
		 Co_Comprobante_M.Cod_Comp, Co_Comprobante_M.Tipo_Comp,  
		 isnull(Co_Comprobante_M.solicitud_codigo,'0') AS Cod_trans, isnull(Co_Comprobante_M.unidad_codigo,'NN') AS org_codigo,
	              Co_Comprobante_M.Glosa, isnull(CO_Diario.D_MontoBs,0) AS D_MontoBs, 
		isnull(CO_Diario.D_MontoDl,0) AS D_MontoDl
	FROM Co_Comprobante_M INNER JOIN
	           CO_Diario ON 
	           Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp
	WHERE (Co_Comprobante_M.estado_codigo = 'APR') AND 
	          (CO_Diario.D_Cuenta = @cuenta) AND 
	          (CO_Diario.D_Subcta1 = @subcta1) AND 
	          (CO_Diario.D_SubCta2 = @subcta2) AND
		  (CO_Diario.D_Aux1= @aux1) AND
		  (CO_Diario.D_Aux2= @aux2) AND
	 	  (CO_Diario.D_Cta_Aux1 LIKE @busca1)AND
	 	  (CO_Diario.D_Cta_Aux2  LIKE @busca2)AND
		  (CO_Diario.D_Cta_Aux3  LIKE @busca3)
		  AND (Co_Comprobante_M.Fecha_transacion BETWEEN 
	         CONVERT(DATETIME, @FInicio, 103) AND 
	        CONVERT(DATETIME, @FFinal, 103)) 
)
/********Movimiento de la cuenta en el Haber****/
INSERT INTO #LMayorAux1_2_3 (fecha,TC,Comp,tipo,cte,org,glosa,Haber,MovSus)
(
	SELECT Co_Comprobante_M.Fecha_transacion,isnull(CO_Diario.H_Cambio,0) AS H_Cambio,
		Co_Comprobante_M.Cod_Comp, Co_Comprobante_M.Tipo_Comp,  
		isnull(Co_Comprobante_M.solicitud_codigo,'0') AS Cod_trans, isnull(Co_Comprobante_M.unidad_codigo,'NN') AS org_codigo,
	             Co_Comprobante_M.Glosa, isnull(CO_Diario.H_MontoBs,0) AS H_MontoBs, 
	             isnull((CO_Diario.H_MontoDl * -1),0) AS H_MontoDl
	FROM Co_Comprobante_M INNER JOIN
	             CO_Diario ON 
	             Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp
	WHERE (Co_Comprobante_M.estado_codigo = 'APR') AND 
	    	   (CO_Diario.H_Cuenta = @cuenta) AND 
		    (CO_Diario.H_Subcta1 = @subcta1) AND 
		    (CO_Diario.H_SubCta2 = @subcta2) AND 
		    (CO_Diario.H_Aux1 = @aux1) AND
		    (CO_Diario.H_Aux2= @aux2) AND
	 	    (CO_Diario.H_Cta_Aux1 LIKE @busca1)AND
	 	    (CO_Diario.H_Cta_Aux2  LIKE @busca2)AND
	 	    (CO_Diario.H_Cta_Aux3  LIKE @busca3) AND
		    (Co_Comprobante_M.Fecha_transacion BETWEEN 
		    CONVERT(DATETIME, @FInicio, 103) AND 
		    CONVERT(DATETIME, @FFinal, 103)) 

)
/******/
/*****tabla de saldos */

/**balance de apertura*/
set @SICtaBs = (SELECT SUM(isnull(DebeSaldoIBs,0))-SUM(isnull(HaberSaldoIBs,0))
				    	     	FROM fo_balance_apertura 
                 				WHERE cuenta = @cuenta
						      AND subcta1 = @subcta1
						      AND subcta2 = @subcta2 
						      AND denominacion_aux1 LIKE @busca1 
						      AND denominacion_aux2 LIKE @busca2
						      AND denominacion_aux3 LIKE @busca3)
set @SICtaSus = (SELECT SUM(isnull(DebeSaldoISus,0))-SUM(isnull(HaberSaldoISus,0))
				    	     	FROM fo_balance_apertura 
                 				WHERE cuenta = @cuenta
						      AND subcta1 = @subcta1
						      AND subcta2 = @subcta2 
						      AND denominacion_aux1 LIKE @busca1 
						      AND denominacion_aux2 LIKE @busca2
						      AND denominacion_aux3 LIKE @busca3)
--
set @SIBs = isnull(@SICtaBs,0)+ ISNULL((SELECT SUM(isnull(CO_Diario.D_MontoBs,0)) 
		FROM Co_Comprobante_M INNER JOIN
		CO_Diario ON 
		Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp
		WHERE  (Co_Comprobante_M.estado_codigo = 'APR') and
		    (CO_Diario.D_Cuenta = @cuenta) AND 
		    (CO_Diario.D_Subcta1 = @subcta1) AND 
		    (CO_Diario.D_SubCta2 = @subcta2) AND 
		    (CO_Diario.D_Aux1 = @aux1) AND
		    (CO_Diario.D_Aux2= @aux2) AND
	 	    (CO_Diario.D_Cta_Aux1 LIKE @busca1)AND
	 	    (CO_Diario.D_Cta_Aux2 LIKE @busca2)AND
	 	    (CO_Diario.D_Cta_Aux3 LIKE @busca3) AND
		    (Co_Comprobante_M.Fecha_transacion < CONVERT(DATETIME, 
		     @FInicio, 102))),0) - 
	       ISNULL((SELECT SUM(isnull(CO_Diario.H_MontoBs,0)) 
			FROM Co_Comprobante_M INNER JOIN
			    CO_Diario ON 
			    Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp
			WHERE  (Co_Comprobante_M.estado_codigo = 'APR') AND
			    (CO_Diario.H_Cuenta = @cuenta) AND 
			    (CO_Diario.H_Subcta1 = @subcta1) AND 
			    (CO_Diario.H_SubCta2 = @subcta2) AND 
			    (CO_Diario.H_Aux1 = @aux1) AND
			    (CO_Diario.H_Aux2= @aux2) AND
		 	    (CO_Diario.H_Cta_Aux1 LIKE @busca1)AND
		 	    (CO_Diario.H_Cta_Aux2 LIKE @busca2)AND
		 	    (CO_Diario.H_Cta_Aux3 LIKE @busca3) AND
			    (Co_Comprobante_M.Fecha_transacion < CONVERT(DATETIME, 
			    @FInicio, 102))),0)

set @SISus =isnull(@SICtaSus,0)+ ISNULL((SELECT SUM(isnull(CO_Diario.D_MontoDl,0)) 
		FROM Co_Comprobante_M INNER JOIN
		CO_Diario ON 
		Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp
		WHERE  (Co_Comprobante_M.estado_codigo = 'APR')  AND
		    (CO_Diario.D_Cuenta = @cuenta) AND 
		    (CO_Diario.D_Subcta1 = @subcta1) AND 
		    (CO_Diario.D_SubCta2 = @subcta2) AND 
		    (CO_Diario.D_Aux1 = @aux1) AND
		    (CO_Diario.D_Aux2= @aux2) AND
	 	    (CO_Diario.D_Cta_Aux1 LIKE @busca1)AND
	 	    (CO_Diario.D_Cta_Aux2 LIKE @busca2)AND
	 	    (CO_Diario.D_Cta_Aux3 LIKE @busca3) AND
		    (Co_Comprobante_M.Fecha_transacion < CONVERT(DATETIME, 
		    @FInicio, 102))),0) - 
		 ISNULL((SELECT SUM(isnull(CO_Diario.H_MontoDl,0)) 
			FROM Co_Comprobante_M INNER JOIN
			    CO_Diario ON 
			    Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp
			WHERE  (Co_Comprobante_M.estado_codigo = 'APR') AND
			    (CO_Diario.H_Cuenta = @cuenta) AND 
			    (CO_Diario.H_Subcta1 = @subcta1) AND 
			    (CO_Diario.H_SubCta2 = @subcta2) AND 
			    (CO_Diario.H_Aux1 = @aux1) AND
			    (CO_Diario.H_Aux2= @aux2) AND
		 	    (CO_Diario.H_Cta_Aux1 LIKE @busca1)AND
		 	    (CO_Diario.H_Cta_Aux2 LIKE @busca2)AND
		 	    (CO_Diario.H_Cta_Aux3 LIKE @busca3) AND
			    (Co_Comprobante_M.Fecha_transacion < CONVERT(DATETIME, 
			    @FInicio, 102))),0)

create table #Saldos1_2_3
(  saldobs money,
   saldosus money
)

INSERT INTO #Saldos1_2_3(SaldoBs, SaldoSus) SELECT @SIBs, @SISus

UPDATE #LMayorAux1_2_3 set
	SIBs =@SIBs,
	SISus=@SISus

DECLARE @IDCta Int

DECLARE qLMayor8 SCROLL CURSOR
	FOR SELECT IDCta FROM #LMayorAux1_2_3 ORDER BY Fecha
OPEN qLMayor8
FETCH FIRST FROM qLMayor8 INTO @IDCta
WHILE @@FETCH_STATUS  = 0
 BEGIN
	SET @SIBs = @SIBs + (SELECT Debe FROM #LMayorAux1_2_3 WHERE IDCta = @IDCta)- (SELECT Haber FROM #LMayorAux1_2_3 WHERE IDCta = @IDCta) 
	SET @SISus = @SISus + (SELECT MovSus FROM #LMayorAux1_2_3 WHERE IDCta = @IDCta) 
	UPDATE #LMayorAux1_2_3 SET SaldoBs = @SIBs WHERE IDCta = @IDCta
	UPDATE #LMayorAux1_2_3 SET SaldoSus = @SISus WHERE IDCta = @IDCta
	FETCH NEXT FROM qLMayor8 INTO @IDCta
 END
CLOSE qLMayor8
DEALLOCATE qLMayor8

select * from #LMayorAux1_2_3 order by Fecha,comp 