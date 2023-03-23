/*
   miércoles, 07 de junio de 2000 11:54:35
   Usuario: sa
   Servidor: sersis
   Base de datos: SAF2000
   Aplicación: 
*/

BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO
COMMIT
BEGIN TRANSACTION
ALTER TABLE dbo.fo_ingresos ADD
	estado_conciliacion varchar(1) NULL
GO
COMMIT
