/*
   miércoles, 07 de junio de 2000 11:34:16
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
ALTER TABLE dbo.fc_DatosGTZ ADD
	status varchar(1) NULL
GO
COMMIT
