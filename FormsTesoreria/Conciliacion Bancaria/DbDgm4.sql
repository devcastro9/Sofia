/*
   miércoles, 07 de junio de 2000 11:26:55
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
EXECUTE sp_rename 'dbo.fc_DatosGTZ.Nro_Cheque', 'Tmp_Nro_Doc_1', 'COLUMN'
GO
EXECUTE sp_rename 'dbo.fc_DatosGTZ.Tmp_Nro_Doc_1', 'Nro_Doc', 'COLUMN'
GO
COMMIT
