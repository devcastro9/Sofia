/*
   miércoles, 07 de junio de 2000 11:16:00
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
CREATE TABLE dbo.fc_DatosGTZ
	(
 	Nro_Cmpte varchar(10) NULL,
	Organismo varchar(150) NULL,
	Fecha_pago smalldatetime NULL,
	Monto float(53) NULL,
	Cambio float(53) NULL,
	Beneficiario varchar(60) NULL,
	Justificacion varchar(200) NULL,
	Cta_Codigo varchar(50) NULL,
	Nro_Cheque varchar(10) NULL,
	Banco varchar(100) NULL,
	Transf_Cheq varchar(15) NULL,
	Literal varchar(255) NULL,
	estado_conciliacion varchar(1) NULL
	) ON [PRIMARY]
GO
COMMIT
