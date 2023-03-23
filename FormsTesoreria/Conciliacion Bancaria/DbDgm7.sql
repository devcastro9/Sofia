/*
   miércoles, 07 de junio de 2000 14:54:49
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
CREATE TABLE dbo.Tmp_pago_detalle
	(
 	Ges_gestion varchar(4) NOT NULL,
	org_codigo varchar(3) NOT NULL,
	codigo_pago int NOT NULL,
	codigo_pago_detalle varchar(9) NOT NULL,
	par_codigo varchar(5) NULL,
	Pro_programa varchar(2) NULL,
	Pro_subprograma varchar(2) NULL,
	Pro_proyecto varchar(2) NULL,
	Pro_actividad varchar(2) NULL,
	cta_codigo varchar(40) NULL,
	cheque_o_trf varchar(1) NULL,
	numero_cheque_trf varchar(20) NULL,
	cta_codigo_destino varchar(40) NULL,
	cheque_o_trf_destino varchar(1) NULL,
	numero_cheque_trf_destino varchar(20) NULL,
	codigo_beneficiario varchar(15) NULL,
	concepto_pago varchar(80) NULL,
	monto_total float(53) NULL,
	Porcentaje float(53) NULL,
	monto_Bolivianos float(53) NULL,
	monto_Dolares float(53) NULL,
	tipo_cambio float(53) NULL,
	deducciones float(53) NULL,
	saldo_bolivianos float(53) NULL,
	fecha_impresion_cheque datetime NULL,
	fecha_aprobacion_tesoreria datetime NULL,
	fecha_pago smalldatetime NULL,
	departamento varchar(50) NULL,
	estado_aprobacion varchar(1) NULL,
	fecha_autorizacion datetime NULL,
	banco_destino varchar(100) NULL,
	ObsEscrita varchar(255) NULL,
	Observacion varchar(255) NULL,
	honorarios varchar(1) NULL,
	beneficiario_destino varchar(60) NULL,
	codigo_dev float(53) NULL,
	usr_usuario varchar(15) NULL,
	fecha_registro datetime NULL,
	hora_registro varchar(8) NULL,
	literal varchar(255) NULL,
	Fecha_Aprobacion datetime NULL,
	estado_conciliacion varchar(1) NULL
	) ON [PRIMARY]
GO
IF EXISTS(SELECT * FROM dbo.pago_detalle)
	 EXEC('INSERT INTO dbo.Tmp_pago_detalle(Ges_gestion, org_codigo, codigo_pago, codigo_pago_detalle, par_codigo, Pro_programa, Pro_subprograma, Pro_proyecto, Pro_actividad, cta_codigo, cheque_o_trf, numero_cheque_trf, cta_codigo_destino, cheque_o_trf_destino, numero_cheque_trf_destino, codigo_beneficiario, concepto_pago, monto_total, Porcentaje, monto_Bolivianos, monto_Dolares, tipo_cambio, deducciones, saldo_bolivianos, fecha_impresion_cheque, fecha_aprobacion_tesoreria, fecha_pago, departamento, estado_aprobacion, fecha_autorizacion, banco_destino, ObsEscrita, Observacion, honorarios, beneficiario_destino, codigo_dev, usr_usuario, fecha_registro, hora_registro, literal, Fecha_Aprobacion, estado_conciliacion)
		SELECT Ges_gestion, org_codigo, codigo_pago, codigo_pago_detalle, par_codigo, Pro_programa, Pro_subprograma, Pro_proyecto, Pro_actividad, cta_codigo, cheque_o_trf, numero_cheque_trf, cta_codigo_destino, cheque_o_trf_destino, numero_cheque_trf_destino, codigo_beneficiario, concepto_pago, monto_total, Porcentaje, monto_Bolivianos, monto_Dolares, tipo_cambio, deducciones, saldo_bolivianos, fecha_impresion_cheque, fecha_aprobacion_tesoreria, CONVERT(smalldatetime, fecha_pago), departamento, estado_aprobacion, fecha_autorizacion, banco_destino, ObsEscrita, Observacion, honorarios, beneficiario_destino, codigo_dev, usr_usuario, fecha_registro, hora_registro, literal, Fecha_Aprobacion, estado_conciliacion FROM dbo.pago_detalle TABLOCKX')
GO
DROP TABLE dbo.pago_detalle
GO
EXECUTE sp_rename 'dbo.Tmp_pago_detalle', 'pago_detalle'
GO
ALTER TABLE dbo.pago_detalle ADD CONSTRAINT
	PK_pago_detalle PRIMARY KEY NONCLUSTERED 
	(
	Ges_gestion,
	org_codigo,
	codigo_pago,
	codigo_pago_detalle
	) ON [PRIMARY]
GO
GRANT REFERENCES ON dbo.pago_detalle TO saf AS dbo
GRANT SELECT ON dbo.pago_detalle TO saf AS dbo
GRANT INSERT ON dbo.pago_detalle TO saf AS dbo
GRANT DELETE ON dbo.pago_detalle TO saf AS dbo
GRANT UPDATE ON dbo.pago_detalle TO saf AS dbo
COMMIT
