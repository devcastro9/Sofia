
-- No existe
IF NOT EXISTS(SELECT * FROM sys.columns
WHERE Name = N'beneficiario_nombre_ref' AND OBJECT_ID = OBJECT_ID(N'ao_solicitud_bitacora'))
BEGIN
     ALTER TABLE ao_solicitud_bitacora
    ADD beneficiario_nombre_ref VARCHAR(350)
    PRINT ' Adicionamos columna beneficiario_nombre_ref'
END  

-- No existe
IF NOT EXISTS(SELECT * FROM sys.columns
WHERE Name = N'beneficiario_codigo_cgi' AND OBJECT_ID = OBJECT_ID(N'ao_solicitud_bitacora'))
BEGIN
     ALTER TABLE ao_solicitud_bitacora
    ADD beneficiario_codigo_cgi VARCHAR(350)
    PRINT ' Adicionamos columna beneficiario_codigo_cgi'
END  

