-- =============================================        
-- Autor: JOSE LUIS RODRIGUEZ        
-- Modificado por:         
-- Sistema: ADMIN.EMPRESA        
-- Objeto: PROCEDIMIENTO ALMACENADO        
-- Nombre del objeto:rp_planilla_rciva         
-- Descripcion: PROCEDIMIENTO ALMACENADO PARA REPORTE        
-- Tabla:         
-- Tablas relacionadas:        
-- Modificado por:        
-- Fecha de creacion: 29/07/2016       
-- Fecha de modificacion: 09/08/2016        
-- Version: 1.0        
-- ================================================ 
    -- EXEC rp_planilla_rciva '2016','P01','1'  
ALTER PROCEDURE rp_planilla_rciva       
        
@ges_gestion varchar(4),        
@planilla_codigo varchar(3),        
@mes_grupo int        
AS        
BEGIN        
DECLARE @dos_mnn DECIMAL(12,2) = 2400;     --    3610
WITH C AS     
(    
  SELECT DISTINCT      
  gc_beneficiario.beneficiario_codigo As ci,      
  gc_beneficiario.beneficiario_denominacion AS nombre,      
  ROUND((ro_pagos_cronograma_Detalle.total_ganado - (CASE  WHEN ro_pagos_cronograma_Detalle.afp1 > 0 THEN ro_pagos_cronograma_Detalle.afp1 ELSE ro_pagos_cronograma_Detalle.afp2 END)),0) As sueldo_neto,  
   -- Verifica si sueldo neto es mayor a 2 salarios nacionales.      
  ROUND(CASE WHEN (ro_pagos_cronograma_Detalle.total_ganado - (CASE  WHEN ro_pagos_cronograma_Detalle.afp1 > 0 THEN ro_pagos_cronograma_Detalle.afp1 ELSE ro_pagos_cronograma_Detalle.afp2 END)) > @dos_mnn THEN  
     @dos_mnn  
  ELSE  
    (ro_pagos_cronograma_Detalle.total_ganado - (CASE  WHEN ro_pagos_cronograma_Detalle.afp1 > 0 THEN ro_pagos_cronograma_Detalle.afp1 ELSE ro_pagos_cronograma_Detalle.afp2 END))   
  END,0)    As minno_imponible,   -- CALCULO dos minimos nacioneal B  
  ISNULL( ro_pagos_cronograma_Detalle.iva_110,0 ) As iva110,       
  ISNULL( ro_pagos_cronograma_Detalle.fisco_a_favor,0) As fisco_favor,          
  ISNULL( ro_pagos_cronograma_Detalle.mes_anterior_mant,0) As mesa_mant,    -- ENTRADA  
  ISNULL( ro_pagos_cronograma_Detalle.saldo_util,0) As saldo_util,      
  ISNULL( ro_pagos_cronograma_Detalle.saldo_a_favor_depend,0) As saldo_fav_depend,    
  -- Saldo favor mes anterior     
  ISNULL((SELECT TOP 1 saldo_a_favor_depend FROM ro_pagos_cronograma_Detalle ant WHERE  ant.beneficiario_codigo = gc_beneficiario.beneficiario_codigo  AND ant.mes_grupo = (@mes_grupo - 1)),0) AS saldof_mesant       
  FROM         dbo.gc_beneficiario  INNER JOIN        
                      dbo.ro_personal_contratado ON dbo.gc_beneficiario.beneficiario_codigo = dbo.ro_personal_contratado.beneficiario_codigo INNER JOIN        
                      dbo.ro_pagos_cronograma_Detalle ON dbo.gc_beneficiario.beneficiario_codigo = dbo.ro_pagos_cronograma_Detalle.beneficiario_codigo INNER JOIN        
                      dbo.rc_cargos ON dbo.ro_personal_contratado.cargo_codigo = dbo.rc_cargos.cargo_codigo INNER JOIN        
                      dbo.rc_puestos ON dbo.ro_personal_contratado.puesto_codigo = dbo.rc_puestos.puesto_codigo INNER JOIN        
                      dbo.ro_pagos_cronograma ON dbo.ro_pagos_cronograma_Detalle.ges_gestion = dbo.ro_pagos_cronograma.ges_gestion INNER JOIN        
                      dbo.gc_unidad_ejecutora ON dbo.ro_personal_contratado.unidad_codigo = dbo.gc_unidad_ejecutora.unidad_codigo INNER JOIN        
                      dbo.rc_planilla_grupo ON dbo.ro_pagos_cronograma_Detalle.planilla_codigo = dbo.rc_planilla_grupo.planilla_codigo        
                      WHERE  ( ro_pagos_cronograma_Detalle.ges_gestion like @ges_gestion  and ro_pagos_cronograma_Detalle.planilla_codigo like @planilla_codigo and ro_pagos_cronograma_Detalle.mes_grupo like @mes_grupo)                            
),     
y AS     
(    
SELECT       
      ci,       
      nombre,        
      sueldo_neto,  
      minno_imponible,        
      ROUND((sueldo_neto - minno_imponible),0) As difsuj_impuesto,  -- CALCULO diferencia sujeta impuesto C    
      ROUND(((sueldo_neto - minno_imponible) * 0.13),0) As impto13,  -- CALCULO impuesto 13% D    
      iva110,    
      -------------------------------------   
      CASE WHEN ROUND((minno_imponible * 0.13),0) > ROUND((sueldo_neto - minno_imponible),0) THEN
			ROUND(((sueldo_neto - minno_imponible) * 0.13),0)
      ELSE
            ROUND((minno_imponible * 0.13),0)
      END  As dosmn13, -- CALCULO 13% dos salarios minimos F      
      -------------------------------------    
      ROUND(CASE WHEN ( ((sueldo_neto - minno_imponible) * 0.13)  
             - iva110 - (minno_imponible * 0.13)      
                  ) > 0 THEN    
         ( ((sueldo_neto - minno_imponible) * 0.13) - iva110 - (minno_imponible * 0.13) )  
      ELSE    
          0        
      END,0) AS fisco_favor,   -- CALCULO saldo favor fisco G   
      --------------------------------------    
      ROUND(CASE WHEN ( ( ((sueldo_neto - minno_imponible) * 0.13)  
                   - iva110 - (minno_imponible * 0.13)      
                  )) <= 0 THEN    
          (( (minno_imponible * 0.13) + iva110) -  ((sueldo_neto - minno_imponible) * 0.13) )         
           
      ELSE    
          0        
      END, 0) AS depend_favor,    -- CALCULO saldo favor dependiente  H     
      (saldof_mesant) AS mesa,      
       mesa_mant,      
      ROUND((saldof_mesant + mesa_mant),0) As mesa_total -- CALCULO saldo favor total  K   
FROM C      
    
)    
-- Seleccion de datos para reporte    
SELECT     
      ci,       
      nombre,        
      sueldo_neto,       
      minno_imponible,       
      difsuj_impuesto,      
      impto13 ,      
      iva110,      
      dosmn13,     
      fisco_favor,    
      depend_favor,    
      mesa,      
      mesa_mant,      
      mesa_total,      
      (depend_favor + mesa_total) As saldof_depend,      
      ROUND(CASE WHEN fisco_favor >= (depend_favor + mesa_total) THEN (depend_favor + mesa_total) ELSE 0 END,0) AS saldo_util,      
      ROUND(CASE WHEN fisco_favor >= 0 THEN fisco_favor - (CASE WHEN fisco_favor >= (depend_favor + mesa_total) THEN (depend_favor + mesa_total) ELSE 0 END) ELSE 0 END,0) AS imp_rete_pagar,      
      ROUND(CASE WHEN (depend_favor + mesa_total) >= 0 THEN (depend_favor + mesa_total) - (CASE WHEN fisco_favor >= (depend_favor + mesa_total) THEN (depend_favor + mesa_total) ELSE 0 END) ELSE 0 END,0) AS saldo_fav_depend          
FROM y    
    
END 