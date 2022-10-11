VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000003&
   Caption         =   "SAF - 2000"
   ClientHeight    =   6135
   ClientLeft      =   885
   ClientTop       =   2820
   ClientWidth     =   9840
   Icon            =   "frmMain1.frx":0000
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5865
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11721
            Text            =   "Estado"
            TextSave        =   "Estado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "01/08/2000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:32 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu Clasificadores 
      Caption         =   "Clasificadores"
   End
   Begin VB.Menu Clasificadores2 
      Caption         =   "Clasificadores2"
      Visible         =   0   'False
      Begin VB.Menu Generales 
         Caption         =   "Generales"
         Begin VB.Menu UnidadesEjecutoras 
            Caption         =   "&Unidades Ejecutoras"
            Shortcut        =   ^U
         End
         Begin VB.Menu Entidades 
            Caption         =   "E&ntidades"
            Shortcut        =   ^N
         End
         Begin VB.Menu TipoTramite 
            Caption         =   "Tipo de Trá&mite (Formularios)"
         End
         Begin VB.Menu CBeneficiarios 
            Caption         =   "Beneficiarios"
         End
         Begin VB.Menu DepartamentosBolivia 
            Caption         =   "Departamentos de Bolivia"
         End
         Begin VB.Menu ProvinciasDepartamentos 
            Caption         =   "Provincias de Departamentos"
         End
         Begin VB.Menu TiposErrores 
            Caption         =   "Tipos de Errores en Documentos"
         End
      End
      Begin VB.Menu CPresupuesto 
         Caption         =   "Gastos"
         Begin VB.Menu PartidasGasto 
            Caption         =   "&Partidas del Gasto"
            Shortcut        =   ^P
         End
         Begin VB.Menu EconómicosGasto 
            Caption         =   "Económicos del Gasto"
         End
         Begin VB.Menu RelacionadorGastoEco 
            Caption         =   "Relacionador Gasto Eco. Vs. Partida"
         End
      End
      Begin VB.Menu Presupuesto 
         Caption         =   "Presupuesto"
         Begin VB.Menu FuentesFinanciamiento 
            Caption         =   "&Fuentes de Financiamiento"
            Shortcut        =   ^F
         End
         Begin VB.Menu OrganismosFinanciadores 
            Caption         =   "&Organismos Financiadores"
            Shortcut        =   ^O
         End
         Begin VB.Menu Convenios 
            Caption         =   "&Convenios con los Financiadores"
            Shortcut        =   ^C
         End
         Begin VB.Menu CategoriaFinanciadores 
            Caption         =   "Ca&tegoria de Financiadores"
            Shortcut        =   ^T
         End
         Begin VB.Menu EstructuraProgramatica 
            Caption         =   "&Estructura Programática"
            Shortcut        =   ^E
         End
         Begin VB.Menu Sisin 
            Caption         =   "S&isin"
            Shortcut        =   ^I
         End
         Begin VB.Menu Deducciones 
            Caption         =   "Deducciones/Retenciones"
         End
         Begin VB.Menu RelacionadorDonacionesOrganismos 
            Caption         =   "Relacionador Donaciones - Organismos"
         End
      End
      Begin VB.Menu CTesoreria 
         Caption         =   "Tesoreria"
         Begin VB.Menu CBancos 
            Caption         =   "Bancos"
         End
         Begin VB.Menu CCtaBancarias 
            Caption         =   "Cuentas Bancarias"
         End
         Begin VB.Menu FormasPago 
            Caption         =   "Formas de Pago"
         End
      End
      Begin VB.Menu Ingresos 
         Caption         =   "Ingresos"
         Begin VB.Menu Rubros 
            Caption         =   "Rubros"
         End
         Begin VB.Menu EconómicosRecursos 
            Caption         =   "Económicos de Recursos"
         End
         Begin VB.Menu RelacionadorRubroEco 
            Caption         =   "Relacionador Rubro Eco. Vs Recursos"
         End
      End
      Begin VB.Menu Contabilidad2 
         Caption         =   "Contabilidad"
         Begin VB.Menu PlanCuentas 
            Caption         =   "Plan de Cuentas"
         End
         Begin VB.Menu RelacionadorCuentaPartidas 
            Caption         =   "Relacionador Cuenta Vs Partidas"
         End
         Begin VB.Menu RelacionadorIngresosCuentas 
            Caption         =   "Relacionador Cuentas Vs Ingresos"
         End
         Begin VB.Menu Depreciaciones 
            Caption         =   "Depreciaciones"
         End
         Begin VB.Menu ClaseAuxiliares 
            Caption         =   "Clase de Auxiliares"
         End
         Begin VB.Menu Inversiones 
            Caption         =   "Inversiones"
         End
      End
      Begin VB.Menu Administrativos 
         Caption         =   "Administrativos"
         Begin VB.Menu Adquisiciones 
            Caption         =   "Adquisiciones"
         End
         Begin VB.Menu Contrataciones 
            Caption         =   "Contrataciones"
         End
         Begin VB.Menu Almacenes2 
            Caption         =   "Almacenes"
         End
         Begin VB.Menu RecursosHumanos2 
            Caption         =   "Recursos Humanos"
         End
      End
   End
   Begin VB.Menu MesaEntrada 
      Caption         =   "Mesa de Entrada"
      Begin VB.Menu RegistroSolicitudes 
         Caption         =   "Registro Solicitudes"
         Begin VB.Menu FormularioF01 
            Caption         =   "Formulario F01"
         End
         Begin VB.Menu FormularioF02 
            Caption         =   "Formulario F02"
         End
         Begin VB.Menu FormularioF03 
            Caption         =   "FormularioF03"
         End
         Begin VB.Menu FormularioF04 
            Caption         =   "FormularioF04"
         End
         Begin VB.Menu FormularioF05 
            Caption         =   "FormularioF05"
         End
         Begin VB.Menu FormularioF06 
            Caption         =   "FormularioF06"
         End
         Begin VB.Menu FormularioF07 
            Caption         =   "FormularioF07"
         End
      End
      Begin VB.Menu Importar 
         Caption         =   "Importar Solicitudes"
      End
      Begin VB.Menu Exportar 
         Caption         =   "Exportar Solicitudes"
      End
      Begin VB.Menu ProgConAdq 
         Caption         =   "Programación Contrataciones y Adquisiciones"
      End
   End
   Begin VB.Menu Procesos 
      Caption         =   "Egresos"
      Begin VB.Menu Compromiso 
         Caption         =   "Ejecucion Presupuestaria"
      End
      Begin VB.Menu Ejecucion 
         Caption         =   "Reportes de E&jecucion"
         Begin VB.Menu EjecucionOrganismo 
            Caption         =   "Ejecucion Por Organismo, Convenio y Categoría"
         End
         Begin VB.Menu EjecucionProyVsPpto 
            Caption         =   "Ejecucion por Organismo, Proyecto Vs. Ppto. de Ley"
         End
         Begin VB.Menu EjecucionUniOrgProyPar 
            Caption         =   "Ejecución por Unidad, Organismo, Proyecto y Partida"
         End
         Begin VB.Menu EjecuciónPorUni 
            Caption         =   "Ejecución del por Unidad"
         End
         Begin VB.Menu RepGraf 
            Caption         =   "Reportes Gráficos"
            Begin VB.Menu repGrafUnidad 
               Caption         =   "Por unidad"
            End
            Begin VB.Menu repGraforga 
               Caption         =   "Por Organismo"
            End
         End
      End
      Begin VB.Menu MnuModPpto 
         Caption         =   "Modificaciones Presupuestarias"
      End
   End
   Begin VB.Menu Tesoreria 
      Caption         =   "Tesorería"
      Begin VB.Menu pp 
         Caption         =   "Pagos Pendientes"
      End
      Begin VB.Menu OperaciónCheques 
         Caption         =   "Operación de Cheques"
      End
      Begin VB.Menu CuentasBancarias2 
         Caption         =   "Movimientos"
         Begin VB.Menu MovCta 
            Caption         =   "Movimiento de Cuentas Bancarias"
         End
         Begin VB.Menu SaldosActuales 
            Caption         =   "Saldos Actuales"
         End
      End
      Begin VB.Menu ic 
         Caption         =   "Traspasos"
      End
      Begin VB.Menu Gastos 
         Caption         =   "Gastos"
      End
      Begin VB.Menu ConsultasPagos 
         Caption         =   "Consultas Pagos"
         Begin VB.Menu PagosEfectuados 
            Caption         =   "Pagos Efectuados"
         End
         Begin VB.Menu PagosEfectuadosRealizar 
            Caption         =   "Pagos Efectuados y por Realizar"
         End
      End
      Begin VB.Menu ManejoCheques 
         Caption         =   "Impresión de Comprobantes de Pago"
         Begin VB.Menu PorColaImpresion 
            Caption         =   "Por Cola de Impresión"
         End
         Begin VB.Menu PorSeleccionComprobantes 
            Caption         =   "Por Selección de Comprobantes"
         End
      End
      Begin VB.Menu ct 
         Caption         =   "Impresión de Transferencias"
      End
      Begin VB.Menu MnuImpCheques 
         Caption         =   "Impresión de cheques"
      End
      Begin VB.Menu MnuCorrCheques 
         Caption         =   "Correlativo de cheques"
      End
      Begin VB.Menu cmpte_nuevo 
         Caption         =   "Imprime Cmpte Nuevo"
      End
   End
   Begin VB.Menu CContabilidad 
      Caption         =   "Contabilidad"
      Begin VB.Menu Comprobantes 
         Caption         =   "Registro Manual"
      End
      Begin VB.Menu ReportesC 
         Caption         =   "Reportes"
         Begin VB.Menu LibroMayor 
            Caption         =   "Libro Mayor"
         End
         Begin VB.Menu LibroMayorAuxiliar 
            Caption         =   "Libro Mayor Auxiliar"
         End
         Begin VB.Menu EstadoResultados 
            Caption         =   "Estado de Resultados"
         End
         Begin VB.Menu BalanceSumasSaldos 
            Caption         =   "Balance de Sumas y Saldos"
         End
         Begin VB.Menu BalanceGeneral 
            Caption         =   "Balance General"
         End
      End
   End
   Begin VB.Menu MnuIngresos 
      Caption         =   "Ingresos"
      Begin VB.Menu EjecucionPresupuestaria 
         Caption         =   "Ejecución Presupuestaria"
      End
      Begin VB.Menu ReportesIngresos 
         Caption         =   "Reportes Ingresos"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Administracion 
      Caption         =   "Procesos Administrativos"
      Begin VB.Menu AdquisicionBienes 
         Caption         =   "Adquisición de Bienes"
         Begin VB.Menu ComprasDirectas 
            Caption         =   "Compras Directas"
         End
         Begin VB.Menu LicitacionesNacionales 
            Caption         =   "Licitaciones Nacionales"
         End
         Begin VB.Menu LicitacionesInternacionales 
            Caption         =   "Licitaciones Internacionales"
         End
      End
      Begin VB.Menu ContratacionServicios 
         Caption         =   "Contratación de Servicios"
         Begin VB.Menu ConsultoresIndividuales 
            Caption         =   "Consultores Individuales"
         End
         Begin VB.Menu EmpresasConsultoras 
            Caption         =   "Empresas Consultoras"
         End
      End
      Begin VB.Menu Almacenes 
         Caption         =   "Almacenes"
         Begin VB.Menu IngresosA 
            Caption         =   "Ingresos"
         End
         Begin VB.Menu SalidasA 
            Caption         =   "Salidas"
         End
      End
   End
   Begin VB.Menu RecursosHumanos 
      Caption         =   "Recursos Humanos"
      Begin VB.Menu AdministracionPersonal 
         Caption         =   "Administración Personal"
      End
      Begin VB.Menu ControlPersonal 
         Caption         =   "Control de Personal"
      End
      Begin VB.Menu CapacitacionPersonal 
         Caption         =   "Capacitación de Personal"
      End
      Begin VB.Menu EvaluacionDesempeño 
         Caption         =   "Evaluación de Desempeño"
      End
   End
   Begin VB.Menu InformacionGerencial 
      Caption         =   "Informacion Gerencial"
   End
   Begin VB.Menu mnuAdmiSistema 
      Caption         =   "Administracion Sistema"
      Begin VB.Menu mnuCambiarClave 
         Caption         =   "Cambiar Clave"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnuNivelAcceso 
         Caption         =   "Nivel de Acceso"
      End
      Begin VB.Menu mnuPrivAcceso 
         Caption         =   "Privilegios de Operación"
      End
   End
   Begin VB.Menu Salida 
      Caption         =   "Salida"
      Begin VB.Menu mnuAcercade 
         Caption         =   "A cerca de ..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepara 
         Caption         =   "-------------------"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir del Sistema"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------- Programas en comentario de JORGE

Private Sub CuentasBancarias_Click()
    Dim e As Long
    e = Shell("c:\SAF-2000\Reportes Tesoreria\cnsPagados.exe", 1)
End Sub

Private Sub BalanceGeneral_Click()
frmBalanceGral.Show
End Sub

Private Sub BalanceSumasSaldos_Click()
frmsumsaldos.Show
End Sub

'Private Sub CBeneficiarios_Click() 'clasifica
'  frmBeneficiario.Show
'End Sub

Private Sub CCtaBancarias_Click()
  CLFrmCtaBco.Show
End Sub

Private Sub Clasificadores_Click()
'  FrmCtaBco.Show
  If UCase(GlUsuario) = "SAF" Or UCase(GlUsuario) = "M_YAÑEZ" Or UCase(GlUsuario) = "A_OZINAGA" Then
    Dim e As Long
    e = Shell("c:\SAF-2000\Clasificadores\clasificadores.exe", 1)
  Else
      MsgBox "El usuario no tiene acceso !", vbInformation + vbOKOnly
  End If
End Sub

Private Sub cmpte_nuevo_Click()
  FrmImprimeComprobanteNuevo.Show
End Sub

Private Sub EjecuciónPorUni_Click()
  glRepPresup = "REP004"
  frmRepPresupuesto.Show
End Sub

Private Sub EjecucionProyVsPpto_Click()
  glRepPresup = "REP002"
  frmRepPresupuesto.Show
End Sub

Private Sub EjecucionUniOrgProyPar_Click()
  glRepPresup = "REP003"
  frmRepPresupuesto.Show
End Sub

'Private Sub Entidades_Click() 'clasifica
'  Dim e As New Entidades
'  e.Show '  Entidades.se
'End Sub

Private Sub EstadoResultados_Click()
  frmEERR.Show
End Sub

Private Sub Exportar_Click()
  Frmexporta.Show vbModal
End Sub

Private Sub FormularioF01_Click()
  FrmF01.Show vbModal
End Sub

Private Sub FormularioF02_Click()
  FrmF02.Show vbModal
End Sub

Private Sub FormularioF03_Click()
  FrmF03.Show vbModal
End Sub

Private Sub FormularioF04_Click()
  FrmF04.Show vbModal
End Sub

Private Sub FormularioF05_Click()
  FrmF05.Show vbModal
End Sub

Private Sub FormularioF06_Click()
  FrmF06.Show vbModal
End Sub

Private Sub FormularioF07_Click()
  FrmF07.Show vbModal
End Sub

Private Sub Gastos_Click()
  'FrmCuentaBancaria.Show
  FrmListadoPagos.Show
End Sub

Private Sub ic_Click()
  ' con Gabriela
  frmtraspasos.Show
End Sub

Private Sub Importar_Click()
  FrmImporta.Show vbModal
End Sub

Private Sub LibroMayor_Click()
  frmLMayor.Show
End Sub

Private Sub LibroMayorAuxiliar_Click()
  FrmLMayorAux.Show
End Sub

Private Sub mnuAcercade_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuCambiarClave_Click()
    FrmCambiarClave.Show vbModal
    'MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

Private Sub MnuCorrCheques_Click()
  FrmCorrelativos.Show
End Sub

Private Sub MnuImpCheques_Click()
'    Dim e As Long
'    e = Shell("c:\SAF-2000\Reportes Tesoreria\cnsTotal.exe", 1)
    FrmChequesNuevo.Show
End Sub

'Private Sub Aprobacion_Click()
'   Dim f As New FrmCompromisoA
'   f.Show
'End Sub
'
'Private Sub CategoriaFinanciadores_Click()
'   Dim f As New CategoriaFinanciador
'   f.Show
'End Sub
'
'Private Sub CBancos_Click()
'   Dim f As New FrmBanco
'   f.Show
'End Sub
'
'Private Sub CBeneficiarios_Click()
'   Dim f As New frmBeneficiario
'   f.Show
'End Sub
'
'Private Sub CCtaBancarias_Click()
'   Dim f As New FrmCtaBco
'   f.Show
'End Sub
'
'Private Sub Compromiso_Click()
'   Dim f As New FrmCompromiso
'   f.Show
'End Sub
'
'Private Sub ConciliacionCtaEspecial_Click()
'   Dim f As New RepConciliacionEspecial04
'   f.Show
'End Sub
'
'Private Sub Convenios_Click()
'   Dim f As New Convenios
'   f.Show
'End Sub
'
'Private Sub Convenios03_Click()
'   Dim f As New RepConvenio03
'   f.Show
'End Sub
'
'Private Sub DesembolsoBid_Click()
'   Dim f As New RepDesembolso06
'   f.Show
'End Sub
'
'Private Sub CPlanCuentas_Click()
'   Dim f As New frmCl_Cuentas
'   f.Show
'End Sub
'
'Private Sub cRelacionadorPartidaCta_Click()
'   Dim f As New FrmC_Relacionador
'   f.Show
'End Sub
'
'Private Sub Devengado_Click()
'   Dim f As New FrmDevengado
'   f.Show
'End Sub
'
'Private Sub DistribucionPorcentual_Click()
'   Dim f As New PorcentajeFtePartida
'   f.Show
'End Sub
'
'Private Sub EjecucionAcumulada_Click()
'   Dim f As New RepEjecucionAcum22a
'   f.Show
'End Sub
'
'Private Sub EjecucionSuecia_Click()
'   Dim f As New RepEjecucionSuecia20a
'   f.Show
'End Sub
'
'Private Sub EjecucionSueciaCat_Click()
'   Dim f As New RepEjecutadoCat19a
'   f.Show
'End Sub
'
'Private Sub EjecucionSueciaDep_Click()
'   Dim f As New RepEjecutadoDep17
'   f.Show
'End Sub
'
'Private Sub EjecucionSueciaUni_Click()
'   Dim f As New RepEjecutadoUni18
'   f.Show
'End Sub
'
'Private Sub Entidades_Click()
'   Dim f As New Entidades
'   f.Show
'End Sub
'
'Private Sub EstructuraProgramatica_Click()
'   Dim f As New EstructuraProgrmatica
'   f.Show
'End Sub
'
'Private Sub FuentesFinanciamiento_Click()
'   Dim f As New FuenteFinanciamiento
'   f.Show
'End Sub
'
'Private Sub InformeSemestralBID_Click()
'   Dim f As New RepInformeBID09
'   f.Show
'End Sub

'-------------------------------------------
'VERIFICA .............

'Private Sub CategoriaFinanciadores_Click()
'   Dim f As New CategoriaFinanciador
'   f.Show
'End Sub
'
'Private Sub CBancos_Click()
'   Dim f As New FrmBanco
'   f.Show
'End Sub
'
'Private Sub CBeneficiarios_Click()
'   Dim f As New frmBeneficiario
'   f.Show
'End Sub
'
'Private Sub CCtaBancarias_Click()
'   Dim f As New FrmCtaBco
'   f.Show
'End Sub
'
'Private Sub ClaseAuxiliares_Click()
'   Dim f As New frmclaseauxiliar
'   f.Show
'End Sub
'
'Private Sub Convenios_Click()
'   Dim f As New Convenios
'   f.Show
'End Sub
'
'Private Sub Deducciones_Click()
'    frmDeducciones.Show
'End Sub
'
'Private Sub Depreciaciones_Click()
'    frmdepreciacion.Show
'End Sub
'
'Private Sub EconómicosRecursos_Click()
'    frmEcoRecurso.Show
'End Sub
'
Private Sub EjecucionOrganismo_Click()
  glRepPresup = "REP001"
  frmRepPresupuesto.Show
End Sub
'
'Private Sub Inversiones_Click()
'    frmInversiones.Show
'End Sub
'
'End Sub
'
'Private Sub mnuCambiarClave_Click()
'    FrmCambiarClave.Show
'End Sub

Private Sub MnuModPpto_Click()
  If UCase(GlUsuario) = "FFL001" Or UCase(GlUsuario) = "M_YAÑEZ" Or UCase(GlUsuario) = "A_ITURRI" Or UCase(GlUsuario) = "I_IMAÑA" Or UCase(GlUsuario) = "F_FLORES" Or UCase(GlUsuario) = "SAF" Then
    FrmModPresup.Show
  Else
    MsgBox "El usuario, no tiene acceso", vbCritical + vbOKOnly, "Acceso restringido"
  End If
End Sub

Private Sub mnuNivelAcceso_Click()
    FrmNivelesAcceso.Show
'    MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

Private Sub mnuSalir_Click()
   Unload Me
   End
End Sub

Private Sub mnuUsuarios_Click()
    FrmSisUsuarios.Show
    'MsgBox "El usuario no tiene acceso", vbInformation + vbCritical
End Sub

'Private Sub PlanCuentas_Click()
'   Dim f As New frmCl_Cuentas
'   f.Show
'End Sub
''
''Private Sub cRelacionadorPartidaCta_Click()
''   Dim f As New FrmC_Relacionador
''   f.Show
''End Sub
'
'Private Sub Entidades_Click()
'   Dim f As New Entidades
'   f.Show
'End Sub
'
'Private Sub EstructuraProgramatica_Click()
'   Dim f As New EstructuraProgrmatica
'   f.Show
'End Sub
'
'
'Private Sub FormasPago_Click()
'    frmFormaPago.Show
'End Sub
'
'Private Sub FuentesFinanciamiento_Click()
'   Dim f As New FuenteFinanciamiento
'   f.Show
'End Sub
'
'Private Sub OrganismosFinanciadores_Click()
'    OrganismoFinanciador.Show
'End Sub
'
'Private Sub ac_Click()
'    If NombreUsuario = "jcc001" Or NombreUsuario = "JCC002" Or NombreUsuario = "MYB159" Then
'        FrmActivacionCheques.Show
'    Else
'        MsgBox "El Usuario NO está autorizado . . . "
'    End If
'End Sub
'
'Private Sub ApruebaComprobante_Click()
'    If usuario2 = "RAG001" Or usuario2 = "rag001" Then
'        FrmApruebaR.Show
'    Else
'        MsgBox "El Usuario NO está autorizado . . . "
'    End If
'End Sub
'
Private Sub Comprobantes_Click()
'    Dim e As Long
'    usuario2 = frmLogin.txtUserName.Text
'    If usuario2 = "cl001" Or usuario2 = "CL001" Or usuario2 = "ram001" Or usuario2 = "RAM001" Or usuario2 = "MYB159" Then
'        e = Shell("D:\saf-2000\Contabilidad\Prueba_Conta.exe", 1)
'    Else
'        MsgBox "El Usuario NO Tiene Acceso !! ..."
'    End If
    frm_ManualConta.Show
End Sub


Private Sub Compromiso_Click()
'ppto
'    Dim e As Long
'    usuario2 = frmLogin.txtUserName.Text
'    If usuario2 = "fff777" Or usuario2 = "FFF777" Or usuario2 = "jqa001" Or usuario2 = "MYB159" Then
'    FrmRegularizacion.Show
'        e = Shell("D:\saf-2000\Ppto\PPTO.exe", 1)
'    Else
'        MsgBox "El Usuario NO Tiene Acceso !! ..."
'    End If

  BuscaTipoAcceso "compromiso"
  FrmRegularizacion.Show
End Sub

Private Sub ct_Click()
    FrmTransferenciasNuevo.Show
End Sub

Private Sub dc_Click()
    If NombreUsuario = "jcc001" Or NombreUsuario = "JCC002" Or NombreUsuario = "MYB159" Then
        FrmDesactivacionCheques.Show
    Else
        MsgBox "El Usuario NO está autorizado . . . "
    End If
End Sub

'Private Sub DepartamentosBolivia_Click() 'clasifica
'    frmDepto.Show
'End Sub

Private Sub EconómicosGasto_Click()
    frmEcoGasto.Show
End Sub

Private Sub EjecucionPresupuestaria_Click()
'    usuario2 = frmLogin.txtUserName.Text
    usuario2 = GlUsuario
'    If UCase(GlUsuario) = "FFL001" Or UCase(GlUsuario) = "F_FLORES" Or UCase(GlUsuario) = "A_ITURRI" Or UCase(GlUsuario) = "J_CRUZ" Or (GlUsuario) = "F_Flores" Or GlUsuario = "J_CAMACHO" Or GlUsuario = "J_Camacho" Or GlUsuario = "M_YAÑEZ" Or UCase(GlUsuario) = "SAF" Then
      FrmIngresosabm.Show
'    Else
'        MsgBox "El Usuario NO Tiene Acceso !! ..."
'    End If
End Sub


'Private Sub ic_Click()
'    Dim e As Long
'    usuario2 = frmLogin.txtUserName.Text
'    If usuario2 = "cl001" Or usuario2 = "CL001" Or usuario2 = "ram001" Or usuario2 = "RAM001" Or usuario2 = "MYB159" Then
'        e = Shell("D:\saf-2000\Contabilidad\Prueba_Conta.exe", 1)
'    Else
'        MsgBox "El Usuario NO Tiene Acceso !! ..."
'    End If
'End Sub

'Private Sub ManejoCheques_Click()
'    FrmActivacionCheques.Show
'End Sub

Private Sub MDIForm_Load()
   Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
   Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
   Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
   Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
   LoadNewDoc
End Sub

Private Sub LoadNewDoc()
   'Static lDocumentCount As Long
   'Dim frmD As frmDocument
   'lDocumentCount = lDocumentCount + 1
   'Set frmD = New frmDocument
   'frmD.Caption = "Document " & lDocumentCount
   'frmD.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If Me.WindowState <> vbMinimized Then
      SaveSetting App.Title, "Settings", "MainLeft", Me.Left
      SaveSetting App.Title, "Settings", "MainTop", Me.Top
      SaveSetting App.Title, "Settings", "MainWidth", Me.Width
      SaveSetting App.Title, "Settings", "MainHeight", Me.Height
   End If
End Sub

Private Sub mnuDatafrmDocument_Click()
'   Dim f As New frmDocument
'   f.Show
End Sub

Private Sub PCE_Click()
Flag_TipComp = "PCE"

End Sub

Private Sub MovCta_Click()
  FrmCuentas.Show
End Sub

Private Sub OperaciónCheques_Click()
    FrmActivacionCheques.Show
End Sub

Private Sub PagosEfectuados_Click()
    FrmPagosRealizados.Show
End Sub

Private Sub PagosEfectuadosRealizar_Click()
    FrmPagosTotal.Show
End Sub

Private Sub PorColaImpresion_Click()
    FrmColaImpresion.Show
End Sub

Private Sub PorSeleccionComprobantes_Click()
    FrmImprimirComprobante.Show
End Sub
Private Sub mnuPrivAcceso_Click()
    frmPrivAcceso.Show
End Sub

'Private Sub pp_Click()
'If NombreUsuario = "rag001" Or NombreUsuario = "RAG001" Or NombreUsuario = "MYB159" Then
''If NombreUsuario = "jcc001" Or NombreUsuario = "JCC001" Then
'    FrmCP.Show
'  Else
'    MsgBox "El usuario NO tiene Acceso !! ...", vbInformation + vbCritical, "Validación"
'  End If
'End Sub

Private Sub pp_Click()
    FrmCP.Show
End Sub

'Private Sub OrganismosFinanciadores_Click()
'   Dim f As New OrganismoFinanciador
'   f.Show
'End Sub
'
'Private Sub PartidasGasto_Click()
'   Dim f As New PartidasGasto
'   f.Show
'End Sub
'
'Private Sub PartidasGasto_Click()
'   Dim f As New PartidasGasto
'   f.Show
'End Sub

Private Sub Regularizacion_Click()
   FrmRegularizacion.Show
End Sub

'Private Sub ReversionCompromiso_Click()
'   Dim f As New FrmCompromisoR
'   f.Show
'End Sub
'
'Private Sub ReversionDevengado_Click()
'   Dim f As New FrmDevengadoR
'   f.Show
'End Sub


'Private Sub ProvinciasDepartamentos_Click() 'clasifica
'    frmProv.Show
'End Sub

Private Sub RelacionadorCuentaPartidas_Click()
    FrmC_Relacionador.Show
End Sub

Private Sub RelacionadorDonacionesOrganismos_Click()
    frmRelDonOrg.Show
End Sub

Private Sub RelacionadorGastoEco_Click()
    frmRelGastoEcoPar.Show
End Sub

Private Sub RelacionadorIngresosCuentas_Click()
    frmRelingresoCta.Show
End Sub

Private Sub RelacionadorRubroEco_Click()
    frmRelRecEcoRec.Show
End Sub

Private Sub repGraforga_Click()
  glRepPresup = "repGraf002"
  frmRepPresGlobal.Show
End Sub

Private Sub repGrafUnidad_Click()
  glRepPresup = "repGraf001"
  frmRepPresGlobal.Show
End Sub

Private Sub ReportesIngresos_Click()
'  FrmModPresup.Show
End Sub

Private Sub SaldosActuales_Click()
  FrmSaldosReales.Show
End Sub

'Private Sub Salida_Click()
'   Unload Me
'   End
'End Sub

'Private Sub Sisin_Click()
'   Dim f As New Sisin
'   f.Show
'End Sub
'
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
   On Error Resume Next
   Select Case Button.Key
      Case "Nuevo"
         LoadNewDoc
      Case "Abrir"
         'TareasPendientes: Agregar código de botón 'Abrir'.
         MsgBox "Agregar código de botón 'Abrir'."
      Case "Guardar"
         'TareasPendientes: Agregar código de botón 'Guardar'.
         MsgBox "Agregar código de botón 'Guardar'."
      Case "Imprimir"
         'TareasPendientes: Agregar código de botón 'Imprimir'.
         MsgBox "Agregar código de botón 'Imprimir'."
      Case "Cortar"
         'TareasPendientes: Agregar código de botón 'Cortar'.
         MsgBox "Agregar código de botón 'Cortar'."
      Case "Copiar"
         'TareasPendientes: Agregar código de botón 'Copiar'.
         MsgBox "Agregar código de botón 'Copiar'."
      Case "Pegar"
         'TareasPendientes: Agregar código de botón 'Pegar'.
         MsgBox "Agregar código de botón 'Pegar'."
      Case "Negrita"
         ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
         Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
      Case "Cursiva"
         ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
         Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
      Case "Subrayado"
         ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
         Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
      Case "Alinear a la izquierda"
         ActiveForm.rtfText.SelAlignment = rtfLeft
      Case "Centrar"
         ActiveForm.rtfText.SelAlignment = rtfCenter
      Case "Alinear a la derecha"
         ActiveForm.rtfText.SelAlignment = rtfRight
   End Select
End Sub

'Private Sub TipoTramite_Click()
'   Dim f As New frmFormulario
'   f.Show
'End Sub
'
'Private Sub UnidadesEjecutoras_Click()
'   Dim f As New UnidadEjecutora
'   f.Show
'End Sub


'Private Sub TiposErrores_Click() 'clasifica
'    frmErrores.Show
'End Sub

'Private Sub TipoTramite_Click()
'    frmFormulario.Show
'End Sub
'
'Private Sub UnidadesEjecutoras_Click()
'    Unidad.Show
'End Sub

'*********************** Freddy ****************************
Public Sub NivelAcceso(vNivelAcceso As Integer)
'Subrutina que habilita o deshabilita las opciones de menu
On Error Resume Next
Dim vNombOpcMenu As String

rsNivelAcceso.Open "Select * From NivelAcceso Where IdNivelAcceso=" & vNivelAcceso, db, adOpenStatic
If rsNivelAcceso.RecordCount > 0 Then
    While Not rsNivelAcceso.EOF
    vNombOpcMenu = LCase(rsNivelAcceso!NombOpcMenu)
    If vNombOpcMenu = "clasificadores" Then Clasificadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "generales" Then Generales.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)  'Presupuesto
                If vNombOpcMenu = "unidadesejecutoras" Then UnidadesEjecutoras.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "entidades" Then Entidades.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "tipotramite" Then TipoTramite.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "cbeneficiarios" Then CBeneficiarios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "departamentosbolivia" Then DepartamentosBolivia.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "provinciasdepartamentos" Then ProvinciasDepartamentos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "tiposerrores" Then TiposErrores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "cpresupuesto" Then CPresupuesto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)  'Presupuesto
                If vNombOpcMenu = "partidasgasto" Then PartidasGasto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "economicosgasto" Then economicosgasto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "RelacionadorGastoEco" Then RelacionadorGastoEco.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "presupuesto" Then Presupuesto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)  'Presupuesto
                If vNombOpcMenu = "fuentesfinanciamiento" Then FuentesFinanciamiento.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "organismosfinanciadores" Then OrganismosFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "convenios" Then Convenios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "categoriafinanciadores" Then CategoriaFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "estructuraprogramatica" Then EstructuraProgramatica.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "sisin" Then Sisin.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "deducciones" Then Deducciones.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "relacionadordonacionesorganismos" Then RelacionadorDonacionesOrganismos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'             If vNombOpcMenu = "cpresupuesto" Then CPresupuesto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)  'Presupuesto
'                If vNombOpcMenu = "unidadesejecutoras" Then UnidadesEjecutoras.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "partidasgasto" Then PartidasGasto.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "estructuraprogramatica" Then EstructuraProgramatica.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "fuentesfinanciamiento" Then FuentesFinanciamiento.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "organismosfinanciadores" Then OrganismosFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "convenios" Then Convenios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "categoriafinanciadores" Then CategoriaFinanciadores.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "distribucionporcentual" Then DistribucionPorcentual.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "entidades" Then Entidades.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "cbeneficiarios" Then CBeneficiarios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
'                If vNombOpcMenu = "tipotramite" Then TipoTramite.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "ctesoreria" Then CTesoreria.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "cbancos" Then CBancos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "cctabancarias" Then CCtaBancarias.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "formaspago" Then FormasPago.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "ingresos" Then Ingresos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "rubros" Then Rubros.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "económicosrecursos" Then EconómicosRecursos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "relacionadorrubroeco" Then RelacionadorRubroEco.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "contabilidad2" Then Contabilidad2.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "plancuentas" Then PlanCuentas.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "relacionadorcuentapartidas" Then RelacionadorCuentaPartidas.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "relacionadoringresoscuentas" Then RelacionadorIngresosCuentas.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "depreciaciones" Then Depreciaciones.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "claseauxiliares" Then ClaseAuxiliares.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "económicosrecursos" Then EconómicosRecursos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "inversiones" Then Inversiones.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "administrativos" Then Administrativos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Tesoreria
                If vNombOpcMenu = "adquisiciones" Then Adquisiciones.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "contrataciones" Then Contrataciones.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "almacenes2" Then Almacenes2.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "recursoshumanos2" Then RecursosHumanos2.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Tesoreria
    
    If vNombOpcMenu = "mesaentrada" Then MesaEntrada.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)     'Contabilidad
            If vNombOpcMenu = "registrosolicitudes" Then RegistroSolicitudes.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof01" Then FormularioF01.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof02" Then FormularioF02.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof03" Then FormularioF03.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof04" Then FormularioF04.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof05" Then FormularioF05.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof06" Then FormularioF06.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "formulariof07" Then FormularioF07.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "copiarestauracion" Then CopiaRestauracion.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "progconadq" Then ProgConAdq.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
    
    If vNombOpcMenu = "procesos" Then Procesos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)          'Egresos
        If vNombOpcMenu = "compromiso" Then Compromiso.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Ejecucion Presupuestaria
        If vNombOpcMenu = "ejecucion" Then Ejecucion.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)     'Reportes de Ejecucion
            If vNombOpcMenu = "ejecucionorganismo" Then EjecucionOrganismo.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "ejecucioncomprobante" Then EjecucionComprobante.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "ejecucioncompromiso" Then EjecucionCompromiso.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "ejecucióndevengado" Then EjecuciónDevengado.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "ejecuciónpagado" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
    
    If vNombOpcMenu = "tesoreria" Then Tesoreria.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
        If vNombOpcMenu = "pp" Then pp.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Pagos pendientes
        If vNombOpcMenu = "pagosefectuados" Then PagosEfectuados.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Activacion de Cheques
        If vNombOpcMenu = "cuentasbancarias2" Then CuentasBancarias2.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Desactivacion de cheques
        If vNombOpcMenu = "manejocheques" Then ManejoCheques.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Impresion de Comprobantes de Trasnferencia
        If vNombOpcMenu = "ic" Then ic.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Impresion de Comprobantes de Pago
        If vNombOpcMenu = "ct" Then ct.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Impresion de Comprobantes de Pago
            If vNombOpcMenu = "mnuimpcheques" Then MnuImpCheques.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Impresion de Cheques
            If vNombOpcMenu = "cuentasbancarias" Then CuentasBancarias.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Impresion de Cheques
        
    If vNombOpcMenu = "ccontabilidad" Then CContabilidad.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)     'Contabilidad
            If vNombOpcMenu = "comprobantes" Then Comprobantes.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
            If vNombOpcMenu = "reportesc" Then ReportesC.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "libromayor" Then LibroMayor.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "libromayorauxiliar" Then LibroMayorAuxiliar.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "balancegeneral" Then BalanceGeneral.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "balancesumassaldos" Then BalanceSumasSaldos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "estadoresultados" Then EstadoResultados.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                
    If vNombOpcMenu = "mnuingresos" Then MnuIngresos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)          'Egresos
        If vNombOpcMenu = "ejecucionpresupuestaria" Then EjecucionPresupuestaria.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Ejecucion Presupuestaria
        If vNombOpcMenu = "reportesingresos" Then ReportesIngresos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)     'Reportes de Ejecucion
    
    If vNombOpcMenu = "administracion" Then Administracion.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
        If vNombOpcMenu = "adquisicionbienes" Then AdquisicionBienes.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "comprasdirectas" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "licitacionesnacionales" Then EjecuciónPagado.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "licitacionesinternacionales" Then LicitacionesInternacionales.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
        If vNombOpcMenu = "contratacionservicios" Then ContratacionServicios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "consultoresindividuales" Then ConsultoresIndividuales.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "empresasconsultoras" Then EmpresasConsultoras.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
        If vNombOpcMenu = "almacenes" Then Almacenes.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "ingresosa" Then IngresosA.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
                If vNombOpcMenu = "salidasa" Then SalidasA.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)

    If vNombOpcMenu = "recursoshumanos" Then RecursosHumanos.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)     'Contabilidad
        If vNombOpcMenu = "administracionpersonal" Then AdministracionPersonal.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
        If vNombOpcMenu = "controlpersonal" Then ControlPersonal.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
        If vNombOpcMenu = "capacitacionpersonal" Then CapacitacionPersonal.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
        If vNombOpcMenu = "evaluaciondesempeño" Then EvaluacionDesempeño.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)
    
    If vNombOpcMenu = "informaciongerencial" Then InformacionGerencial.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)     'Contabilidad

    If vNombOpcMenu = "mnuadmisistema" Then mnuAdmiSistema.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)     'Administracion del sistema
        If vNombOpcMenu = "mnucambiarclave" Then mnuCambiarClave.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Cambiar clave
        If vNombOpcMenu = "mnuusuarios" Then mnuUsuarios.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False)    'Definicion de Usuarios
        If vNombOpcMenu = "mnunivelacceso" Then mnuNivelAcceso.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Cambiar clave
        If vNombOpcMenu = "mnuprivacceso" Then mnuPrivAcceso.Enabled = IIf(rsNivelAcceso!Habilitado = "Si", True, False) 'Privilegios de Operación
    rsNivelAcceso.MoveNext
    Wend
    rsNivelAcceso.MoveFirst
End If
End Sub

'Private Sub TipoTramite_Click() 'clasifica
'  frmFormulario.Show
'End Sub

'Private Sub UnidadesEjecutoras_Click() 'clasifica
''  Dim e As New Unidad
''  e.Show
'  Unidad.Show
'End Sub
