VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frm_ro_SelecBenLiquida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Beneficiarios para la liquidación"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   682
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBenSegunLiq 
      Caption         =   "Adicionar Beneficiarios y Montos de una liquidación anterior:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Adiciona beneficiario al pago..."
      Top             =   5880
      Width           =   5325
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ro_SelecBenLiquida.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ro_SelecBenLiquida.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ro_SelecBenLiquida.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ro_SelecBenLiquida.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ro_SelecBenLiquida.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ro_SelecBenLiquida.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ro_SelecBenLiquida.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ro_SelecBenLiquida.frx":0292
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo cboNroLiq 
      Height          =   315
      Left            =   5520
      TabIndex        =   2
      Top             =   6000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   " "
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TrueOleDBGrid60.TDBGrid GrdDetalle 
      Height          =   4995
      Left            =   120
      OleObjectBlob   =   "frm_ro_SelecBenLiquida.frx":02F0
      TabIndex        =   4
      Top             =   480
      Width           =   9990
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   741
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tool_OrdenarAZ"
            Object.ToolTipText     =   "Ordenar ascendentemente"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tool_OrdenarZA"
            Object.ToolTipText     =   "Ordenar descendentemente"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tool_Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tool_Filtrar"
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tool_Elegir"
            Object.ToolTipText     =   "Elegir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tool_Refrescar"
            Object.ToolTipText     =   "Refrescar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tool_Volver"
            Object.ToolTipText     =   "Volver"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblContar 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro. de registros:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   9990
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuSalir 
         Caption         =   "Volver"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuRefrescar 
         Caption         =   "Refrescar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuOrdenar 
         Caption         =   "Ordenar"
         Begin VB.Menu mnuAscendente 
            Caption         =   "Ascendente"
         End
         Begin VB.Menu mnuDescendente 
            Caption         =   "Descendente"
         End
      End
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFiltrar 
         Caption         =   "Filtrar..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuElegir 
         Caption         =   "Elegir"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
   End
End
Attribute VB_Name = "frm_ro_SelecBenLiquida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTabla As ADODB.Recordset ' recordset de la tabla que se esta procesando
Dim filtro As String ' usado para guardar la cadena de filtro
Dim Gestion As String
Dim CodUnidad As String
Dim CodGrupo As Integer
Dim NroLiq As Integer
Dim NroReg As Integer

Private Sub cboNroLiq_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case 46 ' si presiono suprimir
        cboNroLiq.BoundText = ""
      Case 13 ' si presiono enter
        SendKeys "{Tab}"
      Case 27 ' si presiono escape
        pl_OpcionesGenericas ("Tool_Cancelar")
        
    End Select
End Sub

Private Sub cmdBenSegunLiq_Click()
    Dim SQLs As String
    Dim fechax As String
    Dim horax As String
    Dim TipoCambioUS As Currency
    
    On Error GoTo EtiqError
    'JQ QR
    'DE.dbo_edGetProcessDateTime fechax, horax
    '' OBTIENE EL TIPO DE CAMBIO DEL DOLAR DEL DIA
    SQLs = "select cambio_oficial from ac_tipo_cambio where fecha_cambio = '" & fechax & "'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        TipoCambioUS = IIf(IsNull(rstTemp!cambio_oficial), 0, rstTemp!cambio_oficial)
    End If
    
    If Len(Trim(cboNroLiq.BoundText)) > 0 Then
        SQLs = "select * from ao_pagos_cronograma_detalle "
        SQLs = SQLs & "WHERE ges_gestion = '" & Gestion & "' AND "
        SQLs = SQLs & "codigo_unidad = '" & CodUnidad & "' AND "
        SQLs = SQLs & "codigo_grupo = " & CodGrupo & " AND "
        SQLs = SQLs & "numero_pago = " & cboNroLiq.BoundText & " AND "
        SQLs = SQLs & "codigo_beneficiario NOT IN(SELECT CODIGO_BENEFICIARIO FROM ao_pagos_cronograma_detalle WHERE ges_gestion = '" & Gestion & "' AND codigo_unidad = '" & CodUnidad & "' AND codigo_grupo = " & CodGrupo & " AND numero_pago = " & NroLiq & ")"
    
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        If rstTemp.RecordCount > 0 Then
            Screen.MousePointer = vbHourglass
            SQLs = "Se adicionarán los beneficiarios de la liquidación [" & cboNroLiq.BoundText & "]." & Chr(13)
            SQLs = SQLs & "Desea continuar con el proceso?"
            If vbYes = MsgBox(SQLs, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                
                ' actualiza la tabla ao_pagos_cronograma_detalle
                rstTemp.MoveFirst
                Select Case rstTemp!tipo_moneda & ""
                  Case "Bs"
                    While Not rstTemp.EOF
                        'JQ QR
                        'DE.dbo_ap_PagosGrabaPagoBenef Gestion, CodUnidad, CodGrupo, NroLiq, NroReg, rstTemp!codigo_beneficiario, "Bs", TipoCambioUS, rstTemp!monto_bs / TipoCambioUS, rstTemp!monto_bs, IIf(IsNull(rstTemp!numero_consultoriaHist), "", rstTemp!numero_consultoriaHist), IIf(IsNull(rstTemp!fte_financiamientoHist), "", rstTemp!fte_financiamientoHist), GlUsuario
'                        frm_ro_LiquidaMain.lblEstadoBeneficiario.Caption = rstTemp!codigo_beneficiario ' se guarda codigo beneficiairo
                        rstTemp.MoveNext
                    Wend
                  Case "$US"
                    While Not rstTemp.EOF
                        'JQ QR
                        'DE.dbo_ap_PagosGrabaPagoBenef Gestion, CodUnidad, CodGrupo, NroLiq, NroReg, rstTemp!codigo_beneficiario, "$US", TipoCambioUS, rstTemp!monto_us, rstTemp!monto_us * TipoCambioUS, IIf(IsNull(rstTemp!numero_consultoriaHist), "", rstTemp!numero_consultoriaHist), IIf(IsNull(rstTemp!fte_financiamientoHist), "", rstTemp!fte_financiamientoHist), GlUsuario
'                        frm_ro_LiquidaMain.lblEstadoBeneficiario.Caption = rstTemp!codigo_beneficiario ' se guarda codigo beneficiairo
                        rstTemp.MoveNext
                    Wend
                  Case Else
                    While Not rstTemp.EOF
                        'JQ QR
                        'DE.dbo_ap_PagosGrabaPagoBenef Gestion, CodUnidad, CodGrupo, NroLiq, NroReg, rstTemp!codigo_beneficiario, IIf(IsNull(rstTemp!tipo_moneda), "", rstTemp!tipo_moneda), IIf(IsNull(rstTemp!tc_us), 0, rstTemp!tc_us), rstTemp!monto_us, rstTemp!monto_bs, IIf(IsNull(rstTemp!numero_consultoriaHist), "", rstTemp!numero_consultoriaHist), IIf(IsNull(rstTemp!fte_financiamientoHist), "", rstTemp!fte_financiamientoHist), GlUsuario
'                        frm_ro_LiquidaMain.lblEstadoBeneficiario.Caption = rstTemp!codigo_beneficiario ' se guarda codigo beneficiairo
                        rstTemp.MoveNext
                    Wend
                End Select
                
                Screen.MousePointer = vbDefault
                Unload Me
              Else
                GrdDetalle.SetFocus
            End If
          Else
            MsgBox "No existen beneficiarios para ser adicionados en la liquidación seleccionada.", vbInformation, "Aviso"
        End If
      Else
        MsgBox "No selecciono número de liquidación.", vbInformation, "Aviso"
        cboNroLiq.SetFocus
    End If

    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    ' si se produjo otro tipo de error
    MsgBox "Error: Se produjo un error." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub Form_Load()
    Dim sql As String ' para la consulta de peronal
    
    On Error GoTo EtiqError
    
'    Gestion = frm_ro_LiquidaMain.lblGestion.Caption
    Gestion = 2008
    CodUnidad = frm_ro_LiquidaMain.lblCodUniSol.Caption
    CodGrupo = Val(frm_ro_LiquidaMain.lblCodGrupo.Caption)
    NroLiq = Val(frm_ro_LiquidaMain.grdBeneficiario.Tag)
    NroReg = Val(frm_ro_LiquidaMain.lblEstadoBeneficiario.Tag)
    
    Call pl_LlenaCombos
    
    Call pl_RefrescaList
       
    Screen.MousePointer = vbDefault

    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    ' si se produjo otro tipo de error
    MsgBox "Error: Se produjo un error." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    
End Sub

Private Sub pl_LlenaCombos()
    ' llena combo de liquidaciones anteriores
    Dim SQLs As String
    On Error GoTo EtiqError

    SQLs = "SELECT ao_pagos_cronograma.numero_pago, cast(ao_pagos_cronograma.numero_pago as varchar(5)) + ' - ' + ao_pagos_cronograma.concepto AS concepto "
    SQLs = SQLs & "FROM ao_pagos_grupos INNER JOIN ao_pagos_cronograma ON ao_pagos_grupos.ges_gestion = ao_pagos_cronograma.ges_gestion AND "
    SQLs = SQLs & "ao_pagos_grupos.codigo_unidad = ao_pagos_cronograma.codigo_unidad And ao_pagos_grupos.codigo_grupo = ao_pagos_cronograma.codigo_grupo "
    SQLs = SQLs & " WHERE   ao_pagos_cronograma.ges_gestion = '" & Gestion & "' AND "
    SQLs = SQLs & "ao_pagos_cronograma.codigo_unidad = '" & CodUnidad & "' AND "
    SQLs = SQLs & "ao_pagos_cronograma.codigo_grupo = " & CodGrupo & " AND "
    SQLs = SQLs & "ao_pagos_cronograma.numero_pago < " & NroLiq & " AND "
    SQLs = SQLs & "ao_pagos_cronograma.estado_devengado <>'E' "
    SQLs = SQLs & "ORDER BY ao_pagos_cronograma.numero_pago DESC"

    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        Set cboNroLiq.RowSource = rstTemp
        cboNroLiq.BoundColumn = "numero_pago"
        cboNroLiq.ListField = "concepto"
    End If

    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    ' si se produjo otro tipo de error
    MsgBox "Error: Se produjo un error." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_RefrescaList()
    ' cargamos la informacion de la tabla
    'JQ QR
    'De.dbo_ap_PagosListaBenPSelec glProceso, Gestion, CodUnidad, CodGrupo, NroLiq, NroReg
    'Set rsTabla = De.rsdbo_ap_PagosListaBenPSelec.Clone
    'De.rsdbo_ap_PagosListaBenPSelec.Close
    
    ' se asocia el recordset al grid
    Set GrdDetalle.DataSource = rsTabla
    Call pl_PersonalizaGrid
    lblContar.Caption = "Nro. de beneficiarios: " & rsTabla.RecordCount

End Sub

Private Sub pl_PersonalizaGrid()
    'TITULO:                Procedimiento pl_PersonalizaGrid
    'PROPOSITO:             Personalizar los captions, anchos, etc. del grid
    'EJEMPLO DE LLAMADA:    call pl_PersonalizaGrid
        
    ' define ancho de columnas
    ' define titulos de encabezado de cada columna
    GrdDetalle.Columns(0).Width = 90
    GrdDetalle.Columns(0).Caption = "Paterno"
    GrdDetalle.Columns(1).Width = 90
    GrdDetalle.Columns(1).Caption = "materno"
    GrdDetalle.Columns(2).Width = 90
    GrdDetalle.Columns(2).Caption = "Nombre(s)"
    GrdDetalle.Columns(3).Width = 50
    GrdDetalle.Columns(3).Caption = "Tipo doc."
    GrdDetalle.Columns(4).Width = 90
    GrdDetalle.Columns(4).Caption = "Nro. Doc.."
    GrdDetalle.Columns(5).Width = 40
    GrdDetalle.Columns(5).Caption = "Emisión"
    GrdDetalle.Columns(6).Width = 70
    GrdDetalle.Columns(6).Caption = "Comprometido"
    GrdDetalle.Columns(7).Width = 70
    GrdDetalle.Columns(7).Caption = "Aprobo presupuestos"
    GrdDetalle.Columns(8).Width = 50
    GrdDetalle.Columns(8).Caption = "Tipo moneda"
    GrdDetalle.Columns(9).Width = 50
    GrdDetalle.Columns(9).Caption = "TC $US"
      
End Sub

Private Sub mnuAscendente_Click()
    pl_OpcionesGenericas ("Tool_OrdenarAZ")
End Sub

Private Sub mnuBuscar_Click()
    pl_OpcionesGenericas ("Tool_Buscar")
End Sub

Private Sub mnuDescendente_Click()
    pl_OpcionesGenericas ("Tool_OrdenarZA")
End Sub

Private Sub mnuElegir_Click()
    pl_OpcionesGenericas ("Tool_Elegir")
End Sub

Private Sub mnuFiltrar_Click()
    pl_OpcionesGenericas ("Tool_Filtrar")
End Sub

Private Sub mnuRefrescar_Click()
    pl_OpcionesGenericas ("Tool_Refrescar")
End Sub

Private Sub mnuSalir_Click()
    pl_OpcionesGenericas ("Tool_Volver")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    ' ejecuta una acción del toolbar
    Call pl_OpcionesGenericas(Button.Key)

End Sub

Private Sub pl_OpcionesGenericas(TipoOpcion As String)
    
    'TITULO:                Procedimiento pl_OpcionesGenericas
    'PROPOSITO:             Ejecuta una opcion del toolbar o del menu
    'EJEMPLO DE LLAMADA:    call pl_OpcionesGenericas(TipoOpcion)
    'ENTRADAS:              TipoOpcion = Opción a elegir (Grabar,Borrar, etc)
                            ' Realiza una acción según TipoOpcion
    
    Dim swGuardar As Boolean ' usado para saber si efectivamente se almaceno los datos en la base
    Dim RegPuntero As String ' usada para guardar el código de registro para poder apuntar el el registro seleccionado luego de un refresh
    
    On Error GoTo EtiqError
    
    Select Case TipoOpcion
      
      Case "Tool_OrdenarAZ" ' ordena ascendentemente
        If rsTabla.RecordCount > 0 Then
            Call pg_OrdenaTdbGrid(GrdDetalle, rsTabla, True)
            Call pl_PersonalizaGrid
          Else
            MsgBox "No existen registros para ordenar.", vbInformation, "Aviso"
        End If
        
      Case "Tool_OrdenarZA" ' ordena descendentemente
        If rsTabla.RecordCount > 0 Then
            Call pg_OrdenaTdbGrid(GrdDetalle, rsTabla, False)
            Call pl_PersonalizaGrid
          Else
            MsgBox "No existen registros para ordenar.", vbInformation, "Aviso"
        End If
        
      Case "Tool_Buscar"
        If rsTabla.RecordCount > 0 Then
            Call pg_BuscaTdbGrid(GrdDetalle, rsTabla, GrdDetalle.Columns(GrdDetalle.Col).DataField)
            Call pl_PersonalizaGrid
          Else
            MsgBox "No existen registros para búscar.", vbInformation, "Aviso"
        End If
        
      Case "Tool_Filtrar"
        
        If rsTabla.RecordCount > 0 Then
            Call pl_FiltraDatos
            Call pl_PersonalizaGrid
          Else
            MsgBox "No existen registros para ser filtrados.", vbInformation, "Aviso"
        End If

      Case "Tool_Elegir"
        
        If rsTabla.RecordCount > 0 Then
            If fl_VerificaSeleccion Then
                
                ' actualiza la tabla ao_pagos_cronograma_detalle
                'JQ QR
                'De.dbo_ap_PagosGrabaPagoBenef Gestion, CodUnidad, CodGrupo, NroLiq, NroReg, rsTabla!codigo_beneficiario, rsTabla!tipo_moneda, IIf(IsNull(rsTabla!tc_us), 0, rsTabla!tc_us), 0, 0, "", "", GlUsuario
                frm_ro_LiquidaMain.lblEstadoBeneficiario.Caption = rsTabla!codigo_beneficiario ' se guarda codigo beneficiairo
                
                Unload Me
            End If
          Else
            MsgBox "No existen registros para elegir.", vbInformation, "Aviso"
        End If
        
      Case "Tool_Refrescar"
        
        Screen.MousePointer = vbHourglass
        If Len(filtro) > 0 Then
            filtro = ""
            rsTabla.Filter = "" ' se vacia el filtro
          Else
            Call pl_RefrescaList ' se refresca el recordset para mostrar los datos originales
        End If
        
        Call pl_PersonalizaGrid
        GrdDetalle.Refresh

        If Len(RegPuntero) > 0 Then ' se ubica en el registro actual si existe
            rsTabla.Find rsTabla.Fields(0).Name & "='" & RegPuntero & "'" ' el puntero de registro se ubica en la posicion guardada
        End If
        lblContar.Caption = "Nro. de proyectos aprobados: " & rsTabla.RecordCount
        Screen.MousePointer = vbDefault

      Case "Tool_Volver"
        
        Unload Me
        
    End Select
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub
    
EtiqError:
    Screen.MousePointer = vbDefault
    ' si se produjo otro tipo de error
    MsgBox "Error: Se produjo un error." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    
End Sub

Private Sub pl_FiltraDatos()
    'PROPÓSITO      : Realiza la filtración de una especificación sobre la columna de la celda activa
    
    Dim CadFiltro As String ' usada para almacenar la cedena de filtración
    Dim micriterio As String ' critetio de filtración
    Dim CampoAct As Integer ' nombre del campo activo
    Dim ColFiltro As String ' usada para almacenar la columna activa por el cual se filtrará
    Dim a As Integer ' usada para ver el formato de cadena a filtrar
    
   ' If rsTabla.RecordCount = 0 Then Exit Sub
    micriterio = "Digite " & LCase(GrdDetalle.Columns(GrdDetalle.Col).Caption) & " a filtrar"
    CadFiltro = pg_QuitaEspBlanco(UCase(InputBox(micriterio, "Filtración")))
    ' verificamos que la cadena sea del tipo *a* donde a representa cualquier secuencia de caracteres
    Select Case Len(CadFiltro)
      Case Is >= 3
        If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) <> "*" Then
            ' completamos la cadena al tipo *a*
            CadFiltro = CadFiltro & "*"
          Else
            ' es del tipo a* o a que son cadenas validas
            'CadFiltro = "*" & CadFiltro
        End If
      Case 2
        If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) <> "*" Then
            CadFiltro = CadFiltro & "*"
          Else
            If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) = "*" Then
                ' si ambos son *
                CadFiltro = ""
              Else
                ' es del tipo a* o a que son cadenas validas
                'CadFiltro = "*" & CadFiltro
            End If
        End If
      Case 1
        If CadFiltro = "*" Then CadFiltro = ""
    End Select
    
    On Error GoTo EtiqError
    If Len(CadFiltro) > 0 Then ' si introdujo una cadena a filtrar
        CampoAct = GrdDetalle.Col
        ColFiltro = GrdDetalle.Columns(GrdDetalle.Col).DataField
        
        ' verificamos si la longitud coincide con el tamaño del campo
        If Len(CadFiltro) <= rsTabla.Fields(ColFiltro).DefinedSize Or rsTabla.Fields(ColFiltro).Type = 3 Then
            If rsTabla.Filter = 0 Then ' es la primera filtración
                rsTabla.Filter = ColFiltro & " like " & Chr(39) & CadFiltro & Chr(39)
                filtro = GrdDetalle.Columns(CampoAct).Caption & " -> " & Chr(39) & CadFiltro & Chr(39) ' concatenamos la cadena de filtración
              Else
                rsTabla.Filter = rsTabla.Filter & " AND " & ColFiltro & " like " & Chr(39) & CadFiltro & Chr(39)
                filtro = filtro & ", " & GrdDetalle.Columns(CampoAct).Caption & " -> " & Chr(39) & CadFiltro & Chr(39)  ' concatenamos la cadena de filtración
            End If
            lblContar.Caption = "Nro. de proyectos: " & rsTabla.RecordCount & " Filtro ( " & filtro & " )"
                        
            If rsTabla.RecordCount = 0 Then ' no se encontraron coincidencias
                GrdDetalle.ReOpen
                MsgBox "No se encontró ninguna coincidencia con " & filtro, vbInformation, "Información"
            
            End If
            GrdDetalle.SetFocus
            
          Else ' la longitud de la cadena a filtrar es mayor a la longitud del campo
            MsgBox "La longitud de la cadena a filtrar -> " & CadFiltro & " es mayor a la longitud del permitido por " & GrdDetalle.Columns(CampoAct).Caption, vbInformation, "Información"
        End If
        GrdDetalle.SetFocus
        
      Else ' solo tiene el foco
        GrdDetalle.SetFocus
    End If
    
    On Error GoTo 0 ' desactiva el manejador de errores
    Exit Sub
    
EtiqError:
    Select Case Err.Number
      Case -2147352571
        MsgBox "Error: No se pueden filtrar los datos, los tipos no coinciden." & Chr(13) & Chr(13) & "No se realizo la filtración de datos." & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description, vbCritical, "Error"
      Case Else ' si se produjo otro tipo de error
        MsgBox "Error: No se realizo la filtración de datos." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    End Select

End Sub

Private Function fl_VerificaSeleccion() As Boolean
    'TITULO:                Función fl_VerificaSeleccion
    'PROPOSITO:             Verifica los datos para el registro de grupo de liquidacion
    'EJEMPLO DE LLAMADA:    fl_VerificaSeleccion
    Dim SQLs As String
    
    fl_VerificaSeleccion = True ' asuminos que se cuenta con los datos mnimos para grabar

'' queda pendiente esta verificación
''    ' verificamos si tiene comprobante de comnpromiso
''    SQLs = "SELECT * from ac_ben_comprdeven WHERE tipocomprobante = 'COM' AND aproboTesoreria = 'S' AND codigo_beneficiario =" & rsTabla!codigo_beneficiario
''    Set rstTemp = New ADODB.Recordset
''    rstTemp.Open SQLs, G_DBConnection, adOpenStatic, adLockReadOnly
''    If rstTemp.RecordCount = 0 Then
''        MsgBox "El beneficiario no tiene procesado comprobante de compromiso." & Chr(13) & "Corrija el error e intente seleccionar nuevamente.", vbInformation, "Aviso"
''        fl_VerificaSeleccion = False
''        GrdDetalle.SetFocus
''        Exit Function
''    End If
''
    
''    ' verificamos si lo tipos de moneda base coinciden
''    If Len(Trim(af_LiquidaConsultoria.cboTipoMoneda.BoundText)) > 0 And af_LiquidaConsultoria.cboTipoMoneda.BoundText <> rsTabla!tipo_moneda Then
''        MsgBox "El tipo de mondeda base [" & rsTabla!tipo_moneda & "] es distindo al tipo de moneda base del número de liquidación [" & af_LiquidaConsultoria.cboTipoMoneda.BoundText & "]." & Chr(13) & "Corrija el error e intente seleccionar nuevamente.", vbInformation, "Aviso"
''        fl_VerificaSeleccion = False
''        GrdDetalle.SetFocus
''        Exit Function
''    End If
    
End Function

