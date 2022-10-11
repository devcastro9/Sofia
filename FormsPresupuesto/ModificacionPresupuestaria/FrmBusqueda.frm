VERSION 5.00
Begin VB.Form FrmBusqueda 
   Caption         =   "  Busqueda"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6480
   Icon            =   "FrmBusqueda.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2460
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraBusqueda 
      BackColor       =   &H00808080&
      Height          =   2445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      Begin VB.CommandButton cmdCancelarBusqueda 
         Caption         =   "&Terminar"
         DownPicture     =   "FrmBusqueda.frx":030A
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4440
         Picture         =   "FrmBusqueda.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Frame FraCriterios 
         Height          =   1065
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6225
         Begin VB.TextBox TxtValor 
            Height          =   285
            Left            =   3765
            TabIndex        =   6
            Top             =   645
            Width           =   2265
         End
         Begin VB.ComboBox CmbOperador 
            Height          =   315
            ItemData        =   "FrmBusqueda.frx":0B8E
            Left            =   2565
            List            =   "FrmBusqueda.frx":0BA1
            TabIndex        =   5
            Top             =   630
            Width           =   1065
         End
         Begin VB.ComboBox CmbCampo 
            Height          =   315
            Left            =   165
            TabIndex        =   4
            Top             =   630
            Width           =   2295
         End
         Begin VB.Label LblValor 
            Alignment       =   2  'Center
            Caption         =   "VALOR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   4440
            TabIndex        =   9
            Top             =   255
            Width           =   675
         End
         Begin VB.Label LblOperador 
            Alignment       =   2  'Center
            Caption         =   "OPERADOR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2445
            TabIndex        =   8
            Top             =   255
            Width           =   1365
         End
         Begin VB.Label LblCampo 
            Alignment       =   2  'Center
            Caption         =   "CAMPO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   600
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton CmdEjecutarBusqueda 
         Caption         =   "&Ejecutar"
         DownPicture     =   "FrmBusqueda.frx":0BB8
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         Picture         =   "FrmBusqueda.frx":0FFA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1020
      End
      Begin VB.CommandButton cmdRefrescarBusqueda 
         Caption         =   "&Refrescar"
         DownPicture     =   "FrmBusqueda.frx":143C
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2760
         Picture         =   "FrmBusqueda.frx":187E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   1020
      End
   End
End
Attribute VB_Name = "FrmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub cmdCancelarBusqueda_Click()
    Unload FrmBusqueda
End Sub

Private Sub CmdEjecutarBusqueda_Click()
  Dim cadena_busqueda As String
  cadena_busqueda = ""
  Select Case varbusca
    Case "FOR"          ' Formulacion
        If CmbCampo = "fte_codigo" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "org_codigo" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "pro_programa" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "pro_proyecto" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + CmbOperador + " '" + TxtValor + "' "
        End If
        If CmbCampo = "pro_actividad" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "par_codigo" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "fgs_formulado" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + CmbOperador + " " + TxtValor + " "
        End If
        If CmbCampo = "fgs_adiciones" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + CmbOperador + " " + TxtValor + " "
        End If
        If CmbCampo = "fgs_modificaciones" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + CmbOperador + " " + TxtValor + " "
        End If
        If CmbCampo = "fecha_formulacion" Then
            cadena_busqueda = "fv_formulacion_gasto." + CmbCampo.Text + " = " + "#" + TxtValor + "#"
        End If
            
        If cadena_busqueda <> "" Then
            parametro = cadena_busqueda + " and " + "fv_formulacion_gasto.pro_programa" + " = " + "'10'"
            If OriDes = "F" Then
                
                Call FrmFormulacion.abrir_formulacion                    'Abrir fv_formulacion_gasto
                
                Call FrmFormulacion.Totales
                FrmFormulacion.lblFormulado = Format(montoTotal, "###,###,##0")
                FrmFormulacion.lblAdiciones = Format(montoTotalA, "###,###,##0")
                FrmFormulacion.lblModificaciones = Format(montoTotalM, "###,###,##0")
                FrmFormulacion.lblVigente = Format((montoTotal + montoTotalA + montoTotalM), "###,###,##0")
            End If
            If OriDes = "O" Then
                parametro = cadena_busqueda + " and " + "fv_formulacion_gasto.fgs_modificaciones" + " <= " + "'0'" + " and " + "left(fv_formulacion_gasto.par_codigo,1)" + " <> " + "'1'"
                Call FrmOrigenDestino.abrir_formulacionO                    'Abrir fv_formulacion_gasto ORIGEN
            End If
            If OriDes = "D" Then
                parametro = cadena_busqueda + " and " + "fv_formulacion_gasto.fgs_modificaciones" + " >= " + "'0'"
                Call FrmOrigenDestino.abrir_formulacionD                    'Abrir fv_formulacion_gasto DESTINO
            End If

        Else
            MsgBox "El VALOR no ha sido encontrado, vuelva a intentar . . ."
        End If
    
    Case "TRF"          ' Traspasos
        If CmbCampo = "fte_codigo" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "org_codigo" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "pro_programa" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "pro_proyecto" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + CmbOperador + " '" + TxtValor + "' "
        End If
        If CmbCampo = "pro_actividad" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "par_codigo" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "trn_monto_origen" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + CmbOperador + " " + TxtValor + " "
        End If
        If CmbCampo = "fte_codigo_des" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "org_codigo_des" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "pro_programa_des" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "pro_proyecto_des" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + CmbOperador + " '" + TxtValor + "' "
        End If
        If CmbCampo = "pro_actividad_des" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "par_codigo_des" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "trn_monto_destino" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + CmbOperador + " " + TxtValor + " "
        End If
        If CmbCampo = "resolucion" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " like " + "'%" + TxtValor + "%'"
        End If
        If CmbCampo = "fecha_transaccion" Then
            cadena_busqueda = "po_formulacion_trn." + CmbCampo.Text + " = " + "#" + TxtValor + "#"
        End If
            
        If cadena_busqueda <> "" Then
            
            parametro = cadena_busqueda + " and " + "po_formulacion_trn.tipo_transaccion" + " = " + "'T'" + " or " + "po_formulacion_trn.tipo_transaccion" + " = " + "'F'"
            Call FrmFormulacion.abrir_traspaso
            
        Else
            MsgBox "El VALOR no ha sido encontrado, vuelva a intentar . . ."
        End If
    
    Case Else
         MsgBox "seleccione otro parametro"
      
  End Select
End Sub

Private Sub cmdRefrescarBusqueda_Click()
  Select Case varbusca
    Case "FOR"       ' Formulacion
        parametro = "fv_formulacion_gasto.ges_gestion" + " = " + "'2004'"
        If OriDes = "F" Then
            Call FrmFormulacion.abrir_formulacion                    'Abrir fv_formulacion_gasto
    
            Call FrmFormulacion.Totales
            FrmFormulacion.lblFormulado = Format(montoTotal, "###,###,##0")
            FrmFormulacion.lblAdiciones = Format(montoTotalA, "###,###,##0")
            FrmFormulacion.lblModificaciones = Format(montoTotalM, "###,###,##0")
            FrmFormulacion.lblVigente = Format((montoTotal + montoTotalA + montoTotalM), "###,###,##0")
        End If
        If OriDes = "O" Then
            parametro = "fv_formulacion_gasto.fgs_modificaciones" + " <= " + "'0'" + " and " + "left(fv_formulacion_gasto.par_codigo,1)" + " <> " + "'1'"
            Call FrmOrigenDestino.abrir_formulacionO                    'Abrir fv_formulacion_gasto ORIGEN
        End If
        If OriDes = "D" Then
            parametro = "fv_formulacion_gasto.fgs_modificaciones" + " >= " + "'0'"
            Call FrmOrigenDestino.abrir_formulacionD                    'Abrir fv_formulacion_gasto DESTINO
        End If
    Case "TRF"       ' Traspasos
            parametro = "po_formulacion_trn.tipo_transaccion" + " = " + "'T'" + " or " + "po_formulacion_trn.tipo_transaccion" + " = " + "'F'"
            Call FrmFormulacion.abrir_traspaso
    
    Case Else
         MsgBox "seleccione otro parametro"
      
  End Select


End Sub
