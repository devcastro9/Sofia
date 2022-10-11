VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form mw_ventas_alcance 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   120
         ScaleHeight     =   705
         ScaleWidth      =   10815
         TabIndex        =   1
         Top             =   3600
         Width           =   10875
         Begin VB.CommandButton BtnModDetalle3 
            BackColor       =   &H00FFFFFF&
            Height          =   580
            Left            =   2760
            Picture         =   "mw_ventas_alcance.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Modifica Detalle Elegido"
            Top             =   80
            Width           =   1365
         End
         Begin VB.CommandButton BtnAddDetalle3 
            BackColor       =   &H00FFFFFF&
            Height          =   580
            Left            =   1320
            Picture         =   "mw_ventas_alcance.frx":0A15
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Adiciona Detalle"
            Top             =   80
            Width           =   1365
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Adjudica"
            Height          =   645
            Left            =   7080
            Picture         =   "mw_ventas_alcance.frx":12C5
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprime Nota de Venta"
            Top             =   0
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "mw_ventas_alcance.frx":2A47
         Height          =   720
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Visible         =   0   'False
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   1270
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cod.Proveedor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "adjudica_descripcion"
            Caption         =   "Denominación.Proveedor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "adjudica_monto_dol"
            Caption         =   "Precio.Total.USD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "adjudica_monto_bs"
            Caption         =   "Precio.Total_BOB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "adjudica_cantidad_total"
            Caption         =   "Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fecha_inicio_contrato"
            Caption         =   "Fecha.Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "fecha_fin_contrato"
            Caption         =   "Fecha.Finalizacion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Fecha_envio_proveedor"
            Caption         =   "Fecha.Entrega/Salida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3885.166
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1665.071
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DtgAlcanceC 
         Bindings        =   "mw_ventas_alcance.frx":2A62
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   10830
         _ExtentX        =   19103
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "venta_codigo"
            Caption         =   "Codigo.Venta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "solicitud_tipo"
            Caption         =   "Correl."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "unidad_codigo_tec"
            Caption         =   "Unidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "solicitud_tipo_descripcion"
            Caption         =   "Descripcion.del.Alcance"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "unidad_codigo_tec"
            Caption         =   "Unidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fecha_inicio_alcance"
            Caption         =   "Fecha.Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "fecha_fin_alcance"
            Caption         =   "Fecha.Fin"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "venta_tiempo_dias"
            Caption         =   "Dias.Calendario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4875.024
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "mw_ventas_alcance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_datos6 As New ADODB.Recordset    'ao_ventas_alcance

Private Sub Form_Load()
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "select * from ao_ventas_alcance where venta_codigo= " & nro_licitacion & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'rs_det2.Open "select * from to_cronograma_diario_final where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  AND estado_activo = 'APR' and hora_registro = '0' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Ado_datos6.Recordset = rs_datos6
    Set DtgAlcanceC.DataSource = rs_datos6.DataSource                'Ado_datos6.Recordset
    If rs_datos6.RecordCount > 0 Then
    'If Ado_datos6.Recordset.RecordCount > 0 Then
        DtgAlcanceC.Visible = True
        DtgAlcanceC.AllowUpdate = True
    Else
        DtgAlcanceC.Visible = False
    End If

End Sub
