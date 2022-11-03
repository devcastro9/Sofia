VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fw_orden_pago_comex 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Egresos - Cronograma de Pagos"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10260
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox TxtImporteBs 
      Alignment       =   2  'Center
      DataField       =   "pago_monto_bs"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,###,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      DataSource      =   "fw_compras_comex.Ado_detalle3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   43
      Text            =   "0"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10155
      TabIndex        =   36
      Top             =   0
      Width           =   10215
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         Picture         =   "fw_orden_pago_comex.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   39
         Top             =   60
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1755
         Picture         =   "fw_orden_pago_comex.frx":07D6
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   38
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORDENES DE PAGO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   330
         Left            =   5400
         TabIndex        =   37
         Top             =   120
         Width           =   2835
      End
   End
   Begin VB.Frame FrmCobros 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10095
      Begin VB.TextBox TxtDsctoDol 
         Alignment       =   2  'Center
         DataField       =   "pago_descuento_dol"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   53
         Text            =   "0"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CheckBox Chk_Ret 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Retención Asumida ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   3120
         TabIndex        =   46
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox TxtDsctoBs 
         Alignment       =   2  'Center
         DataField       =   "pago_descuento_bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   44
         Text            =   "0"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox TxtConcepto 
         CausesValidation=   0   'False
         DataField       =   "pago_descripcion"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   525
         Left            =   1815
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   4875
         Width           =   7995
      End
      Begin VB.TextBox TxtMontoBs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "pago_total_bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7125
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   3840
         Width           =   1545
      End
      Begin VB.TextBox TxtDscto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Text            =   "0"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtMontoDol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "pago_total_dol"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7125
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   4320
         Width           =   1545
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9525
         TabIndex        =   9
         Top             =   850
         Width           =   255
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9525
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtCobrador 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1425
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txt_respaldos 
         CausesValidation=   0   'False
         DataField       =   "pago_respaldos"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   1065
         Left            =   1800
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   5520
         Width           =   7995
      End
      Begin VB.CheckBox Chk_fac 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiene Factura ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   3000
         Width           =   1800
      End
      Begin VB.TextBox TxtImporteDol 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "pago_monto_dol"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Text            =   "0"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txt_factura 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "pago_nro_cmpbte_factura"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7725
         TabIndex        =   1
         Text            =   "0"
         Top             =   3000
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DTPFechaProg 
         DataField       =   "pago_fecha_prog"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   1995
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   118423553
         CurrentDate     =   44470
         MinDate         =   36526
      End
      Begin MSDataListLib.DataCombo dtc_codigo 
         Bindings        =   "fw_orden_pago_comex.frx":10C2
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   315
         Left            =   7800
         TabIndex        =   8
         Top             =   1425
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo txt_campo1 
         Bindings        =   "fw_orden_pago_comex.frx":10DC
         DataField       =   "beneficiario_codigo"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   315
         Left            =   7800
         TabIndex        =   14
         Top             =   840
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Txt_descripcion 
         Bindings        =   "fw_orden_pago_comex.frx":10F6
         DataField       =   "beneficiario_codigo"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   315
         Left            =   2400
         TabIndex        =   15
         Top             =   840
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc 
         Bindings        =   "fw_orden_pago_comex.frx":1127
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   315
         Left            =   2415
         TabIndex        =   16
         Top             =   1425
         Visible         =   0   'False
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPFechaPago 
         DataField       =   "pago_fecha_efectiva"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   285
         Left            =   7965
         TabIndex        =   17
         Top             =   1995
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   118423555
         CurrentDate     =   44477
         MaxDate         =   109939
         MinDate         =   36526
      End
      Begin MSDataListLib.DataCombo dtc_cuenta_cod 
         Bindings        =   "fw_orden_pago_comex.frx":1141
         DataField       =   "cta_codigo"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   315
         Left            =   7440
         TabIndex        =   40
         Top             =   2520
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_cuenta_desc 
         Bindings        =   "fw_orden_pago_comex.frx":115B
         DataField       =   "cta_codigo"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   315
         Left            =   2400
         TabIndex        =   41
         Top             =   2520
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "cta_descripcion"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_cuenta_moneda 
         Bindings        =   "fw_orden_pago_comex.frx":118C
         DataField       =   "cta_codigo"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         Height          =   315
         Left            =   1800
         TabIndex        =   54
         Top             =   2520
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "tipo_moneda"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808000&
         X1              =   1440
         X2              =   1440
         Y1              =   3420
         Y2              =   4760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808000&
         X1              =   -240
         X2              =   10080
         Y1              =   4760
         Y2              =   4760
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   6720
         TabIndex        =   52
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   4680
         TabIndex        =   51
         Top             =   4320
         Width           =   210
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Dolares -->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   1560
         TabIndex        =   50
         Top             =   4320
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "USD (Dolar)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   8760
         TabIndex        =   49
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         X1              =   -240
         X2              =   10080
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6750
         TabIndex        =   48
         Top             =   3840
         Width           =   150
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4680
         TabIndex        =   47
         Top             =   3840
         Width           =   210
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Dscto./Retención"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5000
         TabIndex        =   45
         Top             =   3540
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Bancaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   42
         Top             =   2520
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   35
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   225
         TabIndex        =   34
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lbl_obs 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto Cuota:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   225
         TabIndex        =   33
         Top             =   4920
         Width           =   1560
      End
      Begin VB.Label lbl_monto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto a Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3000
         TabIndex        =   32
         Top             =   3540
         Width           =   1440
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7605
         TabIndex        =   31
         Top             =   3540
         Width           =   675
      End
      Begin VB.Label Lbl_Cobrador 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable de CGI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   225
         TabIndex        =   30
         Top             =   1425
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbl_fechas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Programada de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   1995
         Width           =   2580
      End
      Begin VB.Label txtCodigo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "pago_codigo"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5205
         TabIndex        =   28
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nro.Orden Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   3600
         TabIndex        =   27
         Top             =   210
         Width           =   1470
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nro.Adjudica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   7305
         TabIndex        =   26
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Lbl_nombre_fac 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor/Beneficiario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   225
         TabIndex        =   25
         Top             =   840
         Width           =   2190
      End
      Begin VB.Label lbl_adjudica 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "adjudica_codigo"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8625
         TabIndex        =   24
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label lblccertif 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Bolivianos  -->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1530
         TabIndex        =   23
         Top             =   3840
         Width           =   1275
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFF80&
         X1              =   -240
         X2              =   10080
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "BOB (Bs)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   8760
         TabIndex        =   22
         Top             =   3840
         Width           =   825
      End
      Begin VB.Label lbl_plazo 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos de Respaldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   240
         TabIndex        =   21
         Top             =   5640
         Width           =   1440
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "compra_codigo"
         DataSource      =   "fw_compras_comex.Ado_detalle3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   20
         Top             =   195
         Width           =   1365
      End
      Begin VB.Label lblfechaCertif 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Orden de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5595
         TabIndex        =   19
         Top             =   2025
         Width           =   2280
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Cheque/Trf"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6240
         TabIndex        =   18
         Top             =   3000
         Width           =   1395
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   360
      Top             =   7560
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_clasif1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_clasif2 
      Height          =   330
      Left            =   2520
      Top             =   6240
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_clasif2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_clasif3 
      Height          =   330
      Left            =   2520
      Top             =   7560
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_clasif3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_clasif4 
      Height          =   330
      Left            =   360
      Top             =   5880
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_clasif4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_clasif5 
      Height          =   330
      Left            =   2520
      Top             =   5880
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_clasif5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_cuentas 
      Height          =   330
      Left            =   360
      Top             =   6600
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_cuentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "fw_orden_pago_comex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Para_Aceptado As String
Dim rs_clasif1 As New ADODB.Recordset
Dim rs_clasif2 As New ADODB.Recordset
Dim rs_clasif3 As New ADODB.Recordset
Dim rs_clasif4 As New ADODB.Recordset
Dim rs_clasif5 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_cuentas As New ADODB.Recordset

Dim VAR_OCUP, CAL_DOL, CAL_BS   As String

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
    Para_Aceptado = "N"
'    txtSW = "0"
    Unload Me
End Sub

Private Sub BtnGrabar_Click()
'acepta las modificaciones realizadas
If Valida Then
    Dim SQLS As String
    SQLS = ""
   'If txtSW = "ADD" Then
   If swnuevo = 1 Then
      'DB.Execute "Insert INTO ro_Beneficiario_Dependiente (beneficiario_codigo, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, cod_pariente, nomb_pariente, estado_codigo, beneficiario_denominacion, ocupacion_pariente) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & dtc_codigo1.Text & ", '" & dtc_desc1.Text & "', '" & txtEstado.Text & "', '" & nomb2 & "', '" & TxtOcupacion & "')"
      ''" & txtBenef.Caption & "',
       'DB.Execute "Insert INTO ao_solicitud_persona (ges_gestion, unidad_codigo, solicitud_codigo, benef_primer_apellido, benef_segundo_apellido, benef_nombres, benef_direccion_domicilio, benef_telefonos_ref, benef_codigo, puesto_codigo, ocup_codigo, munic_codigo, nivel_educ_codigo, observaciones, benef_fecha, estado_codigo, fecha_registro, usr_codigo) Values ('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
       '('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
      fw_compras_comex.Ado_detalle3.Recordset("ges_gestion") = glGestion
      fw_compras_comex.Ado_detalle3.Recordset("compra_codigo").Value = fw_compras_comex.Ado_datos.Recordset!compra_codigo
      fw_compras_comex.Ado_detalle3.Recordset("adjudica_codigo") = fw_compras_comex.Ado_detalle2.Recordset("adjudica_codigo")          ' fw_compras_comex.Ado_detalle2.Recordset!adjudica_codigo
      fw_compras_comex.Ado_detalle3.Recordset("pago_codigo") = fw_compras_comex.Ado_detalle3.Recordset.RecordCount
      fw_compras_comex.Ado_detalle3.Recordset("beneficiario_codigo") = fw_compras_comex.Ado_detalle2.Recordset("beneficiario_codigo")
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & dtc_codigo1.Text & ", nomb_pariente='" & dtc_desc1.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
   End If
    If Chk_fac.Value = 1 Then
        fw_compras_comex.Ado_detalle3.Recordset("pago_emite_factura").Value = "S"
    Else
        fw_compras_comex.Ado_detalle3.Recordset("pago_emite_factura").Value = "N"
    End If
    If Chk_Ret.Value = 1 Then
        fw_compras_comex.Ado_detalle3.Recordset("retension_asumida").Value = "S"
    Else
        fw_compras_comex.Ado_detalle3.Recordset("retension_asumida").Value = "N"
    End If
    fw_compras_comex.Ado_detalle3.Recordset("pago_descripcion") = TxtConcepto.Text
    fw_compras_comex.Ado_detalle3.Recordset("pago_fecha_prog") = DTPFechaProg.Value
    fw_compras_comex.Ado_detalle3.Recordset("pago_fecha_efectiva").Value = DTPFechaPago.Value

    fw_compras_comex.Ado_detalle3.Recordset("pago_descuento_bs").Value = CDbl(TxtDsctoBs)
    fw_compras_comex.Ado_detalle3.Recordset("pago_descuento_dol").Value = CDbl(TxtDsctoDol)
   If TxtMontoDol.Text = "" Or TxtMontoDol.Text = "0" Then
        fw_compras_comex.Ado_detalle3.Recordset("pago_total_dol").Value = CDbl(TxtMontoBs.Text) / GlTipoCambioOficial
   Else
        fw_compras_comex.Ado_detalle3.Recordset("pago_total_dol").Value = CDbl(TxtMontoDol.Text)
        fw_compras_comex.Ado_detalle3.Recordset("pago_total_bs").Value = CDbl(TxtMontoBs.Text)
   End If
   
   If TxtMontoBs.Text = "" Or TxtMontoBs.Text = "0" Then
        fw_compras_comex.Ado_detalle3.Recordset("pago_total_bs").Value = CDbl(TxtMontoBs.Text) * GlTipoCambioOficial
   Else
        fw_compras_comex.Ado_detalle3.Recordset("pago_total_dol").Value = CDbl(TxtMontoDol.Text)
        fw_compras_comex.Ado_detalle3.Recordset("pago_total_bs").Value = CDbl(TxtMontoBs.Text)
   End If
   'fw_compras_comex.Ado_detalle3.Recordset("pago_monto_bs").Value = fw_compras_comex.Ado_detalle3.Recordset("pago_total_bs").Value
    fw_compras_comex.Ado_detalle3.Recordset("pago_monto_bs").Value = CDbl(TxtImporteBs.Text)
    fw_compras_comex.Ado_detalle3.Recordset("pago_monto_dol").Value = CDbl(TxtImporteDol.Text)       'CDbl(TxtImporteBs.Text) / GlTipoCambioOficial
    
    fw_compras_comex.Ado_detalle3.Recordset("pago_nro_cmpbte_factura").Value = txt_factura.Text         'Factura
    fw_compras_comex.Ado_detalle3.Recordset("pago_nro_autorizacion").Value = ""     'IIf(txtDoc.Text = "", "0", txtDoc.Text)            'Autorizacion
    
    fw_compras_comex.Ado_detalle3.Recordset("pago_respaldos") = txt_respaldos.Text
    'Ado_datos.Recordset!Literal = Literal(CStr(Ado_datos.Recordset!cobranza_total_bs)) + " BOLIVIANOS"
    fw_compras_comex.Ado_detalle3.Recordset("literal").Value = Literal(CStr(fw_compras_comex.Ado_detalle3.Recordset!pago_total_dol)) + " DOLARES AMERICANOS"
    fw_compras_comex.Ado_detalle3.Recordset("poa_codigo").Value = "4.1.1"
    
    fw_compras_comex.Ado_detalle3.Recordset("estado_codigo") = "REG"
    fw_compras_comex.Ado_detalle3.Recordset("usr_codigo") = glusuario
    fw_compras_comex.Ado_detalle3.Recordset("fecha_registro") = Format(Date, "dd/mm/yyyy")
    fw_compras_comex.Ado_detalle3.Recordset("hora_registro") = Format(Time, "HH:mm:ss")
    fw_compras_comex.Ado_detalle3.Recordset("cta_codigo") = dtc_cuenta_cod.Text
    fw_compras_comex.Ado_detalle3.Recordset("tipo_cambio") = TxtDscto.Text
    
'    sino = MsgBox("Desea APROBAR el Registro ? (Ya no podrá modificarlo)", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
'        Select Case fw_compras_comex.Ado_datos.Recordset("modalidad_codigo")
'            Case "INVD"    'INVITACION DIRECTA
'                fw_compras_comex.Ado_detalle3.Recordset("estado_codigo") = "APR"
'                Call GRABA_FICHA
'            Case "CPEX"    'CONVOCATORIA PUBLICA EXTERNA
'                fw_compras_comex.Ado_detalle3.Recordset("estado_codigo") = "APR"
'                Call GRABA_FICHA
'            Case "CPIN"    'CONVOCATORIA PUBLICA INTERNA
'                fw_compras_comex.Ado_detalle3.Recordset("estado_codigo") = "APR"
'                Call GRABA_FICHA
'        End Select
'    Else
'        fw_compras_comex.Ado_detalle3.Recordset("estado_codigo") = "REG"
'    End If

    fw_compras_comex.Ado_detalle3.Recordset.Update
   Para_Aceptado = "S"
   'frm_ao_solicitud_rrhh.Ado_detalle3.Refresh '.Recordset.Requery
'   txtSW = "0"
   Unload Me
End If
End Sub

Function Valida()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
    Valida = True
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
  If (Txt_campo1.Text = "") Then
    MsgBox "Debe registrar ... " + lblprov.Caption, vbCritical + vbExclamation, "Validación de datos"
    Valida = False
  End If
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
'  If (dtc_codigo4.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
'  If txtPat = "" Then
'        Valida = False
'    End If
'    If txtNom = "" Then
'        Valida = False
'    End If
    If TxtMontoBs.Text = "" Or TxtMontoBs.Text = "0" Then
        MsgBox "Debe registrar ... " + lbl_monto.Caption, vbCritical + vbExclamation, "Validación de datos"
        Valida = False
    End If
    
End Function

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux4.BoundText
    dtc_desc5.BoundText = dtc_aux4.BoundText
    dtc_aux5.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux5.BoundText
    dtc_desc5.BoundText = dtc_aux5.BoundText
    dtc_aux4.BoundText = dtc_aux5.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux4.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub Chk_fac_LostFocus()
    TxtDsctoBs.Text = "0"
    If Chk_fac.Value = 1 Then
        Chk_Ret.Visible = False
        TxtDsctoBs.Visible = False
    Else
        Chk_Ret.Visible = True
        TxtDsctoBs.Visible = True
    End If
End Sub

Private Sub Chk_Ret_Click()
    If Chk_Ret.Value = 1 Then       'ASUMIDA
        If fw_compras_comex.Ado_detalle3.Recordset!bien_o_servicio = "S" Then
            'SERVICIOS
            'C1 - 900 / 84.5% = 15.5% (rET) --- 12.5 IUE RET A TERCEROS y 3% IT rETEN A TERCEROS
            'TxtDsctoBs.Visible = False
            TxtDsctoBs.Text = CDbl(TxtImporteBs.Text) * 0.155
            TxtMontoBs.Text = CDbl(TxtImporteBs.Text) - CDbl(TxtDsctoBs.Text)
            TxtMontoDol.Text = CDbl(TxtMontoBs.Text) / GlTipoCambioOficial
        Else
            'BIENES
            'C1 - 900 / 92% = 8% (rET) --- 5% IUE RET A TERCEROS y 3% IT rETEN A TERCEROS
            'TxtDsctoBs.Visible = False
            TxtDsctoBs.Text = CDbl(TxtImporteBs.Text) * 0.08
            TxtMontoBs.Text = CDbl(TxtImporteBs.Text) - CDbl(TxtDsctoBs.Text)
            TxtMontoDol.Text = CDbl(TxtMontoBs.Text) / GlTipoCambioOficial
        End If
    Else                            'RETENIDA
        If fw_compras_comex.Ado_detalle3.Recordset!bien_o_servicios = "S" Then
            'SERVICIOS
            '900 * (900*15.5%) = 139.5
            'PAGO EFEC 760.5
            'TxtDsctoBs.Visible = False
            TxtDsctoBs.Text = CDbl(TxtImporteBs.Text) * 0.155
            TxtMontoBs.Text = CDbl(TxtImporteBs.Text) - CDbl(TxtDsctoBs.Text)
            TxtMontoDol.Text = CDbl(TxtMontoBs.Text) / GlTipoCambioOficial
        Else
            'BIENES
            '900 * (900*8%) = 72
            'PAGO EFEC 828
            'TxtDsctoBs.Visible = False
            TxtDsctoBs.Text = CDbl(TxtImporteBs.Text) * 0.08
            TxtMontoBs.Text = CDbl(TxtImporteBs.Text) - CDbl(TxtDsctoBs.Text)
            TxtMontoDol.Text = CDbl(TxtMontoBs.Text) / GlTipoCambioOficial
        End If
    End If
End Sub

Private Sub dtc_codigo_Click(Area As Integer)
    dtc_desc.BoundText = dtc_codigo.BoundText
End Sub

Private Sub dtc_cuenta_cod_Click(Area As Integer)
    dtc_cuenta_desc.BoundText = dtc_cuenta_cod.BoundText
    dtc_cuenta_moneda.BoundText = dtc_cuenta_cod.BoundText
End Sub

Private Sub dtc_cuenta_cod_LostFocus()
    If dtc_cuenta_moneda.Text = "USD" Then
        TxtImporteBs.Enabled = False
        TxtDsctoBs.Enabled = False
        TxtImporteDol.Enabled = True
        TxtDsctoDol.Enabled = True
        CAL_DOL = "SI"
        CAL_BS = "NO"
    Else
        TxtImporteBs.Enabled = True
        TxtDsctoBs.Enabled = True
        TxtImporteDol.Enabled = False
        TxtDsctoDol.Enabled = False
        CAL_DOL = "NO"
        CAL_BS = "SI"
    End If
End Sub

Private Sub dtc_cuenta_desc_Click(Area As Integer)
    dtc_cuenta_cod.BoundText = dtc_cuenta_desc.BoundText
    dtc_cuenta_moneda.BoundText = dtc_cuenta_desc.BoundText
End Sub

Private Sub dtc_cuenta_moneda_Click(Area As Integer)
    dtc_cuenta_desc.BoundText = dtc_cuenta_moneda.BoundText
    dtc_cuenta_cod.BoundText = dtc_cuenta_moneda.BoundText
End Sub

Private Sub dtc_desc_Click(Area As Integer)
    dtc_codigo.BoundText = dtc_desc.BoundText
End Sub

Private Sub Form_Activate()
'    Set rs_clasif5 = New ADODB.Recordset
'    If rs_clasif5.State = 1 Then rs_clasif5.Close
'    Select Case Glaux
'        Case "PROVI"    'PROVISION DE EQUIPOS
'            rs_clasif5.Open "SELECT * FROM gc_beneficiario  ORDER BY beneficiario_denominacion ", db, adOpenStatic  'where pais_codigo= '" & txt_pais.Text & "'
'        Case "TRANS"    'TRANSPORTE
'            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
'        Case "ADUAN"    'DESADUANIZACION
'            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
'        Case "DESCA"    'DESCARGUIO Y OTROS
'            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
'    End Select
'    Set Ado_clasif5.Recordset = rs_clasif5
'    Txt_descripcion.BoundText = txt_campo1.BoundText
'    TxtDscto.Text = GlTipoCambioOficial
    CAL_DOL = "NO"
    CAL_BS = "NO"
End Sub

Private Sub Form_Load()
    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    rs_clasif1.Open "SELECT * FROM rv_unidad_vs_responsable where unidad_codigo = 'COMEX' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif1.Recordset = rs_clasif1
    dtc_desc.BoundText = dtc_codigo.BoundText
    
    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    rs_clasif5.Open "SELECT * FROM gc_beneficiario ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5
    Txt_descripcion.BoundText = Txt_campo1.BoundText
    
  Set rs_cuentas = New ADODB.Recordset
  If rs_cuentas.State = 1 Then rs_cuentas.Close
  rs_cuentas.Open "SELECT * FROM fc_cuenta_bancaria where cta_codigo_tgn='003' ORDER BY cta_descripcion ", db, adOpenStatic
  Set Ado_cuentas.Recordset = rs_cuentas
  dtc_cuenta_cod.BoundText = dtc_cuenta_desc.BoundText
  dtc_cuenta_desc.BoundText = dtc_cuenta_cod.BoundText
    'txtSW = "0"
	Call SeguridadSet(Me)
End Sub

Private Sub txt_campo1_Change()
    Txt_descripcion.BoundText = Txt_campo1.BoundText
End Sub

Private Sub Txt_campo1_Click(Area As Integer)
    Txt_descripcion.BoundText = Txt_campo1.BoundText
End Sub

Private Sub Txt_descripcion_Change()
    Txt_campo1.BoundText = Txt_descripcion.BoundText
End Sub

Private Sub Txt_descripcion_Click(Area As Integer)
    Txt_campo1.BoundText = Txt_descripcion.BoundText
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    CAL_DOL = "NO"
    CAL_BS = "SI"

    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
    Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub TxtDscto_Change()
On Error GoTo UpdateErr
If CAL_BS = "SI" Then
    If TxtMontoDol.Text = "" Or TxtMontoDol.Text = "," Or TxtMontoDol.Text = "0" Then
        TxtMontoBs.Text = Format(0, "###,###,##0.00")
        Exit Sub
    End If
    If TxtDscto.Text <> "" And TxtDscto.Text <> "," And TxtDscto.Text <> "0" Then
        TxtMontoDol.Text = Format(CDbl(TxtMontoBs) / CDbl(TxtDscto), "###,###,##0.00")
    Else
        MsgBox ("Debe Ingresar el Tipo De Cambio (TDC)")
    End If
End If

If CAL_DOL = "SI" Then
    If TxtMontoBs.Text = "" Or TxtMontoBs.Text = "," Or TxtMontoBs.Text = "0" Then
        TxtMontoDol.Text = Format(0, "###,###,##0.00")
        Exit Sub
    End If
    If TxtDscto.Text <> "" And TxtDscto.Text <> "," And TxtDscto.Text <> "0" Then
        TxtMontoBs.Text = Format(CDbl(TxtMontoDol) * CDbl(TxtDscto), "###,###,##0.00")
    Else
        MsgBox ("Debe Ingresar el Tipo De Cambio (TDC)")
    End If
End If
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub TxtDscto_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
    Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub TxtDsctoBs_Change()
If TxtDsctoBs.Text = "" Then
    TxtDsctoBs.Text = "0"
End If
If TxtImporteBs.Text = "" Then
    TxtImporteBs.Text = "0"
End If
If CAL_BS = "SI" Then
    'If TxtMontoBs.Text = "" Or TxtMontoBs.Text = "," Or TxtMontoBs.Text = "0" Then
    If TxtImporteBs.Text = "" Or TxtImporteBs.Text = "," Or TxtImporteBs.Text = "0" Then
        TxtMontoDol.Text = Format(0, "###,###,##0.00")
        TxtDsctoBs.Text = Format(0, "###,###,##0.00")
        TxtMontoBs.Text = Format(0, "###,###,##0.00")
        Exit Sub
    End If
    If TxtDscto.Text <> "" And TxtDscto.Text <> "," And TxtDscto.Text <> "0" Then
        TxtMontoBs.Text = Format(CDbl(TxtImporteBs) - CDbl(TxtDsctoBs), "###,###,##0.00")
        TxtMontoDol.Text = Format(CDbl(TxtMontoBs) / CDbl(TxtDscto), "###,###,##0.00")
    Else
        MsgBox ("Debe Ingresar el Tipo de Cambio ...")
    End If
    'CAL_DOL = "NO"
    'CAL_BS = "NO"
End If

End Sub

Private Sub TxtImporteBs_Change()
    If TxtDsctoBs.Text = "" Then
        TxtDsctoBs.Text = "0"
    End If
    If TxtImporteBs.Text = "" Then
        TxtImporteBs.Text = "0"
    End If
    If CAL_BS = "SI" Then
        'If TxtMontoBs.Text = "" Or TxtMontoBs.Text = "," Or TxtMontoBs.Text = "0" Then
        If TxtImporteBs.Text = "" Or TxtImporteBs.Text = "," Or TxtImporteBs.Text = "0" Then
            TxtMontoDol.Text = Format(0, "###,###,##0.00")
            TxtDsctoBs.Text = Format(0, "###,###,##0.00")
            TxtMontoBs.Text = Format(0, "###,###,##0.00")
            Exit Sub
        End If
        If TxtDscto.Text <> "" And TxtDscto.Text <> "," And TxtDscto.Text <> "0" Then
            TxtMontoBs.Text = Format(CDbl(TxtImporteBs) - CDbl(TxtDsctoBs), "###,###,##0.00")
            TxtMontoDol.Text = Format(CDbl(TxtMontoBs) / CDbl(TxtDscto), "###,###,##0.00")
        Else
            MsgBox ("Debe Ingresar el Tipo de Cambio ...")
        End If
    End If

End Sub

Private Sub TxtImporteBs_KeyPress(KeyAscii As Integer)
    'CAL_DOL = "NO"
    'CAL_BS = "SI"
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
    Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub TxtImporteDol_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
    Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub TxtImporteDol_Change()
On Error GoTo UpdateErr
    If CAL_DOL = "SI" Then
        If TxtImporteDol.Text = "" Or TxtImporteDol.Text = "," Or TxtImporteDol.Text = "0" Then
            TxtImporteBs.Text = Format(0, "###,###,##0.00")
            Exit Sub
        End If
        If TxtDscto.Text <> "" And TxtDscto.Text <> "," And TxtDscto.Text <> "0" Then
            TxtDsctoBs.Text = "0"
            TxtDsctoDol.Text = "0"
            TxtImporteBs.Text = Format(CDbl(TxtImporteDol) * CDbl(TxtDscto), "###,###,##0.00")
            TxtMontoDol.Text = CDbl(TxtImporteDol.Text) - CDbl(TxtDsctoDol.Text)
            TxtMontoBs.Text = CDbl(TxtImporteBs.Text) - CDbl(TxtDsctoBs.Text)
        Else
            MsgBox ("Debe Ingresar el Tipo de Cambio (TDC)")
        End If
    End If
    
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

