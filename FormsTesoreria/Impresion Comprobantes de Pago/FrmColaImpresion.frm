VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmColaImpresion 
   Caption         =   "IM`PRESION DE COMPROBANTES"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   Icon            =   "FrmColaImpresion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DtGComprobante 
      Height          =   7380
      Left            =   1440
      TabIndex        =   11
      Top             =   1110
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   13018
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraOpciones 
      Height          =   7455
      Left            =   75
      TabIndex        =   6
      Top             =   1050
      Width           =   1305
      Begin VB.CommandButton CmdBorrarRegistroSeleccionado 
         Caption         =   "Borrar registro seleccionado"
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   2460
         Width           =   1140
      End
      Begin VB.CommandButton CmdColaImpresion 
         Caption         =   "Ver Cola de Impresion"
         Height          =   735
         Left            =   105
         TabIndex        =   10
         Top             =   960
         Width           =   1140
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Borrar todos los registros"
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   1710
         Width           =   1140
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   735
         Left            =   105
         Picture         =   "FrmColaImpresion.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5805
         Width           =   1140
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   105
         Picture         =   "FrmColaImpresion.frx":130C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "TODO LO QUE ESTA EN LA COLA DE IMPRESION"
         Top             =   225
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   7845
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      Begin VB.Label Label2 
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   4
         Top             =   690
         Width           =   2460
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   3
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   2
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "COLA DE IMPRESION DE COMPROBANTES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   3090
         TabIndex        =   1
         Top             =   225
         Width           =   6495
      End
   End
End
Attribute VB_Name = "FrmColaImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Dim CryTrans As New CryTransferencia
Dim rsc As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
'Dim CryCmpte As New CryComprobante


Private Sub CmdBorrarRegistroSeleccionado_Click()
'========================================================================================
' Módulo:                   CmdBorrarRegistroSeleccionado
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmColaImpresion
' Descipción :              Borra los registros que son marcados
'                           con el mouse sobre el grid
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================
Dim sino As Variant
   sino = MsgBox("Se eliminaran los comprobantes de la cola de impresión !!!!", vbYesNo, "Mensaje de Advertencia")
   'DataGrid1.Row = I
   'rsCheque("numero_cheque") = DataGrid1.Columns(0)
     Set rsc = New ADODB.Recordset
     If rsc.State = 1 Then rsc.Close
     rsc.Open "SELECT * FROM to_comprobantes ", db, adOpenKeyset, adLockOptimistic
     If rsc.RecordCount > 0 Then
            MsgBox "No existen registros para imprimir", vbInformation + vbCritical, "Validación de datos"
            Exit Sub
     End If

   If rsc.State = 1 Then rsc.Close
   If sino = vbYes Then
     Set rsc = New ADODB.Recordset
     If rsc.State = 1 Then rsc.Close
     rsc.Open "SELECT * FROM to_comprobantes ", db, adOpenKeyset, adLockOptimistic
     If rsc.RecordCount > 0 Then
            If rsc.State = 1 Then rsc.Close
            rsc.Open "SELECT * FROM to_comprobantes where Nro_Cmpte='" & DtGComprobante.Columns(0) & "'", db, adOpenKeyset, adLockOptimistic
            If rsc.RecordCount > 0 Then
               rsc.Delete
            End If

     End If
   End If
   
     Set rsComprobante = New ADODB.Recordset
     If rsComprobante.State = 1 Then rsComprobante.Close
     rsComprobante.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
     If rsComprobante.RecordCount > 0 Then
        Set DtGComprobante.DataSource = rsComprobante
     End If

End Sub

Private Sub CmdColaImpresion_Click()
'========================================================================================
' Módulo:                   CmdColaImpresion
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmColaImpresion
' Descipción :              Muestra los registros de Cheq./Transf. impresas
'                           El fin es ahorrar tiempo para no volver a elegir para
'                           imprimir comprobante de pago
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================
     Set rsComprobante = New ADODB.Recordset
     If rsComprobante.State = 1 Then rsComprobante.Close
     rsComprobante.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
     If rsComprobante.RecordCount > 0 Then
        Set DtGComprobante.DataSource = rsComprobante
     Else
        MsgBox "No existen registros para recuperar", vbInformation + vbCritical, "Validación de datos"
     End If
End Sub

Private Sub cmdImprimir_Click()
Dim sino As String

Dim pname As String         'Stores the printer name
Dim pport As String         'Stores the printer port information
Dim pdriver As String       'Stores the printer driver information


     Set rsComprobante = New ADODB.Recordset
     If rsComprobante.State = 1 Then rsComprobante.Close
     rsComprobante.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
     If rsComprobante.RecordCount > 0 Then
        Set DtGComprobante.DataSource = rsComprobante
     Else
       MsgBox "No existen registros para imprimir", vbInformation + vbCritical
       Exit Sub
     End If
     
   sino = MsgBox("Se imprimiran los comprobantes sin Nro. Transf ...!", vbYesNo, "Mensaje de Advertencia")
   If sino = vbYes Then
        pname = "Epson LX-810"
        pport = "LPT1:  (Puerto de impresora ECP)"
        pdriver = "Epson LX-810"
                
        Call CryCmpte.SelectPrinter(pdriver, pname, pport)
        
        'Cmpte_NroTransferencia
        FrmComprobante.Show
        'Verify
        '''CryTrans.PrintOut
        'Transferencia_Aprobados
        '''CryTrans.Database.
   End If
End Sub

Private Sub CmdLimpiar_Click()
'========================================================================================
' Módulo:                   CmdLimpiar
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmColaImpresion
' Descipción :              Permite el borrado de los registros impresos en
'                           comprobantes de pago
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================

Dim sino As Variant
   sino = MsgBox("Se eliminaran los comprobantes de la cola de impresión !!!!", vbYesNo, "Mensaje de Advertencia")
   If sino = vbYes Then
           'eliminar
           db.Execute "DELETE FROM to_comprobantes"
           'crear
           'db.Execute "create"
           'adicionar
           'db.Execute "insert to nonmbre_tabla (opcion_campos_sep_comas) values (valores corresp. a cada campo) "
           'eliminar
           'db.Execute "DELETE FROM to_comprobantes where condicion"
           'Actualizar
           'db.Execute "UPDATE nombre_tabla SET CAMPO=VALUE, CAMPO=VALOR"
           
'                Set rsComprobante = New ADODB.Recordset
'                If rsComprobante.State = 1 Then rsComprobante.Close
'                rsComprobante.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
'                If rsComprobante.RecordCount > 0 Then
'                           While Not rsComprobante.EOF
'                                 rsComprobante.Delete
'                                 rsComprobante.MoveNext
'                           Wend
'                End If
     MsgBox "Cola de impresion eliminado", vbCritical + vbInformation
     
     Set rsComprobante = New ADODB.Recordset
     If rsComprobante.State = 1 Then rsComprobante.Close
     rsComprobante.Open "SELECT * FROM to_comprobantes", db, adOpenKeyset, adLockOptimistic
     If rsComprobante.RecordCount > 0 Then
        Set DtGComprobante.DataSource = rsComprobante
     Else
        Set DtGComprobante.DataSource = rsNada
     End If
   End If
End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub
