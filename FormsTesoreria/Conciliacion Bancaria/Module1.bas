Attribute VB_Name = "Module1"
Public db As Connection
Public swConciliacion As String ' Puede tener los valores de CHEQUE, TRABSFERENCIA
Public swConciliados  As String
Public swFiltro  As String
Public Cheq_Transf As String
Sub Main()
   Set db = New Connection
   db.CursorLocation = adUseClient
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=\\sersis\saf\udapre.mdb;"
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=d:\saf-1\udapre.mdb;"
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=d:\saf-1\udapre.mdb;"
   db.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2000;Data Source=sersis"
   'FrmConciliacion.Show
   FrmPresentacion.Show
   'FrmCuentas.Show
End Sub


