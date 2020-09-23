Attribute VB_Name = "modCDs"
' *** Sistema de Vendas de CDs
' *** Desenvolvido por Frederico Machado

Global Path  As String
Global banco As Database
Global tabcds   As Recordset

Sub Main()
  Path = App.Path
  If Right$(Path, 1) <> "\" Then Path = Path & "\"
  
  frmSplash.Show
  
  Set banco = OpenDatabase(Path & "cds.mdb")
  Set tabcds = banco.OpenRecordset("select * from cds order by codigo")
End Sub
