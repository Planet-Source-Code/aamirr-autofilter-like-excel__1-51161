Attribute VB_Name = "Connection"
Option Explicit
Public cn As New ADODB.connection
Sub main()
    Call connection
    frmMain.Show
End Sub
Private Sub connection()
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db.mdb;"
End Sub

