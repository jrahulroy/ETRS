Attribute VB_Name = "mMain"

Public cConnect As String


Sub Main()
    cConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb"
    frmLogin.Show vbModal
End Sub

