Attribute VB_Name = "mMain"
Public fMainForm As frmMain
Public cConnect As String


Sub Main()
    Dim fLogin As New frmLogin
    fLogin.Show vbModal
    Unload fLogin


    Set fMainForm = New frmMain
    Load fMainForm
   
   
    
    cConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;User ID=;Data Source=" & App.Path & "\data.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Locale Identifier=1033;Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='danieldave';Jet OLEDB:Global Partial Bulk Ops=2"



    fMainForm.Show
End Sub

