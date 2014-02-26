VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomersBrowse 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10500
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txtFilter 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   5520
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7858
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ID"
         Text            =   "CustomerID"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Phone No"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Address"
         Object.Width           =   3246
      EndProperty
      Picture         =   "frmCustomerBrowse.frx":0000
   End
End
Attribute VB_Name = "frmCustomersBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event OnSelect()

Public pCustomerNo As String
Public pName As String
Public pPhNo As String
Public pAddress As String


Public Sub LoadList()
    Dim li As ListItem
    Dim LV As ListView
    Dim vData As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Dim sSQL  As String
    
    sSQL = "SELECT dbCustomer.* From dbCustomer "
            
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
        
    cn.Open cConnect
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
    Set LV = ListView1
    LV.ListItems.Clear
          
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
       rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("CustomerID") & "")
            li.ListSubItems.Add , , rs("Name") & ""
            li.ListSubItems.Add , , rs("PhoneNo") & ""
            li.ListSubItems.Add , , rs("Address") & ""
            li.Tag = rs("CustomerID")
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    cn.Close
     
End Sub



Private Sub cmdRef_Click()
LoadList
End Sub

Private Sub cmdSelect_Click()
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sSQL As String

Set frmEDIT = New frmCustomerEdit

sSQL = "SELECT dbCustomer.* " & _
        "From dbCustomer " & _
        "WHERE (((dbCustomer.CustomerID)='" & ListView1.SelectedItem.Tag & "'))"
        
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

cn.Open cConnect
rs.Open sSQL, cn

pCustomerNo = rs("CustomerID")
pName = rs("Name")
pPhNo = rs("PhoneNo")
pAddress = rs("Address")

rs.Close
cn.Close

RaiseEvent OnSelect

End Sub

Private Sub Form_Activate()
LoadList
End Sub

Private Sub ListView1_DblClick()
cmdSelect_Click
End Sub
