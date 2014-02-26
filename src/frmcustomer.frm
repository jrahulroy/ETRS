VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmcustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   8520
      TabIndex        =   15
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   6600
      TabIndex        =   14
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEW"
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   8280
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   960
      TabIndex        =   10
      Top             =   5880
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3413
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
            LCID            =   16393
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
            LCID            =   16393
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
   Begin VB.TextBox txtcphno 
      DataField       =   "cphno"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox txtcaddress 
      DataField       =   "caddress"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtcname 
      DataField       =   "cname"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtcid 
      DataField       =   "cid"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtfilter 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "CPHNO"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "CADDRESS"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "CNAME"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "CID"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmcustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rs As Recordset


Private Sub Command1_Click()
txtfilter.Text = ""

End Sub

Private Sub Command2_Click()
rs.AddNew
End Sub

Private Sub Command3_Click()
txtcid.Enabled = True
txtcname.Enabled = True
txtcaddress.Enabled = True
txtcphno.Enabled = True
txtcid.SetFocus


End Sub

Private Sub Command4_Click()
rs.Update

End Sub

Private Sub Command5_Click()
rs.Delete
End Sub





Private Sub Command6_Click()
Me.Hide

End Sub

Private Sub Form_Load()
  Dim sSQL  As String
    Set cn = New Connection                           'Memory Allocation
    Set rs = New Recordset
    
    sSQL = "SELECT dbCustomer.* From dbCustomer WHERE (((dbCustomer.cname) Like '" & txtfilter.Text & "%'))"
    cn.CursorLocation = adUseClient
    
    cn.Open cconnect
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic
    
    
    
    
    

    
    
    
    If (rs.EOF) Then
        MsgBox "There are No Matching Customers", vbCritical
        rs.Close
        
        cn.Close
        Exit Sub
    End If
    
    Set DataGrid1.DataSource = rs
    Set txtcid.DataSource = rs
    Set txtcname.DataSource = rs
    Set txtcaddress.DataSource = rs
    Set txtcphno.DataSource = rs
End Sub

Private Sub txtfilter_Change()
Form_Load
txtfilter.SetFocus

End Sub
