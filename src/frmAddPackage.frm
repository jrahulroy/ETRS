VERSION 5.00
Begin VB.Form frmAddPackage 
   Caption         =   "Add Package"
   ClientHeight    =   7350
   ClientLeft      =   1875
   ClientTop       =   2625
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   13185
   Begin VB.TextBox txtExpenditure 
      DataField       =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txtDistance 
      DataField       =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox txtAmount 
      DataField       =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   735
      Left            =   9240
      TabIndex        =   8
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7440
      TabIndex        =   3
      Top             =   2280
      Width           =   5295
   End
   Begin VB.TextBox txtDestination 
      DataField       =   "Destination"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox txtName 
      DataField       =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtPID 
      DataField       =   "PackageID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   6165
      Left            =   240
      Picture         =   "frmAddPackage.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4725
   End
   Begin VB.Label ID 
      Caption         =   "Package ID :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Package Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Destination :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Distance :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Expenditure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rsPackages As Recordset

Private Sub cmdClose_Click()
If MsgBox("Are you sure ? Do you want to Close", vbYesNo, "Conformation") = vbYes Then
    Me.Hide
Else
    Exit Sub
End If
End Sub

Private Sub cmdSave_Click()
If MsgBox("Do you want to Save", vbYesNo, "Conformation") = vbYes Then
    rsPackages.AddNew
    rsPackages("PackageID") = txtPID
    rsPackages("Name") = txtName
    rsPackages("Destination") = txtDestination
    rsPackages("Description") = txtDescription
    rsPackages("Amount") = txtAmount
    rsPackages("Distance") = txtDistance
    rsPackages("Expenditure") = txtExpenditure
    rsPackages.Update
    Me.Hide
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
    Set cn = New Connection                           'Memory Allocation
    Set rsPackages = New Recordset
    
    cn.CursorLocation = adUseClient
    cn.Open cConnect
    
    rsPackages.Open "Select * from dbPackage order by Packageid", cn, adOpenDynamic, adLockOptimistic
    rsPackages.MoveLast
    txtPID = rsPackages("PackageID") + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsPackages.Close
cn.Close
End Sub
