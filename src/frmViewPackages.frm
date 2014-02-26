VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmViewPackages 
   Caption         =   "View Packages"
   ClientHeight    =   6900
   ClientLeft      =   1695
   ClientTop       =   2625
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   7800
      TabIndex        =   16
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   5520
      TabIndex        =   15
      Top             =   5880
      Width           =   1935
   End
   Begin MSDataListLib.DataList dlPackageID 
      Height          =   5325
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   9393
      _Version        =   393216
   End
   Begin VB.Label txtPackageID 
      Caption         =   "Package ID "
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
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label txtPackageName 
      Caption         =   "Package Name "
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
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label txtDestination 
      Caption         =   "Destination "
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
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label txtDescription 
      Caption         =   "Description "
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
      Height          =   1575
      Left            =   5160
      TabIndex        =   10
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label txtAmount 
      Caption         =   "Amount"
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
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label txtDistance 
      Caption         =   "Distance "
      DataField       =   "Distance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label txtExpenditure 
      Caption         =   "Expenditure"
      DataField       =   "Expenditure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
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
      Left            =   3000
      TabIndex        =   6
      Top             =   360
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
      Left            =   3000
      TabIndex        =   5
      Top             =   960
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
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
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
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
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
      Left            =   3000
      TabIndex        =   2
      Top             =   3960
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
      Left            =   3000
      TabIndex        =   1
      Top             =   4560
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
      Left            =   3000
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
End
Attribute VB_Name = "frmViewPackages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rsPackages As Recordset

Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub dlPackageID_Click()
On Error Resume Next
rsPackages.Move (dlPackageID.SelectedItem - 1), 1

End Sub

Public Sub Form_Load()
    Set cn = New Connection                           'Memory Allocation
    Set rsPackages = New Recordset
    
    
    cn.CursorLocation = adUseClient
    
    cn.Open cConnect
    rsPackages.Open "Select * from dbPackage", cn, adOpenStatic, adLockOptimistic
               
    Set dlPackageID.RowSource = rsPackages
    dlPackageID.ListField = "PackageID"
       
    Set txtPackageID.DataSource = rsPackages
    Set txtPackageName.DataSource = rsPackages
    Set txtDestination.DataSource = rsPackages
    Set txtDescription.DataSource = rsPackages
    Set txtAmount.DataSource = rsPackages
    Set txtExpenditure.DataSource = rsPackages
    Set txtDistance.DataSource = rsPackages
End Sub


Private Sub cmdDelete_Click()
If MsgBox("Are You Sure. Do You want to Delete ?", vbYesNo, "Conformation") = vbYes Then
    rsPackages.Delete
    dlPackageID.ReFill
    dlPackageID.Refresh
End If
End Sub
