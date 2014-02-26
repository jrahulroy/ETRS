VERSION 5.00
Begin VB.Form frmAddVehicle 
   Caption         =   "Add Vehicle"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPermitExpiry 
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
      Left            =   2280
      TabIndex        =   9
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox txtPermitNo 
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
      Left            =   2280
      TabIndex        =   8
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox txtRCNo 
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
      Left            =   2280
      TabIndex        =   7
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   6120
      Width           =   2775
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   735
      Left            =   3600
      TabIndex        =   11
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox txtVehicleID 
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
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox txtType 
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
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.TextBox txtSlNo 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtCapacity 
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
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtDriver 
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
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox txtCleaner 
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
      Left            =   2280
      TabIndex        =   6
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "Permit Expiry"
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
      Left            =   240
      TabIndex        =   19
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Permit Number"
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
      Left            =   240
      TabIndex        =   18
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "RC Number :"
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
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label ID 
      Caption         =   "Vehicle ID :"
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
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Vehicle Type :"
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
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Serial Number :"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblType 
      Caption         =   "Capacity :"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblSlNo 
      Caption         =   "Driver :"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Cleaner"
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
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rsVehicles As Recordset

Private Sub cmdClose_Click()
If MsgBox("Are you sure ? Do you want to Close", vbYesNo, "Conformation") = vbYes Then
    Me.Hide
Else
    Exit Sub
End If
End Sub

Private Sub cmdSave_Click()
If MsgBox("Do you want to Save", vbYesNo, "Conformation") = vbYes Then
    rsVehicles.AddNew
    rsVehicles("VehicleID") = txtVehicleID
    rsVehicles("Type") = txtType
    rsVehicles("Capacity") = txtCapacity
    rsVehicles("SlNo") = txtSlNo
    rsVehicles("Cleaner") = txtCleaner
    rsVehicles("Driver") = txtDriver
    rsVehicles("RCNO") = txtRCNo
    rsVehicles("PermitNo") = txtPermitNo
    rsVehicles("PermitExpiry") = txtPermitExpiry
    rsVehicles.Update
    Me.Hide
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
    Set cn = New Connection                           'Memory Allocation
    Set rsVehicles = New Recordset
    
    cn.CursorLocation = adUseClient
    cn.Open cConnect
    
    rsVehicles.Open "Select * from dbVehicles order by VehicleID", cn, adOpenDynamic, adLockOptimistic
    rsVehicles.MoveLast
    txtVehicleID = CInt(rsVehicles("VehicleID")) + 1
End Sub

