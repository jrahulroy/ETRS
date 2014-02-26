VERSION 5.00
Begin VB.Form frmVehiclesMenu 
   Caption         =   "Vehicles Menu"
   ClientHeight    =   4200
   ClientLeft      =   1800
   ClientTop       =   2100
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8895
   Begin VB.CommandButton cmdDelVehicle 
      Caption         =   "Delete Vehicle"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdVehicleHistory 
      Caption         =   "Vehicle Previous History"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdVehiclesWorked 
      Caption         =   "Vehicles Worked"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CommandButton cmdAddVehicle 
      Caption         =   "Add Vehicle"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   3525
      Left            =   3840
      Picture         =   "frmVehiclesMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4725
   End
End
Attribute VB_Name = "frmVehiclesMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection

Private Sub cmdAddVehicle_Click()
frmAddVehicle.Show vbModal
End Sub

Private Sub cmdDelVehicle_Click()
Set cn = New Connection
cn.Open cConnect
cn.Execute "Delete * from dbVehicles where SlNo = '" & InputBox("Enter the Serial Number") & "'"
cn.Close
End Sub

Private Sub cmdVehicleHistory_Click()
frmVehicleHistory.Show vbModal
End Sub

Private Sub cmdVehiclesWorked_Click()
frmVehiclesWorked.Show vbModal
End Sub
