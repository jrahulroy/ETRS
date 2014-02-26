VERSION 5.00
Begin VB.Form frmCashFlow 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cash Flow"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDailyIncome 
      Caption         =   "Daily Income"
      Height          =   735
      Left            =   5520
      TabIndex        =   3
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CommandButton cmdMonthlyIncome 
      Caption         =   "Monthly Income"
      Height          =   735
      Left            =   5520
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdVehicleIncome 
      Caption         =   "Vehicle Income"
      Height          =   735
      Left            =   5520
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdPackageIncome 
      Caption         =   "Package Income"
      Height          =   735
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   3525
      Left            =   360
      Picture         =   "frmAMountDetails.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4725
   End
End
Attribute VB_Name = "frmCashFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDailyIncome_Click()
frmDailyIncome.Show vbModal
End Sub

Private Sub cmdMonthlyIncome_Click()
frmMonthlyIncome.Show vbModal
End Sub

Private Sub cmdPackageIncome_Click()
frmPackageIncome.Show vbModal
End Sub

Private Sub cmdVehicleIncome_Click()
frmVehicleIncome.Show vbModal
End Sub

