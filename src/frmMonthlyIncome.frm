VERSION 5.00
Begin VB.Form frmMonthlyIncome 
   Caption         =   "Monthly Income"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerateDays 
      Caption         =   "Generate As Per Days"
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerateTours 
      Caption         =   "Generate As Per Tours"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      ItemData        =   "frmMonthlyIncome.frx":0000
      Left            =   3360
      List            =   "frmMonthlyIncome.frx":004F
      TabIndex        =   1
      Text            =   "Year"
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      ItemData        =   "frmMonthlyIncome.frx":00E9
      Left            =   2160
      List            =   "frmMonthlyIncome.frx":0111
      TabIndex        =   0
      Text            =   "Month"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label ID 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date :"
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
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmMonthlyIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rs As Recordset

Private Sub cmdGenerateDays_Click()
Set cn = New Connection
Set rs = New Recordset

cn.Open cConnect
rs.Open "Select * from qrymonthlyincomedays where TourDate Like '" & cmbMonth & "/%/" & cmbYear & "'", cn

Set rptMonthlyIncomeDays.DataSource = rs
rptMonthlyIncomeDays.Title = "  " & cmbMonth & "/" & cmbYear & " "
rptMonthlyIncomeDays.WindowState = vbMaximized
rptMonthlyIncomeDays.Show vbModal

End Sub

Private Sub cmdGenerateTours_Click()
Set cn = New Connection
Set rs = New Recordset

cn.Open cConnect
rs.Open "Select * from qrymonthlyincometours where TourDate Like '" & cmbMonth & "/%/" & cmbYear & "'", cn, adOpenStatic, adLockOptimistic

Set rptMonthlyIncomeTours.DataSource = rs
rptMonthlyIncomeTours.Title = "  " & cmbMonth & "/" & cmbYear & " "
rptMonthlyIncomeTours.WindowState = vbMaximized
rptMonthlyIncomeTours.Show vbModal

End Sub
