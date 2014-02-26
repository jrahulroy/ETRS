VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmDailyIncome 
   Caption         =   "Daily Income"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3015
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   3855
      _Version        =   524288
      _ExtentX        =   6800
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2010
      Month           =   3
      Day             =   18
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   2055
   End
End
Attribute VB_Name = "frmDailyIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rs As Recordset
Public rpt As DataReport


Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub cmdGenerate_Click()
Set cn = New Connection
Set rs = New Recordset

cn.Open cConnect
rs.Open "Select * from qryDailyIncomeTours where TourDate Like '" & Calendar1.Value & "'", cn, adOpenStatic, adLockOptimistic


Set rptDailyIncome.DataSource = rs
rptDailyIncome.Title = "  " & Calendar1.Value & " "
rptDailyIncome.WindowState = vbMaximized
rptDailyIncome.Show vbModal
End Sub
