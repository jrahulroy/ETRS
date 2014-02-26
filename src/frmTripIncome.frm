VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPackageIncome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Package Income"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "frmTripIncome.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dgIncome 
      Height          =   2535
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   -2147483644
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.TextBox txtTotalTours 
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
      Left            =   10080
      TabIndex        =   17
      Top             =   8880
      Width           =   3975
   End
   Begin VB.TextBox txtNetIncome 
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
      Left            =   3360
      TabIndex        =   11
      Top             =   10080
      Width           =   3975
   End
   Begin VB.TextBox txtTotalExpenditure 
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
      Left            =   3360
      TabIndex        =   10
      Top             =   9480
      Width           =   3975
   End
   Begin VB.TextBox txtTotalAmount 
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
      Left            =   3360
      TabIndex        =   9
      Top             =   8880
      Width           =   3975
   End
   Begin VB.ComboBox cmbSdMonth 
      Height          =   315
      ItemData        =   "frmTripIncome.frx":1604E
      Left            =   3360
      List            =   "frmTripIncome.frx":16076
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox cmbSdYear 
      Height          =   315
      ItemData        =   "frmTripIncome.frx":160A1
      Left            =   4560
      List            =   "frmTripIncome.frx":160F0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid dgAggregate 
      Height          =   1695
      Left            =   240
      TabIndex        =   6
      Top             =   6240
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   2990
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin MSDataListLib.DataCombo dcbPackageID 
      Height          =   390
      Left            =   3360
      TabIndex        =   0
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      Text            =   "Select Package"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tours Conducted :"
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
      Left            =   7800
      TabIndex        =   18
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Statistics for the Month"
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
      Top             =   8160
      Width           =   5175
   End
   Begin VB.Label lblVehicle 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1320
      TabIndex        =   15
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Net Income :"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   10200
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Expenditure :"
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
      Left            =   960
      TabIndex        =   13
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Amount :"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tours Executed under Package :"
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
      TabIndex        =   8
      Top             =   5760
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Transactions under Package :"
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
      TabIndex        =   7
      Top             =   2520
      Width           =   5175
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
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frmPackageIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rsPackages As Recordset
Public rsIncome As Recordset
Public rsAggregate As Recordset
Public rsTotalIncome As Recordset



Private Sub cmdRefresh_Click()
rsIncome.Close
Set rsIncome = New Recordset
rsIncome.Open "SELECT TransID, TourID, TourDate, NumberofPersons, TotalAmount, TotalExpenditure, NetIncome from qryTransPackage where packageid = '" & dcbPackageID & "' and TourDate Like '" & cmbSdMonth & "/%/" & cmbSdYear & "'", cn

Set dgIncome.DataSource = rsIncome

rsTotalIncome.Close
rsTotalIncome.Open "Select  Count(TourID) As TotalTours, Sum(TotalAmount) As TotalAmount, Sum(TotalExpenditure) As TotalExpenditure from qryAggregatePackageIncome where PackageID = '" & dcbPackageID & "' and TourDate Like '" & cmbSdMonth & "/%/" & cmbSdYear & "' Group by PackageID", cn
On Error Resume Next
txtTotalAmount = rsTotalIncome("TotalAmount")
txtTotalExpenditure = rsTotalIncome("TotalExpenditure")
txtNetIncome = txtTotalAmount - txtTotalExpenditure
txtTotalTours = rsTotalIncome("TotalTours")

rsAggregate.Close
Set rsAggregate = New Recordset
rsAggregate.Open "Select TourID, TourDate, TotalTransactions, TotalCustomers, TotalAmount, TotalExpenditure, NetIncome from qryTourPackageIncome where PackageID = '" & dcbPackageID & "' and TourDate Like '" & cmbSdMonth & "/%/" & cmbSdYear & "'", cn
Set dgAggregate.DataSource = rsAggregate
End Sub

Private Sub Form_Load()
Set cn = New Connection
Set rsPackages = New Recordset
Set rsTotalIncome = New Recordset

cn.CursorLocation = adUseClient
cn.Open cConnect

rsPackages.Open "Select * from dbPackage", cn, adOpenDynamic, adLockOptimistic
rsTotalIncome.Open "Select  Count(TourID) As TotalTours, Sum(Customers) As TotalCustomers, Sum(TotalAmount) As TotalAmount, Sum(TotalExpenditure) As TotalExpenditure, Sum(NetIncome) As TotalIncome from qryAggregatePackageIncome where PackageID = '" & dcbPackageID & "' and TourDate Like '" & cmbSdMonth & "/%/" & cmbSdYear & "' Group by TourID", cn


Set rsIncome = New Recordset
Set rsAggregate = New Recordset

rsIncome.Open "SELECT TourID, TourDate, NumberofPersons, TotalAmount, TotalExpenditure, NetIncome from qryPackageIncome where packageid = '" & dcbPackageID & "' order by tourid", cn
rsAggregate.Open "Select   TotalTransactions, TotalCustomers, TotalAmount, TotalExpenditure, NetIncome from qryTourPackageIncome where PackageID = '" & dcbPackageID & "' and TourDate Like '" & cmbSdMonth & "/%/" & cmbSdYear & "'", cn

Set dcbPackageID.RowSource = rsPackages
dcbPackageID.ListField = "PackageID"
End Sub

