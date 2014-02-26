VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVehicleIncome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Vehicle Income"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   10
      Top             =   9360
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
      TabIndex        =   9
      Top             =   8760
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
      TabIndex        =   8
      Top             =   8160
      Width           =   3975
   End
   Begin VB.ComboBox cmbSdMonth 
      Height          =   315
      ItemData        =   "frmVehicleIncome.frx":0000
      Left            =   3360
      List            =   "frmVehicleIncome.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox cmbSdYear 
      Height          =   315
      ItemData        =   "frmVehicleIncome.frx":0053
      Left            =   4560
      List            =   "frmVehicleIncome.frx":00A2
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid dgAggregate 
      Height          =   4935
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8705
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
   Begin MSDataListLib.DataCombo dcbVehicleID 
      Height          =   390
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      Text            =   "Select Vehicle"
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
      TabIndex        =   13
      Top             =   9480
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
      TabIndex        =   12
      Top             =   8880
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
      TabIndex        =   11
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tours Executed by Vehicle :"
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
      Top             =   2400
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
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblVehicle 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frmVehicleIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rsVehicles As Recordset
Public rsAggregate As Recordset
Public rsTotalIncome As Recordset



Private Sub cmdRefresh_Click()
rsTotalIncome.Close
rsTotalIncome.Open "SELECT Sum(Customers) As Customers, Sum(TotalAmount) As TotalAmount, Sum(TotalExpenditure) As TotalExpenditure, Sum(NetIncome) As NetIncome, VehicleID From QryVehiclesIncome WHERE (TourDate Like '" & cmbSdMonth & "/%/" & cmbSdYear & "' and VehicleID = '" & dcbVehicleID & "') GROUP BY VehicleID", cn
On Error Resume Next
txtTotalAmount = rsTotalIncome("TotalAmount")
txtTotalExpenditure = rsTotalIncome("TotalExpenditure")
txtNetIncome = rsTotalIncome("NetIncome")

rsAggregate.Close
Set rsAggregate = New Recordset
rsAggregate.Open "Select TourDate, TourID, Customers, TotalAmount, TotalExpenditure, NetIncome from qryVehiclesIncome where VehicleID = '" & dcbVehicleID & "' and Tourdate Like '" & cmbSdMonth & "/%/" & cmbSdYear & "'", cn
Set dgAggregate.DataSource = rsAggregate
End Sub

Private Sub Form_Load()
Set cn = New Connection
Set rsVehicles = New Recordset
Set rsTotalIncome = New Recordset

cn.CursorLocation = adUseClient
cn.Open cConnect

rsVehicles.Open "Select * from dbVehicles order by VehicleID", cn, adOpenDynamic, adLockOptimistic
rsTotalIncome.Open "SELECT Sum(Customers) As Customers, Sum(TotalAmount) As TotalAmount, Sum(TotalExpenditure) As TotalExpenditure, Sum(NetIncome) As NetIncome, VehicleID From QryVehiclesIncome WHERE (TourDate = '" & "3/15/2010" & "' and VehicleID = '411') GROUP BY VehicleID", cn

Set rsAggregate = New Recordset
rsAggregate.Open "Select TourDate, TourID, Customers, TotalAmount, TotalExpenditure, NetIncome from qryVehiclesIncome where VehicleID = '" & dcbVehicleID & "'", cn

Set dcbVehicleID.RowSource = rsVehicles
dcbVehicleID.ListField = "VehicleID"
End Sub

