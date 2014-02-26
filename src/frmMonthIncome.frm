VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMonthIncome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   11535
   ScaleWidth      =   10230
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
      Width           =   2775
   End
   Begin VB.ComboBox cmbSdMonth 
      Height          =   315
      ItemData        =   "frmMonthIncome.frx":0000
      Left            =   3240
      List            =   "frmMonthIncome.frx":0028
      TabIndex        =   7
      Text            =   "Month"
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cmbSdYear 
      Height          =   315
      ItemData        =   "frmMonthIncome.frx":0053
      Left            =   4440
      List            =   "frmMonthIncome.frx":00A2
      TabIndex        =   6
      Text            =   "Year"
      Top             =   480
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
      TabIndex        =   2
      Top             =   6240
      Width           =   9615
      _ExtentX        =   16960
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
   Begin MSDataGridLib.DataGrid dgIncome 
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4471
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
      TabIndex        =   5
      Top             =   5760
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Transactions under Vehicle :"
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
      TabIndex        =   4
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
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmMonthIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rsVehicles As Recordset
Public rsIncome As Recordset
Public rsAggregate As Recordset
Public rsTotalIncome As Recordset



Private Sub cmdRefresh_Click()
rsIncome.Close
Set rsIncome = New Recordset
rsIncome.Open "SELECT dbReservations.TransID, dbReservations.TransDate, dbReservations.NumberofPersons, dbReservations.TotalAmount, dbReservations.BookDate FROM (dbVehicles INNER JOIN dbEngaged ON dbVehicles.VehicleID = dbEngaged.VehicleID) INNER JOIN dbReservations ON (dbEngaged.TourID = dbReservations.TourID) AND (dbEngaged.TourDate = dbReservations.BookDate) WHERE (((dbEngaged.TourDate) Like '" & cmbSdMonth & "/%/" & cmbSdYear & "') AND ((dbEngaged.VehicleID)='" & dcbVehicleID & "'))", cn
Set dgIncome.DataSource = rsIncome

rsTotalIncome.Close
rsTotalIncome.Open "SELECT Sum(Customers) As Customers, Sum(TotalAmount) As TotalAmount, Sum(TotalExpenditure) As TotalExpenditure, Sum(NetIncome) As NetIncome, VehicleID From QryVehiclesIncome WHERE (TourDate = '" & "3/15/2010" & "' and VehicleID = '411') GROUP BY VehicleID", cn
txtTotalAmount = rsTotalIncome("TotalAmount")
txtTotalExpenditure = rsTotalIncome("TotalExpenditure")
txtNetIncome = rsTotalIncome("NetIncome")

rsAggregate.Close
Set rsAggregate = New Recordset
rsAggregate.Open "Select TourDate, TourID, Customers, TotalAmount, TotalExpenditure, NetIncome from qryVehiclesIncome where VehicleID = '" & dcbVehicleID & "'", cn
Set dgAggregate.DataSource = rsAggregate
End Sub

Private Sub Form_Load()
Set cn = New Connection
Set rsVehicles = New Recordset
Set rsTotalIncome = New Recordset

cn.CursorLocation = adUseClient
cn.Open cConnect

rsVehicles.Open "Select * from dbVehicles", cn, adOpenDynamic, adLockOptimistic
rsTotalIncome.Open "SELECT Sum(Customers) As Customers, Sum(TotalAmount) As TotalAmount, Sum(TotalExpenditure) As TotalExpenditure, Sum(NetIncome) As NetIncome, VehicleID From QryVehiclesIncome WHERE (TourDate = '" & "3/15/2010" & "' and VehicleID = '411') GROUP BY VehicleID", cn
MsgBox rsTotalIncome("NetIncome")


Set rsIncome = New Recordset
Set rsAggregate = New Recordset

rsIncome.Open "SELECT dbReservations.TourID, Sum(dbReservations.NumberofPersons) AS SumOfNumberofPersons, Sum(dbReservations.TotalAmount) AS SumOfTotalAmount FROM (dbVehicles INNER JOIN dbEngaged ON dbVehicles.VehicleID = dbEngaged.VehicleID) INNER JOIN dbReservations ON (dbEngaged.TourDate = dbReservations.BookDate) AND (dbEngaged.TourID = dbReservations.TourID) WHERE (((dbEngaged.TourDate)='" & txtDate & "') AND ((dbEngaged.VehicleID)='" & dcbVehicleID & "')) GROUP BY dbReservations.TourID", cn
rsAggregate.Open "Select TourDate, TourID, Customers, TotalAmount, TotalExpenditure, NetIncome from qryVehiclesIncome where VehicleID = '" & dcbVehicleID & "'", cn

Set dcbVehicleID.RowSource = rsVehicles
dcbVehicleID.ListField = "VehicleID"
End Sub

