VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVehicleHistory 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Vehicle History"
   ClientHeight    =   8460
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11445
   Begin VB.TextBox txtPetrol 
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
      TabIndex        =   16
      Top             =   7320
      Width           =   3975
   End
   Begin VB.TextBox txtTours 
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
      TabIndex        =   10
      Top             =   6360
      Width           =   3975
   End
   Begin VB.TextBox txtTotalWorked 
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
      Top             =   6960
      Width           =   3975
   End
   Begin VB.TextBox txtIdleDays 
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
      TabIndex        =   8
      Top             =   7560
      Width           =   3975
   End
   Begin VB.TextBox txtTotalDistance 
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
      TabIndex        =   7
      Top             =   6360
      Width           =   3975
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      ItemData        =   "frmVehicleHistory.frx":0000
      Left            =   3720
      List            =   "frmVehicleHistory.frx":004F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      ItemData        =   "frmVehicleHistory.frx":00E9
      Left            =   2520
      List            =   "frmVehicleHistory.frx":0111
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid dgHistory 
      Height          =   3735
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6588
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
      Left            =   2520
      TabIndex        =   1
      Top             =   240
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Amount of Petrol :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   17
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Number of Tours :"
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
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Days Worked"
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
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Idle Days :"
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
      Top             =   7560
      Width           =   2895
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
      TabIndex        =   12
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Distance Travelled :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   11
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label lblVehicle 
      BackStyle       =   0  'Transparent
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
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
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
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frmVehicleHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cn As Connection
Public rsVehicles As Recordset
Public rsTotal As Recordset
Public rsVehicleHistory As Recordset


Private Sub cmdRefresh_Click()
rsTotal.Close
Set rsTotal = New Recordset
rsTotal.Open "Select COunt(TourDate) As TotalTour, Count(TourID) As WorkedDays, 30 - Count(TourDate) As IdleDays, Sum(Distance) As TotalDistance, Sum(Petrol) As TotalPetrol From qryParticularVehicleHistory where VehicleID = '" & dcbVehicleID & "' AND TourDate Like '" & cmbMonth & "/%/" & cmbYear & "' group by tourdate", cn
On Error Resume Next
txtTours = rsTotal("TotalTour")
txtTotalWorked = rsTotal("WorkedDays")
txtIdleDays = rsTotal("IdleDays")
txtTotalDistance = rsTotal("TotalDistance") * 2
txtPetrol = rsTotal("TotalPetrol")


rsVehicleHistory.Close
Set rsVehicles = New Recordset
rsVehicleHistory.Open "Select TourDate, TourID, Customers, Distance, Petrol from qryParticularVehicleHistory where VehicleID = '" & dcbVehicleID & "' AND TourDate Like '" & cmbMonth & "/%/" & cmbYear & "'", cn
Set dgHistory.DataSource = rsVehicleHistory

End Sub

Private Sub Form_Load()
Set cn = New Connection
Set rsVehicles = New Recordset
Set rsTotal = New Recordset

cn.CursorLocation = adUseClient
cn.Open cConnect
rsVehicles.Open "Select * from dbVehicles order by vehicleid", cn, adOpenDynamic, adLockOptimistic
Set rsVehicleHistory = New Recordset

rsVehicleHistory.Open "Select TourDate, TourID, Customers, Distance, Petrol from qryParticularVehicleHistory where VehicleID = '" & dcbVehicleID & "' AND TourDate Like '" & cmbMonth & "/%/" & cmbYear & "'", cn
rsTotal.Open "Select COunt(TourDate) As TotalTour, Count(TourID) As WorkedDays, 30 - Count(TourDate) As IdleDays, Sum(Distance) As TotalDistance, Sum(Petrol) As TotalPetrol From qryParticularVehicleHistory where VehicleID = '" & dcbVehicleID & "' AND TourDate Like '" & cmbMonth & "/%/" & cmbYear & "'", cn
Set dgHistory.DataSource = rsVehicleHistory

Set dcbVehicleID.RowSource = rsVehicles
dcbVehicleID.ListField = "VehicleID"
End Sub
