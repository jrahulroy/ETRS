VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVehiclesWorked 
   Caption         =   "Vehicles Worked"
   ClientHeight    =   6675
   ClientLeft      =   2010
   ClientTop       =   2475
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   8055
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
      Left            =   2880
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Generate"
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox cmbSdMonth 
      Height          =   315
      ItemData        =   "frmVehiclesWorked.frx":0000
      Left            =   2760
      List            =   "frmVehiclesWorked.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cmbSdYear 
      Height          =   315
      ItemData        =   "frmVehiclesWorked.frx":0053
      Left            =   3960
      List            =   "frmVehiclesWorked.frx":00A2
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgWorked 
      Height          =   4095
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7223
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
   Begin VB.Label lblType 
      Caption         =   "Starting Date :"
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
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmVehiclesWorked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rsVehiclesWorked As Recordset


Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub cmdRefresh_Click()
rsVehiclesWorked.Close
Set rsVehiclesWorked = New Recordset
rsVehiclesWorked.Open "SELECT dbEngaged.VehicleID, Count(dbEngaged.TourID) AS Trips, dbVehicles.Type, dbVehicles.SlNo FROM dbVehicles INNER JOIN dbEngaged ON dbVehicles.VehicleID = dbEngaged.VehicleID WHERE (((dbEngaged.TourDate) Like '" & cmbSdMonth & "/%/" & cmbSdYear & "')) GROUP BY dbEngaged.VehicleID, dbVehicles.Type, dbVehicles.SlNo ORDER BY dbEngaged.VehicleID", cn
Set dgWorked.DataSource = rsVehiclesWorked
End Sub

Private Sub Form_Load()
Set cn = New Connection
Set rsVehiclesWorked = New Recordset

cn.CursorLocation = adUseClient
cn.Open cConnect

rsVehiclesWorked.Open "SELECT dbEngaged.VehicleID, Count(dbEngaged.TourID) AS NumberOfTours, dbVehicles.Type, dbVehicles.SlNo FROM dbVehicles INNER JOIN dbEngaged ON dbVehicles.VehicleID = dbEngaged.VehicleID WHERE (((dbEngaged.TourDate) Like '" & cmbSdMonth & "/%/" & cmbSdYear & "')) GROUP BY dbEngaged.VehicleID, dbVehicles.Type, dbVehicles.SlNo ORDER BY dbEngaged.VehicleID", cn

End Sub

