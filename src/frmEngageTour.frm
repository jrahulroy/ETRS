VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmEngageTour 
   Caption         =   "EngageTour"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   10200
      TabIndex        =   6
      Top             =   8880
      Width           =   2655
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   8880
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid dgVehicles 
      Height          =   2535
      Left            =   720
      TabIndex        =   4
      Top             =   6000
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   21
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
         Name            =   "Tahoma"
         Size            =   11.25
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
         MarqueeStyle    =   3
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcbPackageID 
      Height          =   405
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   714
      _Version        =   393216
      Style           =   2
      Text            =   "Select Package"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTourDate 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtTourID 
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
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin MSDataListLib.DataCombo dcbGuideID 
      Height          =   405
      Left            =   2160
      TabIndex        =   3
      Top             =   3960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   714
      _Version        =   393216
      Style           =   2
      Text            =   "Select Guide"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   6360
      TabIndex        =   24
      Top             =   240
      Width           =   3975
      _Version        =   524288
      _ExtentX        =   7011
      _ExtentY        =   4260
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
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   10560
      Picture         =   "frmEngageTour.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2385
   End
   Begin VB.Label txtGuideName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label txtPhNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   22
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Label txtPackageName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   21
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label txtDestination 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label txtDistance 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   19
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label txtAmount 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle :"
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
      TabIndex        =   17
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No :"
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
      Left            =   6480
      TabIndex        =   16
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   840
      TabIndex        =   15
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount :"
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
      Left            =   6480
      TabIndex        =   14
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance :"
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
      Left            =   6480
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination :"
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
      Left            =   840
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   840
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Guide ID :"
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
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Date :"
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
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label ID 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour ID :"
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
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmEngageTour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As Connection
Public rsPackages As Recordset
Public rsVehicles As Recordset
Public rsGuides As Recordset
Public rsEngaged As Recordset

Private Sub Calendar1_Click()
txtTourDate = Calendar1.Value
End Sub

Private Sub cmdClose_Click()
If MsgBox("Are you sure ? Do you want to Close", vbYesNo, "Conformation") = vbYes Then
    Me.Hide
Else
    Exit Sub
End If
End Sub

Private Sub cmdSave_Click()
If MsgBox("Do you want to Save", vbYesNo, "Conformation") = vbYes Then
    rsEngaged.AddNew
    rsEngaged("TourID") = txtTourID
    rsEngaged("TourDate") = txtTourDate
    rsEngaged("PackageID") = dcbPackageID
    rsEngaged("GuideID") = dcbGuideID
    rsEngaged("VehicleID") = rsVehicles("VehicleID")
    rsEngaged.Update
Else
    Exit Sub
End If
End Sub

Private Sub dcbPackageID_Change()
rsPackages.Requery
rsPackages.Find "PackageID = '" & dcbPackageID & "'"
End Sub
Private Sub dcbGuideID_Change()
rsGuides.Requery
rsGuides.Find "GuideID = '" & dcbGuideID & "'"
End Sub

Private Sub Form_Load()
    
    txtTourDate = Date
    
    Set cn = New Connection                           'Memory Allocation
    Set rsPackages = New Recordset
    Set rsVehicles = New Recordset
    Set rsGuides = New Recordset
    Set rsEngaged = New Recordset
    
    cn.CursorLocation = adUseClient
    cn.Open cConnect
    
    rsPackages.Open "Select * from dbPackage order by packageid", cn
    rsEngaged.Open "Select * from dbEngaged order by tourid", cn, adOpenDynamic, adLockOptimistic
    
    rsVehicles.Open "Select * from dbVehicles order by vehicleid", cn
    rsGuides.Open "Select * from dbGuide order by guideid", cn
    
    Set dcbPackageID.RowSource = rsPackages
    dcbPackageID.ListField = "PackageID"
    Set txtPackageName.DataSource = rsPackages
    txtPackageName.DataField = "Name"
    Set txtDestination.DataSource = rsPackages
    txtDestination.DataField = "Destination"
    Set txtDistance.DataSource = rsPackages
    txtDistance.DataField = "Distance"
    Set txtAmount.DataSource = rsPackages
    txtAmount.DataField = "Amount"
    
    Set dcbGuideID.RowSource = rsGuides
    dcbGuideID.ListField = "GuideID"
    Set txtGuideName.DataSource = rsGuides
    txtGuideName.DataField = "Name"
     Set txtPhNo.DataSource = rsGuides
    txtPhNo.DataField = "PhoneNo"
    
    Set dgVehicles.DataSource = rsVehicles
    
    rsEngaged.MoveLast
    txtTourID = rsEngaged("TourID") + 1
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsPackages.Close
rsVehicles.Close
cn.Close
End Sub

Private Sub LoadEngaged()
    rsVehicles.Close
    rsGuides.Close
    rsVehicles.Open "Select * from dbVehicles Where VehicleID Not in (Select VehicleID From qryEngagedVehicles Where TourDate = '" & txtTourDate & "') order by Vehicleid", cn
    rsGuides.Open "Select * from dbGuide where GuideID Not In (Select GuideID from qryEngagedGuides where TourDate = '" & txtTourDate & "') order by guideid", cn

    Set dcbGuideID.RowSource = rsGuides
    dcbGuideID.ListField = "GuideID"
    Set txtGuideName.DataSource = rsGuides
    txtGuideName.DataField = "Name"
     Set txtPhNo.DataSource = rsGuides
    txtPhNo.DataField = "PhoneNo"
    
    Set dgVehicles.DataSource = rsVehicles
End Sub

Private Sub txtTourDate_LostFocus()
LoadEngaged
End Sub
