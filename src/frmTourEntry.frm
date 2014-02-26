VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTourEntry 
   Caption         =   "Reserve A Tour"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "frmTourEntry.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dgBookDate 
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3413
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save && Print"
      Height          =   855
      Left            =   11880
      TabIndex        =   6
      Top             =   9240
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dgReservations 
      Height          =   2295
      Left            =   240
      TabIndex        =   18
      Top             =   6720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4048
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcbPackage 
      Height          =   390
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.TextBox txtNumberofPersons 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0"
      Top             =   9240
      Width           =   2055
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0"
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete Customer"
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
      Left            =   11640
      TabIndex        =   5
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Customer"
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
      Left            =   11640
      TabIndex        =   4
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox txtTransID 
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
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtTransDate 
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo dcbCustomers 
      Height          =   390
      Left            =   2160
      TabIndex        =   3
      Top             =   5520
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      Text            =   "Select Customer"
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
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   11520
      Picture         =   "frmTourEntry.frx":B223F
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2385
   End
   Begin VB.Label txtDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "txtDescription"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6960
      TabIndex        =   28
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label txtAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   7080
      TabIndex        =   27
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label txtAmount 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   2040
      TabIndex        =   26
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label txtDestination 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination "
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
      Left            =   6960
      TabIndex        =   25
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label txtTourID 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour ID "
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
      Left            =   2040
      TabIndex        =   24
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label9 
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
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label8 
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
      Left            =   5160
      TabIndex        =   22
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
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
      Left            =   5160
      TabIndex        =   21
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
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
      Left            =   5280
      TabIndex        =   20
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Customers :"
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
      Left            =   360
      TabIndex        =   19
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblNumberofPersons 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Persons :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   16
      Top             =   9240
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      Top             =   9720
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Group Members :"
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
      Left            =   360
      TabIndex        =   12
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   17160
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Upcoming Tours :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   17160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblTourNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Trans ID :"
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
      Left            =   10560
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblTransNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblPackage 
      BackStyle       =   0  'Transparent
      Caption         =   "Package Name:"
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
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmTourEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsPackage As Recordset
Public rsResCustomers As Recordset
Public rsReservations As Recordset
Public rsCustomers As Recordset

Public rsReceipt As Recordset
Public rsReceiptCustomer As Recordset

Public rsTours As Recordset

Public cn As Connection
Public rcn As Connection

Private Sub cmdAdd_Click()


rsResCustomers.AddNew
rsResCustomers("TransID") = txtTransID
rsResCustomers("CName") = dcbCustomers
rsResCustomers("CAddress") = txtAddress

dgReservations.Refresh
txtNumberofPersons = dgReservations.ApproxCount
txtTotalAmount = dgReservations.ApproxCount * txtAmount

End Sub

Private Sub cmdDel_Click()
On Error Resume Next
rsResCustomers.Delete
End Sub

Private Sub cmdSave_Click()

rsReservations.AddNew
rsReservations("TransID") = txtTransID
rsReservations("TransDate") = txtTransDate
rsReservations("TourID") = rsTours("TourID")
rsReservations("PAmount") = txtAmount
rsReservations("NumberofPersons") = txtNumberofPersons
rsReservations("TotalAmount") = txtTotalAmount

rsResCustomers.Update
rsReservations.Update
cn.Close

Set rcn = New Connection
Set rsReceipt = New Recordset
Set rsReceiptCustomer = New Recordset

rcn.Open cConnect
rsReceipt.Open "Select * From qryReservationDetails Where TransID = '" & txtTransID & "'", rcn
rsReceiptCustomer.Open "Select * From dbResCustomers Where TransID = '" & txtTransID & "'", rcn
Set rptReceiptCustomers.DataSource = rsReceiptCustomer
Set rptReceipt.DataSource = rsReceipt
Unload Me
rptReceipt.Show
rptReceiptCustomers.Show
End Sub

Private Sub dcbCustomers_Change()
rsCustomers.Requery
rsCustomers.Find "Name = '" & dcbCustomers & "'"
End Sub

Private Sub dcbPackage_Change()
rsPackage.Requery
rsPackage.Find "Name  = '" & dcbPackage & "'"
rsTours.Close
rsTours.Open "Select TourDate, TourID, Capacity, Bookings from qryFreeSeats Where PackageID = '" & rsPackage("PackageID") & "'", cn
Set dgBookDate.DataSource = rsTours

End Sub

Private Sub Form_Load()

Set cn = New Connection
Set rsPackage = New Recordset
Set rsReservations = New Recordset
Set rsResCustomers = New Recordset
Set rsCustomers = New Recordset
Set rsTours = New Recordset


cn.Open cConnect
cn.CursorLocation = adUseClient

rsPackage.Open "SELECT * From dbPackage", cn, adOpenDynamic
rsCustomers.Open "SELECT * From dbCustomer", cn, adOpenDynamic

rsTours.Open "Select TourDate, TourID, Capacity, Bookings As FreeSeats from qryFreeSeats Where PackageID = '" & rsPackage("PackageID") & "'", cn

rsReservations.Open "SELECT * From dbReservations", cn, adOpenDynamic, adLockOptimistic
rsReservations.MoveLast
txtTransID = rsReservations("TransID") + 1

rsResCustomers.Open "SELECT * From dbResCustomers Where TransID = '" & txtTransID & "'", cn, adOpenDynamic, adLockOptimistic


Set dcbPackage.RowSource = rsPackage
dcbPackage.ListField = "Name"

Set dcbCustomers.RowSource = rsCustomers
dcbCustomers.ListField = "Name"

Set txtDestination.DataSource = rsPackage
txtDestination.DataField = "Destination"
Set txtDescription.DataSource = rsPackage
txtDescription.DataField = "Description"
Set txtAmount.DataSource = rsPackage
txtAmount.DataField = "Amount"
Set txtTourID.DataSource = rsPackage
txtTourID.DataField = "PackageID"

Set txtAddress.DataSource = rsCustomers
txtAddress.DataField = "Address"


Set dgReservations.DataSource = rsResCustomers

txtTransDate = Date
End Sub

