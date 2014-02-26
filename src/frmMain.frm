VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   " Express Travel Reservation System"
   ClientHeight    =   9885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":324A
   ScaleHeight     =   659
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   932
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEngageTour 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Engage Tour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Left            =   9480
      TabIndex        =   6
      Top             =   1800
      Width           =   3960
   End
   Begin VB.Label lblCashFlow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Flow"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Left            =   9480
      TabIndex        =   5
      Top             =   3000
      Width           =   3960
   End
   Begin VB.Label lblVehiclesMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicles Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Left            =   9480
      TabIndex        =   4
      Top             =   4200
      Width           =   3960
   End
   Begin VB.Label lblRevTour 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reserve A Tour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Left            =   6720
      TabIndex        =   3
      Top             =   9000
      Width           =   6720
   End
   Begin VB.Label lblPackages 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Packages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Left            =   9480
      TabIndex        =   2
      Top             =   7800
      Width           =   3960
   End
   Begin VB.Label lblGuides 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guides"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Left            =   9480
      TabIndex        =   1
      Top             =   6600
      Width           =   3960
   End
   Begin VB.Label lblCustomers 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Customers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Left            =   9480
      TabIndex        =   0
      Top             =   5400
      Width           =   3960
   End
   Begin VB.Image imgCustomers 
      Height          =   855
      Left            =   9360
      Picture         =   "frmMain.frx":AA82B
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   4245
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   9360
      Picture         =   "frmMain.frx":B3CC3
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   4245
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   9360
      Picture         =   "frmMain.frx":BD15B
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   4245
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   6360
      Picture         =   "frmMain.frx":C65F3
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   7245
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   9360
      Picture         =   "frmMain.frx":CFA8B
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   4245
   End
   Begin VB.Image Image5 
      Height          =   855
      Left            =   9360
      Picture         =   "frmMain.frx":D8F23
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   4245
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   9360
      Picture         =   "frmMain.frx":E23BB
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   4245
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   7800
      Width           =   2640
   End
   Begin VB.Label lblChangePassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   675
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   9000
      Width           =   4320
   End
   Begin VB.Image Image8 
      Height          =   855
      Left            =   240
      Picture         =   "frmMain.frx":EB853
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   4605
   End
   Begin VB.Image Image7 
      Height          =   855
      Left            =   240
      Picture         =   "frmMain.frx":F4CEB
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   2805
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub lblAbout_Click(Index As Integer)
frmAbout.Show vbModal
End Sub

Private Sub lblCashFlow_Click()
frmCashFlow.Show vbModal
End Sub

Private Sub lblChangePassword_Click(Index As Integer)
frmChangePassword.Show vbModal
End Sub

Private Sub lblCustomers_Click()
frmCustomersLookup.Show vbModal
End Sub

Private Sub lblEngageTour_Click()
frmEngageTour.Show
End Sub

Private Sub lblGuides_Click()
frmGuidesLookup.Show vbModal
End Sub

Private Sub lblPackages_Click()
frmPackagesMenu.Show vbModal
End Sub

Private Sub lblRevTour_Click()
frmTourEntry.Show
End Sub

Private Sub lblVehiclesMenu_Click()
frmVehiclesMenu.Show
End Sub
