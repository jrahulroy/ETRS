VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About Express_Travel_Reservation_System"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14115
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   14115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About Express_Travel_Reservation_System"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   705
      Left            =   9045
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2745
      Width           =   1467
   End
   Begin VB.Image Image3 
      Height          =   3195
      Left            =   360
      Picture         =   "frmAbout.frx":0000
      Top             =   240
      Width           =   4245
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   10800
      Picture         =   "frmAbout.frx":F09C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   9000
      Picture         =   "frmAbout.frx":FD5B
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "Under the Guidance of Mr.Vijay Kumar, Head of DCME"
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   4800
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label lblDescription 
      Caption         =   "Developed as a part of the Project Work towards the completion of the Diploma by Rahul, Ramesh, Imran, Ashok, Bal Reddy, Ramesh"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   4770
      TabIndex        =   4
      Tag             =   "App Description"
      Top             =   1125
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      Caption         =   "Express Travel Reservation System "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   4770
      TabIndex        =   3
      Tag             =   "Application Title"
      Top             =   240
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   4680
      X2              =   10455
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   4680
      X2              =   10455
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      Height          =   225
      Left            =   4770
      TabIndex        =   2
      Tag             =   "Version"
      Top             =   780
      Width           =   4095
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: This Software Uses Data.mdb that should at the location of the EXE file to Enable Database  Functionality of the Software"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   5055
      TabIndex        =   1
      Tag             =   "Warning: ..."
      Top             =   2745
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub picIcon_Click()

End Sub
