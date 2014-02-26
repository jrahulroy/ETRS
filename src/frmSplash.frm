VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9225
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.Timer Timer1 
         Interval        =   6000
         Left            =   2400
         Top             =   360
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   9000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Image Image1 
         Height          =   2280
         Left            =   240
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2280
      End
      Begin VB.Image Image2 
         Height          =   1695
         Left            =   2640
         Picture         =   "frmSplash.frx":242E4
         Top             =   360
         Width           =   6180
      End
      Begin VB.Image Image3 
         Height          =   585
         Left            =   240
         Picture         =   "frmSplash.frx":2F04A
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   8505
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Me.Hide
Unload Me
frmMain.Show
End Sub
