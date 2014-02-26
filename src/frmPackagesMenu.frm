VERSION 5.00
Begin VB.Form frmPackagesMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Packages Menu"
   ClientHeight    =   3900
   ClientLeft      =   1695
   ClientTop       =   2175
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   7995
   Begin VB.CommandButton cmdViewPackages 
      Caption         =   "View Packages"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete Packages"
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Packages"
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   3525
      Left            =   240
      Picture         =   "frmPackagesMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3285
   End
End
Attribute VB_Name = "frmPackagesMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
frmAddPackage.Show vbModal
End Sub

Private Sub cmdDel_Click()
frmDelPackage.Show vbModal
End Sub

Private Sub cmdViewPackages_Click()
frmViewPackages.Show vbModal
End Sub

