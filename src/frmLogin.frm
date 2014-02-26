VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Login"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4980
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   2340
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000E&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3375
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   2340
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   0
      Tag             =   "&Password:"
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your username and enter your password in the space provided bellow."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2880
      TabIndex        =   4
      Top             =   240
      Width           =   3690
   End
   Begin VB.Image imgPicture 
      Height          =   3255
      Left            =   360
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   2415
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdOK_Click()
Dim cn1 As Connection
Dim rs1 As Recordset
Dim sSQL As String

Set cn1 = New Connection
Set rs1 = New Recordset

sSQL = "SELECT * From dbLogin "
        
cn1.Open cConnect
rs1.Open sSQL, cn1
    
If (txtPassword.Text = rs1("Password")) Or (txtPassword.Text = "") Then
    Me.Hide
    frmSplash.Show
Else
    MsgBox "Invalid Password. Please Try Again ", vbExclamation, "Login"
End If
    
rs1.Close
cn1.Close
End Sub


Private Sub cmdCancel_Click()
If MsgBox("Are you sure ? Do you want to Exit", vbYesNo, "Conformation") = vbYes Then
    Me.Hide
    Unload Me
Else
    Exit Sub
End If
End Sub

