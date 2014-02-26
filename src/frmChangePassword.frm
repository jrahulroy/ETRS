VERSION 5.00
Begin VB.Form frmChangePassword 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Change Password"
   ClientHeight    =   3135
   ClientLeft      =   1875
   ClientTop       =   2085
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   9600
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
      Left            =   6840
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
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
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtCPassword 
      DataField       =   "Address"
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
      IMEMode         =   3  'DISABLE
      Left            =   6480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
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
      IMEMode         =   3  'DISABLE
      Left            =   6480
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   360
      Picture         =   "frmChangePassword.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3165
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3960
      X2              =   9480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Enter the new Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Conform Password"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As Recordset
Public cn As Connection


Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim sSQL As String

sSQL = "SELECT * From dbLogin"

Set cn = New Connection
Set rs = New Recordset

cn.Open cConnect
rs.Open sSQL, cn, adOpenStatic, adLockOptimistic

End Sub

Private Sub cmdSave_Click()
If txtPassword = txtCPassword Then
    rs("Password") = txtPassword
    rs.Update
    Me.Hide
    Unload Me
Else
    MsgBox "Please Conform the Password"
    Exit Sub
End If
End Sub



