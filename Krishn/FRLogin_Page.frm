VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Admin Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   21120
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   21120
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   19680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19680
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   12975
      Left            =   0
      Picture         =   "FRLogin_Page.frx":0000
      ScaleHeight     =   12915
      ScaleWidth      =   17235
      TabIndex        =   0
      Top             =   0
      Width           =   17295
      Begin VB.CommandButton Command1 
         Caption         =   "<<  Back to Welcome Page"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Image Image1 
      Height          =   4275
      Left            =   18240
      Picture         =   "FRLogin_Page.frx":7A8F0
      Top             =   8040
      Width           =   4830
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   18000
      TabIndex        =   8
      Top             =   3720
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Index           =   0
      Left            =   18000
      TabIndex        =   3
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   17880
      X2              =   22560
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   18000
      TabIndex        =   2
      Top             =   2040
      Width           =   960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Show
    Unload Form2
End Sub

Private Sub Command2_Click()
    If (Text1.Text = "student" And Text2.Text = "student") Then
        Form3.Show
        Unload Form2
    Else
        Form2.Show
        Text1.Text = ""
        Text2.Text = ""
        MsgBox "Wrong Username or Password!"
    End If
End Sub

Private Sub Command3_Click()
    If (Text1.Text = "admin" And Text2.Text = "admin") Then
        Form4.Show
        Unload Form2
    Else
        Form2.Show
        Text1.Text = ""
        Text2.Text = ""
        MsgBox "Wrong Username or Password!"
    End If
End Sub
