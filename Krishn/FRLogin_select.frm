VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000005&
   Caption         =   "User Login"
   ClientHeight    =   3015
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   6360
      Picture         =   "FRLogin_select.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Data Data1 
      BackColor       =   &H8000000B&
      Caption         =   "Employee's Record"
      Connect         =   "Access"
      DatabaseName    =   "F:\VB Project\Krishn\MYDATABASE1.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   12120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "mytable1"
      Top             =   12000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      BorderStyle     =   0  'None
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7560
      Width           =   3495
   End
   Begin VB.TextBox Text9 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   10920
      Width           =   3375
   End
   Begin VB.TextBox Text8 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   10080
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   9240
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   8400
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6720
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5880
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
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
      Left            =   480
      TabIndex        =   11
      Top             =   6000
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< Back to Login Page"
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
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   10920
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000005&
      Caption         =   "Calculate Gross Salary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   960
      TabIndex        =   6
      Top             =   9600
      Width           =   3615
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000005&
      Caption         =   "Search Record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   960
      TabIndex        =   5
      Top             =   8760
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   -360
      Picture         =   "FRLogin_select.frx":07C3
      ScaleHeight     =   3975
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   720
      Width           =   6615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   15000
      Picture         =   "FRLogin_select.frx":768F
      ScaleHeight     =   7455
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.Image Image1 
      Height          =   4290
      Left            =   16080
      Picture         =   "FRLogin_select.frx":159E8
      Top             =   7920
      Width           =   6720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No. : "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   7800
      TabIndex        =   24
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name : "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   7800
      TabIndex        =   19
      Top             =   8400
      Width           =   1740
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mother's Name : "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   7800
      TabIndex        =   18
      Top             =   9240
      Width           =   1830
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of the Bank : "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   7800
      TabIndex        =   17
      Top             =   10080
      Width           =   2100
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Ammount :"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   7800
      TabIndex        =   16
      Top             =   10920
      Width           =   2115
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Code : "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7800
      TabIndex        =   15
      Top             =   6720
      Width           =   1980
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Name : "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   7800
      TabIndex        =   14
      Top             =   5880
      Width           =   2085
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Employees's Code or Name : "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Top             =   5520
      Width           =   4065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Select your choice)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   11040
      Width           =   2250
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What would you like to do?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   960
      TabIndex        =   4
      Top             =   7920
      Width           =   3990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "student"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome User: "
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
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   2310
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim str, s1, s2, s3, s4, s5, s6, s7 As String

Private Sub Command1_Click()
    If Option2.Value = True Then
        rs.Open str, cn, adOpenDynamic, adLockPessimistic
        Dim i As Integer
        i = 0
        While Not rs.EOF
            With rs
                If rs!Emp_Code = Text1.Text Or rs!Emp_Name = Text1.Text Then
                    s1 = rs!Emp_Name
                    s2 = rs!Emp_Code
                    s3 = rs!Father_Name
                    s4 = rs!Mother_Name
                    s5 = rs!Bank_Name
                    s6 = rs!Amt_Balance
                    s7 = rs!Contact
                    Text2.Text = s1
                    Text5.Text = s2
                    Text6.Text = s3
                    Text7.Text = s4
                    Text8.Text = s5
                    Text9.Text = s6
                    Text10.Text = s7
                    i = 1
                End If
            End With
            rs.MoveNext
        Wend
        If i = 0 Then
            Text2.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            MsgBox "Record Not Found!!"
        End If
        rs.Close
    Else
        If Option3.Value = True Then
            Text2.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            rs.Open str, cn, adOpenDynamic, adLockPessimistic
        Dim j As Integer
        While Not rs.EOF
            With rs
                If rs!Emp_Code = Text1.Text Or rs!Emp_Name = Text1.Text Then
                    MsgBox "Gross Salary of " & rs!Emp_Name & " is " & rs!Amt_Balance
                    j = 1
                End If
            End With
            rs.MoveNext
        Wend
        End If
        If j = 0 Then
            MsgBox "Record Not Found!!"
        End If
        rs.Close
    End If
End Sub

Private Sub Command2_Click()
    Form2.Show
    Unload Form3
End Sub

Private Sub Command3_Click()
    MsgBox "Email id: anand27krishn@gmail.com"
    MsgBox "Facebook: www.facebook.com/anand27krishn"
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    str = "select * from mytable1"
    cn.Open "provider = Microsoft.jet.oledb.4.0; data source=F:\VB Project\Krishn\mydatabase1.mdb"
End Sub
