VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text22 
      BackColor       =   &H0080FFFF&
      DataField       =   "Amt_Balance"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   10680
      Width           =   3375
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H0080FFFF&
      DataField       =   "Emp_Code"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   7080
      Width           =   3375
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H0080FFFF&
      DataField       =   "Emp_Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   7680
      Width           =   3375
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H0080FFFF&
      DataField       =   "Contact"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   8280
      Width           =   3375
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H0080FFFF&
      DataField       =   "Father_Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   8880
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080FFFF&
      DataField       =   "Mother_Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   9480
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FFFF&
      DataField       =   "Bank_Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   10050
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< Back to Login Page"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   0
      Width           =   3615
   End
   Begin VB.Data Data1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Employee's Record"
      Connect         =   "Access"
      DatabaseName    =   "F:\VB Project\Krishn\MYDATABASE1.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "mytable1"
      Top             =   6240
      Width           =   3375
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4560
      Width           =   3975
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   3360
      Width           =   3735
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   960
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7695
      ScaleWidth      =   8655
      TabIndex        =   2
      Top             =   5400
      Width           =   8655
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   4680
         Width           =   4695
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   4080
         Width           =   4695
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3240
         TabIndex        =   21
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   2880
         Width           =   4695
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2280
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3240
         TabIndex        =   18
         Top             =   1680
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   1080
         Width           =   4695
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00404080&
         Caption         =   "ADD RECORD"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5400
         Width           =   8655
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Note : All fields are required]"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   2355
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Balance :"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   4680
         Width           =   3060
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name :"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   4080
         Width           =   1980
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Name :"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   2700
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name :"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   2700
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No. : "
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   2520
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee's Name : "
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   3240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Code :"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2700
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":: Make New Entry or Edit Record ::"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   8655
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   4
         X1              =   0
         X2              =   8640
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000040C0&
         BorderWidth     =   7
         Index           =   0
         X1              =   0
         X2              =   8640
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   8655
      TabIndex        =   1
      Top             =   480
      Width           =   8655
      Begin VB.TextBox Text19 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   39
         Top             =   120
         Width           =   3735
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Find Total Number of Records"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3840
         UseMaskColor    =   -1  'True
         Width           =   7695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Calculate Salary"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   7695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Delete Record"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         UseMaskColor    =   -1  'True
         Width           =   7695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Search Record"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   7695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Employee Code:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   420
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3900
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Salary"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9960
      TabIndex        =   49
      Top             =   10680
      Width           =   2805
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of the Bank"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9960
      TabIndex        =   48
      Top             =   10080
      Width           =   2640
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mother's Name"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9960
      TabIndex        =   47
      Top             =   9480
      Width           =   2145
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9960
      TabIndex        =   46
      Top             =   8880
      Width           =   2145
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9960
      TabIndex        =   45
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Name"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9960
      TabIndex        =   44
      Top             =   7680
      Width           =   2475
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9960
      TabIndex        =   43
      Top             =   7080
      Width           =   2145
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   9000
      TabIndex        =   42
      Top             =   6240
      Width           =   1920
   End
   Begin VB.Image Image2 
      Height          =   4200
      Left            =   17520
      Picture         =   "FRwelcome_admin.frx":0000
      Top             =   2040
      Width           =   6720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Name :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   2
      Left            =   9000
      TabIndex        =   38
      Top             =   1560
      Width           =   2565
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "The Requested Record is >>"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8760
      TabIndex        =   30
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Balance :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   1
      Left            =   9000
      TabIndex        =   29
      Top             =   4560
      Width           =   2430
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Of the Bank :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   1
      Left            =   9000
      TabIndex        =   28
      Top             =   3960
      Width           =   2670
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mother's Name :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   1
      Left            =   9000
      TabIndex        =   27
      Top             =   3360
      Width           =   2235
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   1
      Left            =   9000
      TabIndex        =   26
      Top             =   2760
      Width           =   2100
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No. : "
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   1
      Left            =   9000
      TabIndex        =   25
      Top             =   2160
      Width           =   1845
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   1
      Left            =   9000
      TabIndex        =   24
      Top             =   960
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   8640
      Left            =   8640
      Picture         =   "FRwelcome_admin.frx":426D
      Top             =   6240
      Width           =   15360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Welcome Admin : 12345"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   23775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim str, s1, s2, s3, s4, s5, s6, s7 As String

Private Sub Command1_Click()
    rs.Open str, cn, adOpenDynamic, adLockPessimistic
    Dim i As Integer
    i = 0
    While Not rs.EOF
        With rs
            If rs!Emp_Code = Text19.Text Or rs!Emp_Name = Text19.Text Then
                s1 = rs!Emp_Name
                s2 = rs!Emp_Code
                s3 = rs!Father_Name
                s4 = rs!Mother_Name
                s5 = rs!Bank_Name
                s6 = rs!Amt_Balance
                s7 = rs!Contact
                Text11.Text = s1
                Text10.Text = s2
                Text15.Text = s3
                Text16.Text = s4
                Text17.Text = s5
                Text18.Text = s6
                Text14.Text = s7
                i = 1
            End If
        End With
        rs.MoveNext
    Wend
    If i = 0 Then
        Text11.Text = ""
        Text10.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        Text14.Text = ""
        MsgBox "Record Not Found!!"
    End If
    rs.Close
End Sub

Private Sub Command2_Click()
    rs.Open str, cn, adOpenDynamic, adLockPessimistic
    Dim i As Integer
    i = 0
    While Not rs.EOF
        With rs
            If rs!Emp_Code = Text19.Text Then
                .Delete
                MsgBox "Record Found and Deleted"
                i = 1
            End If
        End With
        rs.MoveNext
    Wend
    If i = 0 Then MsgBox "Record Not Found!!"
    rs.Close
End Sub

Private Sub Command3_Click()
    rs.Open str, cn, adOpenDynamic, adLockPessimistic
    Dim j As Integer
    j = 0
    While Not rs.EOF
        With rs
            If rs!Emp_Code = Text19.Text Then
                MsgBox "Employee's Salary is is Rs " & rs!Amt_Balance
                i = 1
            End If
        End With
        rs.MoveNext
    Wend
    If i = 0 Then MsgBox "Record Not Found!!"
    rs.Close
End Sub

Private Sub Command5_Click()
    rs.Open str, cn, adOpenDynamic, adLockPessimistic
    Dim k As Integer
    k = 0
    While Not rs.EOF
        k = k + 1
        rs.MoveNext
    Wend
    rs.Close
    MsgBox "Total Numbers of Record in Database is " & k
End Sub

Private Sub Command6_Click()
    rs.Open str, cn, adOpenDynamic, adLockPessimistic
    While Not rs.EOF
        If rs!Emp_Code = Text1.Text Then
            MsgBox "Record with the same Employee Code is already in the database"
            GoTo OUTSIDE
        End If
        rs.MoveNext
    Wend
    rs.AddNew
    rs!Emp_Name = Text2.Text
    rs!Emp_Code = Text1.Text
    rs!Father_Name = Text6.Text
    rs!Mother_Name = Text7.Text
    rs!Bank_Name = Text8.Text
    rs!Amt_Balance = Text9.Text
    rs!Contact = Text5.Text
    rs.Update
    rs.Close
    Data1.Refresh
    MsgBox "Record Added Successfully"
OUTSIDE:
    Text2.Text = ""
    Text1.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text5.Text = ""
End Sub

Private Sub Command7_Click()
    Form2.Show
    Unload Form5
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    str = "select * from mytable1"
    cn.Open "provider = Microsoft.jet.oledb.4.0; data source=F:\VB Project\Krishn\mydatabase1.mdb"
End Sub

