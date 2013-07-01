VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14445
   ScaleHeight     =   8415
   ScaleWidth      =   14445
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5280
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1080
      TabIndex        =   7
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "LOGIN NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1080
      TabIndex        =   6
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "LAST NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "AGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "FIRST NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "STUDENT REGISTRATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim roll As Integer


Private Sub Command1_Click()
rs.AddNew
roll = roll + 1
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = roll + rs.Fields(5).Value
rs.Update
rs1.AddNew
rs1.Fields(0) = Text4.Text
rs1.Fields(1) = Text5.Text
rs1.Update

MsgBox ("Registration Successful !!"), vbInformation
Form1.Hide
Form11.Show
End Sub

Private Sub Form_Load()
cn.Open "ocrs", "scott", "tiger"
rs.Open "select * from stud", cn, adOpenDynamic, adLockOptimistic
rs1.Open "select * from login", cn, adOpenDynamic, adLockOptimistic

MsgBox ("Connected to The Database")
'If conn.State = 0 Then
'conn.Open "new"
'End If
End Sub

