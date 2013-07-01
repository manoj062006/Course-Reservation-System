VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form10"
   ScaleHeight     =   5265
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5280
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Student ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "OCRS LOGIN :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   6495
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Msg As String



Private Sub Command1_Click()

Do Until rs.EOF

If rs.Fields("login").Value = Text1.Text And rs.Fields("password").Value = Text2.Text Then
MsgBox ("Login Successful !")
Form10.Hide
Form12.Show
Exit Sub
Else
rs.MoveNext
End If
Loop
Msg = MsgBox("Invalid password, try again!", vbOKCancel)
rs.MoveFirst
If (Msg = 1) Then
Form10.Show
Else
End
End If

End Sub

Private Sub Form_Load()

cn.Open "ocrs", "scott", "tiger"
rs.Open "select * from login", cn, adOpenDynamic, adLockOptimistic

End Sub
