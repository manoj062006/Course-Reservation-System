VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form9"
   ScaleHeight     =   6525
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
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
      Left            =   7560
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Enrollment Number"
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
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "First Name"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "College Selected"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Course Selected"
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
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Last Name"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Reservation Details:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Thanks for registering !@!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cn As New ADODB.Connection

Private Sub Command1_Click(Index As Integer)
Form9.Hide
End Sub

Private Sub Form_Load()
cn.Open "ocrs", "scott", "tiger"
rs.Open "select * from stud order by rollno", cn, adOpenDynamic, adLockOptimistic
rs1.Open "select * from course order by cid", cn, adOpenDynamic, adLockOptimistic


rs.MoveLast
  Text1.Text = rs.Fields(0)
  Text2.Text = rs.Fields(1)
  Text5.Text = rs.Fields(5)
rs1.MoveFirst
  Text3.Text = rs1.Fields(0)
  Text4.Text = rs1.Fields(1)
  

 

End Sub

