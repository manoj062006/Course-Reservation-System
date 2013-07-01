VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form12"
   ScaleHeight     =   7230
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Proceed to Course Selection"
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   6120
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   2055
      Left            =   5640
      TabIndex        =   6
      Top             =   2880
      Width           =   4095
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Mechanical"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   5040
      Width           =   3615
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Electronics and Instrumentation"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   4320
      Width           =   3615
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Electrical and Electronics"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3600
      Width           =   3615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Electronic and Communication"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Computer Science"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Course Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Course Catalogue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form7.Show
Form12.Hide
End Sub

Private Sub Option1_Click()
If Option1.Enabled = True Then
Text2.Text = "Computer Science and Engineering: In this course student will get to learrn about computer....."
End If
End Sub

Private Sub Option2_Click()

If Option2.Enabled = True Then
Text2.Text = "Electronics and Communication Engineering:In this course student will get to learrn about Electronics....."

End If
End Sub

Private Sub Option3_Click()

If Option3.Enabled = True Then
Text2.Text = "Electrical and Electronics Engineering:In this course student will get to learrn about Electrical....."
End If
End Sub

Private Sub Option4_Click()

If Option4.Enabled = True Then
Text2.Text = "Electronics and Instrumentation Engineering: In this course student will get to learrn about Instrumentation....."
End If
End Sub

Private Sub Option5_Click()

If Option5.Enabled = True Then
Text2.Text = "Mechanical Engineering : In this course student will get to learrn about Mechanical....."
End If
End Sub
