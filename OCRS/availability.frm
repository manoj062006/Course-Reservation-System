VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form8"
   ScaleHeight     =   5805
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm and proceed to payment"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Seats Available"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Availability"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Course"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "College"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cn As New ADODB.Connection
Private Sub Command1_Click()
rs.MoveFirst
If Text1.Text = "AIHT" And Text2.Text = "CSE" Then
rs.Fields(1).Value = rs.Fields(1).Value - 1
GoTo x
Else
If Text1.Text = "AIHT" And Text2.Text = "ECE" Then
rs.Fields(2).Value = rs.Fields(2).Value - 1
GoTo x
Else
If Text1.Text = "AIHT" And Text2.Text = "MECH" Then
rs.Fields(3).Value = rs.Fields(3).Value - 1
Else
If Text1.Text = "AIHT" And Text2.Text = "EEE" Then
rs.Fields(4).Value = rs.Fields(4).Value - 1
GoTo x
Else
If Text1.Text = "Harvard" And Text2.Text = "CSE" Then
rs.Fields(1).Value = rs.Fields(1).Value - 1
GoTo x
Else
If Text1.Text = "Harvard" And Text2.Text = "ECE" Then
rs.Fields(2).Value = rs.Fields(2).Value - 1
GoTo x
Else
If Text1.Text = "Harvard" And Text2.Text = "MECH" Then
rs.Fields(3).Value = rs.Fields(3).Value - 1
Else
If Text1.Text = "Harvard" And Text2.Text = "EEE" Then
rs.Fields(4).Value = rs.Fields(4).Value - 1
GoTo x
Else
If Text1.Text = "Stanford" And Text2.Text = "CSE" Then
rs.Fields(1).Value = rs.Fields(1).Value - 1
GoTo x
Else
If Text1.Text = "Stanford" And Text2.Text = "ECE" Then
rs.Fields(2).Value = rs.Fields(2).Value - 1
GoTo x
Else
If Text1.Text = "Stanford" And Text2.Text = "MECH" Then
rs.Fields(3).Value = rs.Fields(3).Value - 1
Else
If Text1.Text = "Stanford" And Text2.Text = "EEE" Then
rs.Fields(4).Value = rs.Fields(4).Value - 1
GoTo x
Else
If Text1.Text = "Carnegie" And Text2.Text = "CSE" Then
rs.Fields(1).Value = rs.Fields(1).Value - 1
GoTo x
Else
If Text1.Text = "Carnegie" And Text2.Text = "ECE" Then
rs.Fields(2).Value = rs.Fields(2).Value - 1
GoTo x
Else
If Text1.Text = "Carnegie" And Text2.Text = "MECH" Then
rs.Fields(3).Value = rs.Fields(3).Value - 1
Else
If Text1.Text = "Carnegie" And Text2.Text = "EEE" Then
rs.Fields(4).Value = rs.Fields(4).Value - 1
GoTo x
Else
If Text1.Text = "Berkeley" And Text2.Text = "CSE" Then
rs.Fields(1).Value = rs.Fields(1).Value - 1
GoTo x
Else
If Text1.Text = "Berkeley" And Text2.Text = "ECE" Then
rs.Fields(2).Value = rs.Fields(2).Value - 1
GoTo x
Else
If Text1.Text = "Berkeley" And Text2.Text = "MECH" Then
rs.Fields(3).Value = rs.Fields(3).Value - 1
Else
If Text1.Text = "Berkeley" And Text2.Text = "EEE" Then
rs.Fields(4).Value = rs.Fields(4).Value - 1
GoTo x

Else
If Text3.Text = 0 Then
MsgBox "Course Unavailable"
Else
MsgBox "Unexpected Error"

x:
rs.Update

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
Form8.Hide
Form2.Show
End Sub

Private Sub Form_Load()
cn.Open "ocrs", "scott", "tiger"
rs.Open "select * from avail", cn, adOpenDynamic, adLockOptimistic
rs1.Open "select * from course ", cn, adOpenDynamic, adLockOptimistic

rs1.MoveLast
Text1.Text = rs1.Fields(0)
Text2.Text = rs1.Fields(1)
Text3.Text = rs1.Fields(3)
End Sub

