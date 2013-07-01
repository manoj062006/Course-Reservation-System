VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   LinkTopic       =   "Form7"
   ScaleHeight     =   5235
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10920
      TabIndex        =   6
      Top             =   -120
      Width           =   150
   End
   Begin VB.ComboBox Combo5 
      Height          =   288
      Left            =   4440
      TabIndex        =   5
      Top             =   2880
      Width           =   5172
   End
   Begin VB.ComboBox Combo4 
      Height          =   288
      Left            =   4440
      TabIndex        =   4
      Top             =   1920
      Width           =   5172
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Availablity"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "College Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "WELCOME TO O.C.R.S."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim cid As Integer

Private Sub Command1_Click()
rs.AddNew
rs.MoveFirst
cid = cid + 1
rs.Fields(0) = Combo4.Text
rs.Fields(1) = Combo5.Text
rs.Fields(2) = cid
If Combo4.Text = "AIHT" And Combo5.Text = "CSE" Then
rs.Fields(3) = rs1.Fields(1).Value
Else
If Combo4.Text = "AIHT" And Combo5.Text = "ECE" Then
rs.Fields(3) = rs1.Fields(2).Value
Else
If Combo4.Text = "AIHT" And Combo5.Text = "MECH" Then
rs.Fields(3) = rs1.Fields(3).Value
Else
If Combo4.Text = "AIHT" And Combo5.Text = "EEE" Then
rs.Fields(3) = rs1.Fields(4).Value
Else
If Combo4.Text = "Harvard" And Combo5.Text = "CSE" Then
rs.Fields(3) = rs1.Fields(1).Value
Else
If Combo4.Text = "Harvard" And Combo5.Text = "ECE" Then
rs.Fields(3) = rs1.Fields(2).Value
Else
If Combo4.Text = "Harvard" And Combo5.Text = "MECH" Then
rs.Fields(3) = rs1.Fields(3).Value
Else
If Combo4.Text = "Harvard" And Combo5.Text = "EEE" Then
rs.Fields(3) = rs1.Fields(4).Value
Else
If Combo4.Text = "Stanford" And Combo5.Text = "CSE" Then
rs.Fields(3) = rs1.Fields(1).Value
Else
If Combo4.Text = "Stanford" And Combo5.Text = "ECE" Then
rs.Fields(3) = rs1.Fields(2).Value
Else
If Combo4.Text = "Stanford" And Combo5.Text = "MECH" Then
rs.Fields(3) = rs1.Fields(3).Value
Else
If Combo4.Text = "Stanford" And Combo5.Text = "EEE" Then
rs.Fields(3) = rs1.Fields(4).Value
Else
If Combo4.Text = "Carnegie" And Combo5.Text = "CSE" Then
rs.Fields(3) = rs1.Fields(1).Value
Else
If Combo4.Text = "Carnegie" And Combo5.Text = "ECE" Then
rs.Fields(3) = rs1.Fields(2).Value
Else
If Combo4.Text = "Carnegie" And Combo5.Text = "MECH" Then
rs.Fields(3) = rs1.Fields(3).Value
Else
If Combo4.Text = "Carnegie" And Combo5.Text = "EEE" Then
rs.Fields(3) = rs1.Fields(4).Value
Else
If Combo4.Text = "Berkeley" And Combo5.Text = "CSE" Then
rs.Fields(3) = rs1.Fields(1).Value
Else
If Combo4.Text = "Berkeley" And Combo5.Text = "ECE" Then
rs.Fields(3) = rs1.Fields(2).Value
Else
If Combo4.Text = "Berkeley" And Combo5.Text = "MECH" Then
rs.Fields(3) = rs1.Fields(3).Value
Else
If Combo4.Text = "Berkeley" And Combo5.Text = "EEE" Then
rs.Fields(3) = rs1.Fields(4).Value
Else: MsgBox "Invalid Selection"
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
rs.Update
MsgBox ("Course Available !")
Form8.Show
Form7.Hide


End Sub

Private Sub Form_Load()
cn.Open "ocrs", "scott", "tiger"
rs.Open "select * from course", cn, adOpenDynamic, adLockOptimistic
rs1.Open "select * from avail", cn, adOpenDynamic, adLockOptimistic

Combo4.AddItem ("AIHT")
Combo4.AddItem ("Harvard")
Combo4.AddItem ("Stanford")
Combo4.AddItem ("Carnegie")
Combo4.AddItem ("Berkeley")
Combo5.AddItem ("CSE")
Combo5.AddItem ("ECE")
Combo5.AddItem ("MECH")
Combo5.AddItem ("EEE")
End Sub
