VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      Caption         =   "Net Banking"
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Debit Card"
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Credit Card"
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   75
   End
   Begin VB.Label Label2 
      Caption         =   "Select Payment type:"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Online Payment  "
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
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()
Form2.Show
End Sub

Private Sub Option2_Click()
Form4.Show
Form2.Hide
End Sub

Private Sub Option3_Click()
Form5.Show
Form2.Hide
End Sub

Private Sub Option4_Click()
Form3.Show
Form2.Hide
End Sub
