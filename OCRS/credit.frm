VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11796
   LinkTopic       =   "Form4"
   ScaleHeight     =   5280
   ScaleWidth      =   11796
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo8 
      Height          =   288
      Left            =   4320
      TabIndex        =   8
      Top             =   1680
      Width           =   3492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   492
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Payment"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Credit Card Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Select bank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Credit Card Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Transaction Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
MsgBox ("Hot Payment Successful")
Form9.Show
Form4.Hide
End Sub

Private Sub Command2_Click()
Form2.Show
Form4.Hide
End Sub

Private Sub Form_Load()
Combo8.AddItem ("BARCLAYS")
Combo8.AddItem ("Royal Bank of Scotland")
Combo8.AddItem ("Swiss Bank")
Combo8.AddItem ("CITI Bank")
Combo8.AddItem ("Deutsche Bank")
End Sub
