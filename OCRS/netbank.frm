VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5364
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11820
   LinkTopic       =   "Form3"
   ScaleHeight     =   10260
   ScaleWidth      =   18984
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo7 
      Height          =   288
      Left            =   4200
      TabIndex        =   8
      Top             =   1920
      Width           =   3492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   372
      Left            =   2640
      TabIndex        =   7
      Top             =   4440
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Payment"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   2880
      Width           =   3495
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
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Transaction ID"
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
      Left            =   840
      TabIndex        =   2
      Top             =   2880
      Width           =   2655
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
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Net Banking"
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
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox ("Hot Payment Successful")
Form9.Show
Form3.Hide
End Sub

Private Sub Command2_Click()
Form2.Show
Form3.Hide
End Sub

Private Sub Form_Load()
Combo7.AddItem ("BARCLAYS")
Combo7.AddItem ("Royal Bank of Scotland")
Combo7.AddItem ("Swiss Bank")
Combo7.AddItem ("CITI Bank")
Combo7.AddItem ("Deutsche Bank")
End Sub
