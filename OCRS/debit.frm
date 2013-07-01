VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   5412
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   12780
   LinkTopic       =   "Form5"
   ScaleHeight     =   5412
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo9 
      Height          =   288
      Left            =   4560
      TabIndex        =   8
      Top             =   1680
      Width           =   3492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   372
      Left            =   3120
      TabIndex        =   7
      Top             =   4200
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Payment"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Debit Card Payment"
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
      Left            =   1200
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
      Left            =   1200
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
      Left            =   1200
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
      Left            =   1200
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox ("Hot Payment Successful")
Form9.Show
Form5.Hide
End Sub

Private Sub Command2_Click()
Form2.Show
Form5.Hide
End Sub

Private Sub Form_Load()
Combo9.AddItem ("BARCLAYS")
Combo9.AddItem ("Royal Bank of Scotland")
Combo9.AddItem ("Swiss Bank")
Combo9.AddItem ("CITI Bank")
Combo9.AddItem ("Deutsche Bank")
End Sub
