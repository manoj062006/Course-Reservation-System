VERSION 5.00
Begin VB.Form Form11 
   ClientHeight    =   4536
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   9948
   LinkTopic       =   "Form11"
   ScaleHeight     =   4536
   ScaleWidth      =   9948
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Proceed"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Your Login ID  and Password has been sent to your email ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   8295
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form10.Show
Form11.Hide
End Sub

