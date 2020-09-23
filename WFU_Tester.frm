VERSION 5.00
Begin VB.Form WFU_Tester 
   Caption         =   "Hell Yeah!"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox rText 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   2
      Text            =   "11"
      Top             =   240
      Width           =   465
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Unload Tray form and return to normal size"
      Height          =   435
      Left            =   1230
      TabIndex        =   1
      Top             =   1800
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Resize (Final)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Top             =   420
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   $"WFU_Tester.frx":0000
      Height          =   825
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   4155
   End
   Begin VB.Label Label1 
      Caption         =   "Number to Resize"
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   1305
   End
End
Attribute VB_Name = "WFU_Tester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()
ResizeTray rText.Text, 2
End Sub

Private Sub Command2_Click()
ResizeTray rText.Text, 1
Load WFU_Tray
Encapsulate_Tray
End Sub

