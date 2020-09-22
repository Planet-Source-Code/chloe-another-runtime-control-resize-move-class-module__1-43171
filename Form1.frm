VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
' Just used MoveMove for effect....you can pass
' a control in anytime.
'*************************************************
Option Explicit
Public obj As CRESIZECTL
Private Sub Form_Load()
    Set obj = New CRESIZECTL
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  obj.DisallowResize
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    obj.AllowResize Picture1
End Sub
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    obj.AllowResize Text1
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    obj.AllowResize Command1
End Sub

