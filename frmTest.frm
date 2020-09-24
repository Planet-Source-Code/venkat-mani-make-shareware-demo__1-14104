VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Make Shareware"
   ClientHeight    =   4056
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6396
   LinkTopic       =   "Form1"
   ScaleHeight     =   4056
   ScaleWidth      =   6396
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   372
      Left            =   2640
      TabIndex        =   0
      Top             =   3480
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6252
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op4 As Integer
Dim op1 As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim lresult As Long
Dim sKeyValue As String
Dim op3 As Integer
Dim op As String
lresult = GetRegValue("\Software\venky\shareware", "uses", sKeyValue)
op = sKeyValue
op3 = Val(op)
lresult = GetRegValue("\Software\venky\shareware", "value", sKeyValue)
op1 = sKeyValue
op4 = Val(op1)
If op4 <= op3 Then
Label1.Caption = "This Program has been used " & op1 & " times"
Dim op2 As String
op4 = op4 + 1
op2 = op4
lresult = SetRegValue("\Software\venky\shareware", "value", op2)
Else
MsgBox "Please Register to Use"
Unload frmTest
End If
End Sub

