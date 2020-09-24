VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Test ShareWare Demo-venky_dude"
   ClientHeight    =   3432
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5796
   LinkTopic       =   "Form1"
   ScaleHeight     =   3432
   ScaleWidth      =   5796
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   372
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   1692
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
      Height          =   732
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   5412
   End
End
Attribute VB_Name = "frmmain"
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
Dim op9 As Date
Dim op8 As Date
Dim op10 As Date
Dim op11 As String
lresult = GetRegValue("\Software\venky\shareware", "lock", sKeyValue)
If sKeyValue = "false" Then
Label1.Caption = "Registered Copy'"
Exit Sub
Exit Sub
End If
lresult = GetRegValue("\Software\venky\shareware", "days", sKeyValue)
If Not sKeyValue = "" Then
op = sKeyValue
op9 = op
op8 = Date
lresult = GetRegValue("\Software\venky\shareware", "value", sKeyValue)
op11 = sKeyValue
op10 = op11
If op10 > op8 Then
frmregister.Show
Exit Sub
End If
If op8 < op9 Then
Label1.Caption = (op9 - op8) & " days left"
Exit Sub
Else
frmregister.Show
End If
Else
lresult = GetRegValue("\Software\venky\shareware", "uses", sKeyValue)
op = sKeyValue
op3 = Val(op)
If op3 = 0 Then
MsgBox "Make A Lock First"
Unload frmmain
End If
lresult = GetRegValue("\Software\venky\shareware", "value", sKeyValue)
op1 = sKeyValue
op4 = Val(op1)
If op4 <= op3 Then
Label1.Caption = "This Program has been run " & op1 & " times from a maximum of " & op3 & " times"
Dim op2 As String
op4 = op4 + 1
op2 = op4
lresult = SetRegValue("\Software\venky\shareware", "value", op2)
Else
frmregister.Show
End If
End If
End Sub

