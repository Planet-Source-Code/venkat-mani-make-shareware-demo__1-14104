VERSION 5.00
Begin VB.Form frmregister 
   Caption         =   "Test Shareware Demo-venky_dude"
   ClientHeight    =   2532
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4068
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2532
   ScaleWidth      =   4068
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   492
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   492
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1452
   End
   Begin VB.Label Label3 
      Caption         =   "Please read Readme.txt file"
      Height          =   372
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label2 
      Caption         =   "Reg no:"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "Trial period expired . Register to continue"
      Height          =   372
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3012
   End
End
Attribute VB_Name = "frmregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please Enter Serial no"
Exit Sub
End If
If Text1.Text = "a1234" Then
Dim lresult As Long
lresult = SetRegValue("\Software\venky\shareware", "lock", "false")
If lresult = o Then MsgBox "Registration Successful"
Unload Me
frmmain.Show
Else: MsgBox "Bad Serial No"
End If
End Sub

Private Sub Command2_Click()
Unload Me
Unload frmmain
End Sub
