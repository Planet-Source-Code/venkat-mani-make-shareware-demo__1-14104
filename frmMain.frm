VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Shareware Lock  Demo -venky_dude"
   ClientHeight    =   3828
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5784
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3828
   ScaleWidth      =   5784
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   492
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Lock"
      Height          =   492
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Lock"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Please Send Your suggestions to venky_dude@yahoo.com . Check out my website at http://www.geocities.com/venky_dude"
      Height          =   852
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   4692
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmshare.Show

End Sub

Private Sub Command2_Click()
Dim lresult As Long
lresult = DeleteRegKey("\Software\venky", "value")
lresult = DeleteRegKey("\Software\venky", "days")
lresult = DeleteRegKey("\Software\venky", "uses")
lresult = DeleteRegKey("\Software\venky", "lock")
lresult = DeleteRegKey("\Software", "venky")
End Sub

Private Sub Command3_Click()
Unload Me
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lresult             As Long

' Remove the test data from the registry

End Sub
