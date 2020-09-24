VERSION 5.00
Begin VB.Form frmshare 
   Caption         =   "Shareware Lock Demo -venky_dude"
   ClientHeight    =   4056
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6396
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4056
   ScaleWidth      =   6396
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   492
      Left            =   3720
      TabIndex        =   14
      Top             =   3480
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Height          =   492
      Left            =   1440
      TabIndex        =   13
      Top             =   3480
      Width           =   1452
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1812
      Left            =   4080
      TabIndex        =   9
      Top             =   1680
      Width           =   2172
      Begin VB.OptionButton Option2 
         Caption         =   "10 Uses"
         Height          =   372
         Index           =   5
         Left            =   0
         TabIndex        =   12
         Top             =   1320
         Width           =   1572
      End
      Begin VB.OptionButton Option2 
         Caption         =   "5 Uses"
         Height          =   372
         Index           =   4
         Left            =   0
         TabIndex        =   11
         Top             =   840
         Width           =   1572
      End
      Begin VB.OptionButton Option2 
         Caption         =   "1 Use"
         Height          =   372
         Index           =   3
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1572
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1692
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   2172
      Begin VB.OptionButton Option2 
         Caption         =   "1 day"
         Height          =   372
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1572
      End
      Begin VB.OptionButton Option2 
         Caption         =   "15 days"
         Height          =   372
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   840
         Width           =   1572
      End
      Begin VB.OptionButton Option2 
         Caption         =   "30 days"
         Height          =   372
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   1320
         Width           =   1572
      End
   End
   Begin VB.OptionButton Option1 
      Height          =   372
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Top             =   1320
      Width           =   372
   End
   Begin VB.OptionButton Option1 
      Height          =   372
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   372
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Expire after uses"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   2532
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Expire after days"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Type of Shareware Lock"
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
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4812
   End
End
Attribute VB_Name = "frmshare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim lresult As Long
Dim sKey As String
Dim sSubkey As String
Dim sSubkey1 As String
Dim sKeyValue1 As String
sKey = "\Software\venky\shareware"
frmMain.Command2.Enabled = True
If Option1(0).Value = True Then
sSubkey = "days"
If Option2(0).Value = True Then sKeyValue1 = (Date + 30)
If Option2(1).Value = True Then sKeyValue1 = (Date + 15)
If Option2(2).Value = True Then sKeyValue1 = (Date + 1)
lresult = SetRegValue(sKey, sSubkey, sKeyValue1)
lresult = SetRegValue(sKey, "value", Date)
lresult = SetRegValue(sKey, "lock", "true")
End If
If Option1(1).Value = True Then
sSubkey = "uses"
If Option2(3).Value = True Then sKeyValue1 = 1
If Option2(4).Value = True Then sKeyValue1 = 5
If Option2(5).Value = True Then sKeyValue1 = 10
lresult = SetRegValue(sKey, sSubkey, sKeyValue1)
lresult = SetRegValue(sKey, "value", "1")
lresult = SetRegValue(sKey, "lock", "true")
End If
MsgBox "Lock Made"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Option1(0).Value = True Then
Option2(3).Value = False
Option2(4).Value = False
Option2(5).Value = False
End If
CreateRegKey ("\Software\venky\shareware")
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
Option2(3).Value = False
Option2(4).Value = False
Option2(5).Value = False
End If
If Option1(1).Value = True Then
Option2(0).Value = False
Option2(1).Value = False
Option2(2).Value = False
End If
End Sub

Private Sub Option2_Click(Index As Integer)
If Option2(2).Value = True Then
Option1(1).Value = False
Option1(0).Value = True
End If
If Option2(1).Value = True Then
Option1(1).Value = False
Option1(0).Value = True
End If
If Option2(0).Value = True Then
Option1(1).Value = False
Option1(0).Value = True
End If
If Option2(3).Value = True Then
Option1(0).Value = False
Option1(1).Value = True
End If
If Option2(4).Value = True Then
Option1(0).Value = False
Option1(1).Value = True
End If
If Option2(5).Value = True Then
Option1(0).Value = False
Option1(1).Value = True
End If
End Sub
