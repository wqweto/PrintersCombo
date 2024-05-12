VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1944
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4836
   LinkTopic       =   "Form1"
   ScaleHeight     =   1944
   ScaleWidth      =   4836
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Disabled"
      Height          =   192
      Left            =   336
      TabIndex        =   2
      Top             =   168
      Width           =   2112
   End
   Begin Project1.ctxOwnerDrawCombo ctxOwnerDrawCombo1 
      Height          =   336
      Left            =   336
      TabIndex        =   1
      Top             =   588
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   593
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   516
      Left            =   2940
      TabIndex        =   0
      Top             =   1176
      Width           =   1356
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    ctxOwnerDrawCombo1.Extension.Enabled = (Check1.Value = vbUnchecked)
End Sub

Private Sub Form_Load()
    Set ctxOwnerDrawCombo1.Font = SystemIconFont
    ctxOwnerDrawCombo1.RegisterExtension New cPrintersCombo
End Sub

Private Sub Command1_Click()
    MsgBox "User wants to print to " & ctxOwnerDrawCombo1.Text & " w/ index " & ctxOwnerDrawCombo1.ListIndex, vbExclamation
End Sub

