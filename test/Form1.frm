VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1944
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   LinkTopic       =   "Form1"
   ScaleHeight     =   1944
   ScaleWidth      =   4788
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctxOwnerDrawCombo ctxOwnerDrawCombo1 
      Height          =   336
      Left            =   252
      TabIndex        =   1
      Top             =   336
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

Private Sub Form_Load()
    Set ctxOwnerDrawCombo1.Font = SystemIconFont
    ctxOwnerDrawCombo1.RegisterExtension New cPrintersCombo
End Sub

Private Sub Command1_Click()
    MsgBox "User wants to print to " & ctxOwnerDrawCombo1.Text & " w/ index " & ctxOwnerDrawCombo1.ListIndex, vbExclamation
End Sub

