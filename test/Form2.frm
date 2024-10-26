VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5532
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   5532
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   516
      Left            =   2940
      TabIndex        =   1
      Top             =   1176
      Width           =   1356
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Disabled"
      Height          =   192
      Left            =   336
      TabIndex        =   0
      Top             =   168
      Width           =   2112
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ctxOwnerDrawCombo1 As ctxOwnerDrawCombo
Attribute ctxOwnerDrawCombo1.VB_VarHelpID = -1

Private Sub Check1_Click()
    ctxOwnerDrawCombo1.Extension.Enabled = (Check1.Value = vbUnchecked)
End Sub

Private Sub Form_Load()
    With Controls.Add("Project1.ctxOwnerDrawCombo", "ctxOwnerDrawCombo1")
        .Move 336, 588, 3540
        .Visible = True
        Set ctxOwnerDrawCombo1 = .object
    End With
    Set ctxOwnerDrawCombo1.Font = SystemIconFont
    ctxOwnerDrawCombo1.RegisterExtension New cPrintersCombo
End Sub

Private Sub Command1_Click()
    MsgBox "User wants to print to " & ctxOwnerDrawCombo1.Text & " w/ index " & ctxOwnerDrawCombo1.ListIndex, vbExclamation
End Sub

