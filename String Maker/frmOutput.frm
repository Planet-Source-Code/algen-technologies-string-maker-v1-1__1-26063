VERSION 5.00
Begin VB.Form frmOutput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Output String"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7470
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   4095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText txtOutput.Text


End Sub

Private Sub Form_Load()

End Sub

Private Sub txtOutput_GotFocus()
Command1.SetFocus

End Sub
