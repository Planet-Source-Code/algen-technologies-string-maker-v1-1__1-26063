VERSION 5.00
Begin VB.Form frmSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set-Up Panel"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5535
   Begin VB.TextBox txtDataObj 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtRecVar 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
   Begin VB.CheckBox chkBrakeLine 
      Caption         =   "Breake formated text lines at VbNewLine"
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CheckBox chkBrkSymbol 
      Caption         =   "Use add string symbol instead of ""Var = Var +...""  *"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "DataBase Object  Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "* Does not apply to ""SQL Create table"" String Maker."
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Recipient Variable Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkBrkSymbol_Click()
BrakeSymbol = chkBrkSymbol.Value

End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtDataObj_Change()
DataObject = txtDataObj.Text

End Sub

Private Sub txtRecVar_Change()
For x = 1 To Len(txtRecVar)
   If Mid(txtRecVar, x, 1) = " " Then
     MsgBox "Var Name Cant Have Spaces"
     txtRecVar = RecipientVar
     Exit Sub
   End If
Next x

RecipientVar = txtRecVar

     
End Sub
