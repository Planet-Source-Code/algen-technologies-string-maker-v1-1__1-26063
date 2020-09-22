VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "String Input"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6765
   Begin VB.CommandButton Command3 
      Caption         =   "Copy Formated Output"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Paste From Clipboard"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtInput 
      Height          =   2535
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type the text to be formated here"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
txtInput.Text = Clipboard.GetText

End Sub

Private Sub Command2_Click()
InputIsDirty = False
txtInput.Text = ""


End Sub

Private Sub txtInput_Change()
Dim Comillas As String
Dim CR1 As String
Dim CR2 As String
Dim NewLineVar As String
Dim NewLineSym As String
Dim WorkString As String

Comillas = Chr(34) & " & chr(34) & " & Chr(34)
CR1 = Chr(34) & " & vbNewLine " & vbNewLine
CR2 = Chr(34) & " & vbNewLine _" & vbNewLine
NewLineVar = RecipientVar & " = " & RecipientVar & " & " & Chr(34)
NewLineSym = "& " & Chr(34)
WorkString = ""


If txtInput.Text = "" Then Exit Sub

'Start string formating
'1.- setting the initial string
If RecipientVar = "" Then RecipientVar = "Var"

WorkString = RecipientVar & " = " & Chr(34)

'2.- Check the string for newlines
For x = 1 To Len(txtInput.Text)
'Debug.Print Mid(txtInput.Text, x, 1), Asc(Mid(txtInput.Text, x, 1))



Select Case Mid(txtInput.Text, x, 1)

    Case Chr(34)
        WorkString = WorkString & Comillas
        
    Case Chr(13)
        If BrakeSymbol = False Then
        WorkString = WorkString & CR1 & NewLineVar
        Else
        WorkString = WorkString & CR2 & NewLineSym
        End If
    
    Case Chr(10)
        ' do nothin
        ' Just to get rid of line feed
        
        
    Case Else
        WorkString = WorkString & Mid(txtInput.Text, x, 1)
        
End Select

Next x


'3.- Close the string
WorkString = WorkString & Chr(34)

'4- Update output window

frmOutput.txtOutput.Text = WorkString


End Sub

