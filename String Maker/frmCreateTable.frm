VERSION 5.00
Begin VB.Form frmCreateTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Create Table"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdEndCreate 
      Caption         =   "End Create"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cboxDataType 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Text            =   "Select Type"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddField 
      Caption         =   "Add"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   5175
      Begin VB.TextBox txtFieldSize 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtFieldName 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Size:"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Field Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Field Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtTableName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5175
      Begin VB.Label Label1 
         Caption         =   "Table Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboxDataType_Click()
If cboxDataType.Text = DTArray(0) Or cboxDataType.Text = DTArray(13) Then
txtFieldSize.Enabled = True
Else
txtFieldSize.Enabled = False
End If

End Sub

Private Sub cboxDataType_Change()
'cboxDataType.Text = DTArray(0)
'txtFieldSize.Enabled = True



End Sub

Private Sub cboxDataType_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub cmdAddField_Click()
Dim again As String


If frmOutput.txtOutput.Text = "" Then
MsgBox "Must create a table first!"
cmdCreate.Enabled = True
Exit Sub
End If

If txtFieldName.Text = "" Then
MsgBox "Must give the field a name"
txtFieldName.SetFocus
Exit Sub
End If

If txtFieldSize.Enabled = True And IsNumeric(txtFieldSize) = False Then
MsgBox "Field Size must have numeric values only"
txtFieldSize.SetFocus
Exit Sub
End If

If txtFieldSize.Enabled = True And (CInt(txtFieldSize.Text) > 255 Or CInt(txtFieldSize.Text) < 1) = True Then
MsgBox "Must set a valid field size"
txtFieldSize.SetFocus
Exit Sub
End If

If cboxDataType.Text = "Select Type" Then
MsgBox "Must select a data type"
Exit Sub
End If



If txtFieldSize.Enabled = False Then
again = RecipientVar & " = " & RecipientVar & " & " & Chr(34) & "[" & txtFieldName.Text & "] " & cboxDataType.Text & "," & Chr(34) & vbNewLine
Else
again = RecipientVar & " = " & RecipientVar & " & " & Chr(34) & "[" & txtFieldName.Text & "] " & cboxDataType.Text & "(" & txtFieldSize.Text & ")," & Chr(34) & vbNewLine
End If

frmOutput.txtOutput.Text = frmOutput.txtOutput.Text & again

txtFieldName.SetFocus


End Sub

Private Sub cmdClear_Click()
frmOutput.txtOutput.Text = ""
cmdCreate.Enabled = True


End Sub

Private Sub cmdCreate_Click()
Dim again As String

If RecipientVar = "" Then RecipientVar = "Var"


If txtTableName.Text = "" Then
MsgBox "Must Enter a table name"
Exit Sub
End If

again = RecipientVar & " = " & Chr(34) & "CREATE TABLE [" & txtTableName.Text & "] (" & Chr(34) & " "
again = again & vbNewLine

frmOutput.txtOutput.Text = again
cmdCreate.Enabled = False
cmdEndCreate.Enabled = True




End Sub

Private Sub cmdEndCreate_Click()
Dim again As String

If frmOutput.txtOutput.Text = "" Then
MsgBox "There is no table to end"
Exit Sub
End If

again = Mid(frmOutput.txtOutput.Text, 1, Len(frmOutput.txtOutput.Text) - 4)
again = again & Chr(34) & vbNewLine
again = again & RecipientVar & " = " & RecipientVar & " & " & Chr(34) & ")" & Chr(34)

frmOutput.txtOutput.Text = again

End Sub

Private Sub Form_GotFocus()
If frmOutput.txtOutput.Text = "" Then cmdCreate.Enabled = True

End Sub

Private Sub Form_Load()
Dim o As Integer
DTArray(0) = "Binary"
DTArray(1) = "BIT"
DTArray(2) = "BYTE"
DTArray(3) = "COUNTER"
DTArray(4) = "Currency"
DTArray(5) = "Date/Time"
DTArray(6) = "GUID"
DTArray(7) = "Single"
DTArray(8) = "Double"
DTArray(9) = "Short"
DTArray(10) = "Long"
DTArray(11) = "LONGTEXT"
DTArray(12) = "LONGBINARY"
DTArray(13) = "Text"

For o = 0 To 13
cboxDataType.AddItem DTArray(o)
Next o

End Sub

Private Sub txtFieldName_GotFocus()
SelectAll txtFieldName

End Sub

Private Sub txtFieldSize_GotFocus()
SelectAll txtFieldSize

End Sub

Private Sub txtTableName_GotFocus()
SelectAll txtTableName

End Sub
