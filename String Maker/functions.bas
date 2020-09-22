Attribute VB_Name = "Functions"
'Main Variables
Public QuickFormatOpen As Boolean


'Setup Panel
Public RecipientVar As String
Public DataObject As String
Public BrakeSymbol As Boolean
Public BrakeLine As Boolean

'String Input
Public InputIsDirty As Boolean

'String Output
Public OutString As String

'Other VArs
Public DTArray(13) As String




Public Sub InitStringOutput(NewOutput As Boolean)
If NewOutput = True Then OutString = ""
OutString = OutString & RecipientVar & " = "

End Sub

Public Sub FormatStringOutput(FullInputStr As String)
Dim InputStr As String

For x = 1 To Len(FullInputStr)


InputStr = Mid(FullInputStr, x, 1)

If InputStr = vbNewLine And BrakeLine = True Then
    Select Case BrakeSymbol
        Case False
            OutString = OutString & RecipientVar & " = "
        Case True
            OutString = OutString & Chr(34) & " _" & vbNewLine & "& " & Chr(34)
End Select

End If

Next x

End Sub

Public Sub SelectAll(textCont As Object)
With textCont
    .SelStart = 0
    .SelLength = Len(textCont.Text)
End With

End Sub

Public Sub Normalize(thisForm As Object)
If thisForm.WindowState = vbMinimized Then
thisForm.WindowState = vbNormal
End If

End Sub




