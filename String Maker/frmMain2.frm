VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H0080C0FF&
   Caption         =   "String Maker V1.1"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8400
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Project1.TrayArea TrayArea 
      Left            =   600
      Top             =   1440
      _ExtentX        =   900
      _ExtentY        =   900
      ToolTip         =   "String Maker"
   End
   Begin VB.Menu mnExit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mnToolbx 
      Caption         =   "&ToolBoxes"
      Begin VB.Menu mnSetup 
         Caption         =   "Setup Panel"
      End
      Begin VB.Menu mnInput 
         Caption         =   "Input String"
      End
      Begin VB.Menu mnOutput 
         Caption         =   "Output String"
      End
      Begin VB.Menu mnusqlcreatetable 
         Caption         =   "SQL CreateTable"
      End
   End
   Begin VB.Menu mnupop 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnupopformat 
         Caption         =   "Format a string"
      End
      Begin VB.Menu mnupopcreatetable 
         Caption         =   "Create SQL Table..."
      End
      Begin VB.Menu menusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnupopsmg 
         Caption         =   "String Maker GUI..."
      End
      Begin VB.Menu mnupopexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
QuickFormatOpen = False
mnupopformat.Enabled = False


frmSetup.Show
frmOutput.Show
frmInput.Show

frmOutput.Top = Me.Height - frmOutput.Height - 1000
frmOutput.Left = Me.Width - frmOutput.Width - 1000


Set TrayArea.Icon = Me.Icon
TrayArea.Visible = True




frmInput.SetFocus

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
Cancel = 1
frmMain.Hide
mnupopformat.Enabled = True

Exit Sub
End If

TrayArea.Visible = False



End Sub

Private Sub mnExit_Click()
TrayArea.Visible = False

End

End Sub

Private Sub mnInput_Click()
frmInput.Show
End Sub

Private Sub mnOutput_Click()
frmOutput.Show
End Sub

Private Sub mnSetup_Click()
frmSetup.Show
End Sub

Private Sub mnuHelp_Click()
Dim way As String
way = App.Path & "\String maker.txt"

Shell "Notepad " & way, vbNormalFocus


End Sub

Private Sub mnupopcreatetable_Click()
If QuickFormatOpen = True Then
Unload frmInput2
QuickFormatOpen = False
End If

mnupopformat.Enabled = False


frmMain.Show
Normalize frmMain

frmCreateTable.Show
frmCreateTable.SetFocus

End Sub

Private Sub mnupopexit_Click()
TrayArea.Visible = False
End

End Sub

Private Sub mnupopformat_Click()



frmInput2.Show
frmInput2.SetFocus
QuickFormatOpen = True

End Sub

Private Sub mnupopsmg_Click()
If QuickFormatOpen = True Then
Unload frmInput2
QuickFormatOpen = False
End If

mnupopformat.Enabled = False



frmMain.Show
Normalize frmMain
frmMain.SetFocus
End Sub

Private Sub mnusqlcreatetable_Click()
frmCreateTable.Show


End Sub

Private Sub TrayArea_DblClick()
If QuickFormatOpen = True Then
Unload frmInput2
QuickFormatOpen = False
End If

mnupopformat.Enabled = False



frmMain.Show
Normalize frmMain
frmMain.SetFocus

End Sub

Private Sub TrayArea_MouseDown(Button As Integer)
If QuickFormatOpen = True Then
mnupopformat.Enabled = False
End If


If Button = vbRightButton Then
Me.PopupMenu mnupop, 8
End If

End Sub

