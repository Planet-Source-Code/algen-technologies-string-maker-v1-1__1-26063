VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "String Maker V1"
   ClientHeight    =   3915
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Main.frx":0000
      Top             =   600
      Width           =   4935
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3660
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
            Text            =   "String Warp:"
            TextSave        =   "String Warp:"
            Key             =   "SWarp"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
            Text            =   "Code Warp"
            TextSave        =   "Code Warp"
            Key             =   "CWarp"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Button 
      Caption         =   "Clear"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btnCopyString 
      Caption         =   "Copy String"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btnShowString 
      Caption         =   "Show String"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Variable Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mainMenuSM 
      Caption         =   "&String Maker"
      Begin VB.Menu mainMenuSMShow 
         Caption         =   "&Show Coded String..."
      End
      Begin VB.Menu mainMenuSMClip 
         Caption         =   "Send to &ClipBoard"
      End
      Begin VB.Menu mainMenuSMLine 
         Caption         =   "-"
      End
      Begin VB.Menu mainMenuSMExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mainMenuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mainMenuOptionsSW 
         Caption         =   "Set &String Warp..."
      End
      Begin VB.Menu mainMenuOptionsCW 
         Caption         =   "Set &Code Warp..."
      End
      Begin VB.Menu mainMenuOptionsVariable 
         Caption         =   "Use &Variable"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mainMenuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mainMenuHelpH 
         Caption         =   "&Help..."
      End
      Begin VB.Menu mainMenuHelpA 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public sWarpVar As Integer, cWarpVAr As Integer


Private Sub btnShowString_Click(Index As Integer)
Debug.Print Str(Asc(Text1.Text))
End Sub

Private Sub Button_Click(Index As Integer)
Text2.Text = ""

End Sub

Private Sub Form_Load()


sWarpVar = 30
Status.Panels(1).Text = "String Warp Not Yet Implemented"
Status.Panels(1).Enabled = False
Status.Panels(2).Text = "Code Warp Not Yet Implemented"
Status.Panels(2).Enabled = False

End Sub

Private Sub mainMenuSMExit_Click()
End

End Sub

