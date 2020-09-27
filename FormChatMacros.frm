VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormChatMacros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Macros"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "FormChatMacros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save to File"
      Height          =   345
      Left            =   2550
      TabIndex        =   20
      ToolTipText     =   "Save Observer Log Macros to File"
      Top             =   3030
      Width           =   1155
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load from File"
      Height          =   345
      Left            =   1230
      TabIndex        =   19
      ToolTipText     =   "Load Observer Log Macros from File"
      Top             =   3030
      Width           =   1245
   End
   Begin VB.TextBox txtMacro 
      Height          =   255
      Index           =   7
      Left            =   1230
      TabIndex        =   18
      Top             =   2640
      Width           =   4905
   End
   Begin VB.TextBox txtMacro 
      Height          =   255
      Index           =   6
      Left            =   1230
      TabIndex        =   17
      Top             =   2340
      Width           =   4905
   End
   Begin VB.TextBox txtMacro 
      Height          =   255
      Index           =   5
      Left            =   1230
      TabIndex        =   16
      Top             =   2040
      Width           =   4905
   End
   Begin VB.TextBox txtMacro 
      Height          =   255
      Index           =   4
      Left            =   1230
      TabIndex        =   15
      Top             =   1740
      Width           =   4905
   End
   Begin VB.TextBox txtMacro 
      Height          =   255
      Index           =   3
      Left            =   1230
      TabIndex        =   14
      Top             =   1440
      Width           =   4905
   End
   Begin VB.TextBox txtMacro 
      Height          =   255
      Index           =   2
      Left            =   1230
      TabIndex        =   13
      Top             =   1140
      Width           =   4905
   End
   Begin VB.TextBox txtMacro 
      Height          =   255
      Index           =   1
      Left            =   1230
      TabIndex        =   12
      Top             =   840
      Width           =   4905
   End
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Index           =   7
      Left            =   330
      TabIndex        =   11
      Top             =   2640
      Width           =   700
   End
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Index           =   6
      Left            =   330
      TabIndex        =   10
      Top             =   2340
      Width           =   700
   End
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Index           =   5
      Left            =   330
      TabIndex        =   9
      Top             =   2040
      Width           =   700
   End
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Index           =   4
      Left            =   330
      TabIndex        =   8
      Top             =   1740
      Width           =   700
   End
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Index           =   3
      Left            =   330
      TabIndex        =   7
      Top             =   1440
      Width           =   700
   End
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Index           =   2
      Left            =   330
      TabIndex        =   6
      Top             =   1140
      Width           =   700
   End
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Index           =   1
      Left            =   330
      TabIndex        =   5
      Top             =   840
      Width           =   700
   End
   Begin VB.TextBox txtMacro 
      Height          =   255
      Index           =   0
      Left            =   1230
      TabIndex        =   3
      Top             =   540
      Width           =   4905
   End
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Index           =   0
      Left            =   330
      TabIndex        =   2
      Top             =   540
      Width           =   700
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5190
      TabIndex        =   1
      Top             =   3030
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   4230
      TabIndex        =   0
      Top             =   3030
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   90
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Caption                     Macro Text"
      Height          =   195
      Left            =   390
      TabIndex        =   4
      Top             =   210
      Width           =   2535
   End
End
Attribute VB_Name = "FormChatMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLoad_Click()
Dim MP As String 'macro file path
On Error GoTo somerr
MP = GetSetting(App.Title, "Settings", "MacroFileName", App.Path)

If Dir(MP) <> "" Then CmDialog1.InitDir = MP Else CmDialog1.InitDir = App.Path
CmDialog1.DialogTitle = "Load Macros From File  (ext =scm)"
CmDialog1.CancelError = True
CmDialog1.DefaultExt = "scm"
CmDialog1.FileName = "*.scm"
CmDialog1.ShowOpen
If Err = cdlCancel Then Exit Sub

MFN = CmDialog1.FileName
MP = GetPath(CmDialog1.FileName)
sfnum = FreeFile
Open MFN For Input As #sfnum
For I = 0 To 7
        Line Input #sfnum, A$
        TxtCaption(I).Text = A$
        If EOF(sfnum) Then Exit For
        Line Input #sfnum, A$
        txtMacro(I).Text = A$
        If EOF(sfnum) Then Exit For
Next I
SaveSetting App.Title, "Settings", "MacroFileName", MP

getout:
On Error Resume Next
Close sfnum

Exit Sub
somerr:
If Err = cdlCancel Then Exit Sub
MsgBox "There was an error loading " + MFN + vbCrLf + Err.Description
Resume getout
End Sub

Private Sub cmdOK_Click()
SaveMacros MacroFileName
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim MFN As String 'macro file path
On Error GoTo somerr
MP = GetPath(GetSetting(App.Title, "Settings", "MacroFileName", App.Path))

If Dir(MP) = "" Then CmDialog1.InitDir = MP Else CmDialog1.InitDir = App.Path
CmDialog1.DialogTitle = "Save Chat Macros to File (ext=scm)"
CmDialog1.CancelError = True
CmDialog1.DefaultExt = "scm"
CmDialog1.FileName = "*.scm"
CmDialog1.ShowSave
If Err = cdlCancel Then Exit Sub

MFN = CmDialog1.FileName

sfnum = FreeFile
Open MFN For Output As #sfnum
For I = 0 To 7
        Print #sfnum, TxtCaption(I).Text
        Print #sfnum, txtMacro(I).Text
Next I
SaveSetting App.Title, "Settings", "MacroFileName", MFN

getout:
On Error Resume Next
Close sfnum

Exit Sub
somerr:
If Err = cdlCancel Then Exit Sub
MsgBox "There was an error saving " + MFN + vbCrLf + Err.Description
Resume getout

End Sub

Private Sub Form_Load()
For I = 0 To 7
        TxtCaption(I) = ChatMacro(0, I)
        txtMacro(I) = ChatMacro(1, I)
Next I
Refresh
End Sub

