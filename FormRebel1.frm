VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "TenTec Rebel Control Box"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptBand 
      Caption         =   "20m"
      Height          =   195
      Index           =   1
      Left            =   6480
      TabIndex        =   27
      Top             =   600
      Width           =   975
   End
   Begin VB.OptionButton OptBand 
      Caption         =   "40m"
      Height          =   195
      Index           =   0
      Left            =   6480
      TabIndex        =   26
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Macro"
      Height          =   375
      Index           =   7
      Left            =   7680
      TabIndex        =   25
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Macro"
      Height          =   375
      Index           =   6
      Left            =   6840
      TabIndex        =   24
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Macro"
      Height          =   375
      Index           =   5
      Left            =   6000
      TabIndex        =   23
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Macro"
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   22
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Macro"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   21
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Macro"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   20
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Macro"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   19
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Macro"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   18
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtSend 
      Height          =   975
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   6495
   End
   Begin VB.CommandButton cmdSweep 
      Caption         =   "Sweep"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Pic1 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1555.948
      ScaleMode       =   0  'User
      ScaleTop        =   1024
      ScaleWidth      =   3075
      TabIndex        =   15
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Dwn"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Dwn"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   13
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Dwn"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   12
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   10
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdCommPort 
      Caption         =   "Comm Port"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin MSCommLib.MSComm Comm 
      Left            =   8760
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   57600
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   615
      Left            =   8640
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.Frame FrameBW 
      Caption         =   "Bandwidth "
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
      Begin VB.OptionButton optBW 
         Caption         =   "Narrow"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton optBW 
         Caption         =   "Medium"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optBW 
         Caption         =   "Wide"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtF 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Text            =   "14000000"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label LBLS 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   6960
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label LBLsig 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Signal Strength "
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DoEscape As Boolean
Dim Sweeping As Boolean
Dim BW As String
Dim Freq
Dim Sig
Dim Cport


Sub SendMacro(idx As Integer)
If Comm.PortOpen = True Then
    Comm.Output = "M" + ChatMacro(1, idx) + vbCrLf
    txtSend = ChatMacro(1, idx)
End If
End Sub
Sub ReadS()

Comm.Output = "?S" + vbCr

End Sub
Sub ReadF()
Comm.Output = "?F" + vbCr

End Sub
Sub ReadBW()
Comm.Output = "?B" + vbCr

End Sub

Private Sub cmd_Click(Index As Integer)
SendMacro (Index)

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    FormChatMacros.Show
    For I = 0 To 7
        cmd(I).Caption = ChatMacro(0, I)
    Next I
End If
End Sub

Private Sub cmdCommPort_Click()
On Error Resume Next
Cport = InputBox("Enter Comport", "Rebel", Cport)
Comm.PortOpen = False
Comm.CommPort = Cport
Comm.PortOpen = True
SaveSetting App.Title, "Settings", "Cport", Str$(Cport)

End Sub

Private Sub cmdDown_Click(Index As Integer)
Freq = Freq - 10 ^ (2 + Index)
Comm.Output = "F" + Trim(Str(Freq)) + vbCrLf
cmdUpdate_Click
End Sub

Private Sub cmdSweep_Click()
If Comm.PortOpen = False Then Exit Sub
Pic1.Cls
DoEscape = False
For I = 0 To 100
    Freq = 7000000 + 1000 * I
    Sig = -1
    Comm.Output = "F" + Trim(Str(Freq)) + vbCrLf
    t = Now + 0.2 / 86400
    While Now < t
        DoEvents
    Wend
    ReadS
    t = Now + 2 / 86400
    While Sig = -1 And Now < t
        DoEvents
    
    Wend
    If DoEscape Then Exit For
    Pic1.Line (I, 0)-(I, Sig)

Next I
End Sub

Private Sub cmdUp_Click(Index As Integer)
Freq = Freq + 10 ^ (2 + Index)

Comm.Output = "F" + Trim(Str(Freq)) + vbCrLf
cmdUpdate_Click

End Sub

Private Sub cmdUpdate_Click()

If Not Comm.PortOpen Then Comm.PortOpen = True
ReadF
ReadBW
ReadS

End Sub

Private Sub Comm_OnComm()
Dim A$
Dim B
'inBuf = inBuf + Comm.Input
'While InStr(inBuf, vbCr) > 0
'    A$ = Left$(inBuf, InStr(inBuf, vbCr))
'    inBuf = Right$(inBuf, Len(inBuf) - Len(A$))
A$ = Comm.Input
B = InStr(A$, vbCrLf)
'A$ = RemoveUnPrintable(A$)

B = Len(A$)
LBLS = A$
    If InStr(A$, "F") > 0 Then
        txtF.Text = ExtractStr(A$, "F", vbCrLf)
        Freq = Val(txtF)
        
        A$ = Replace(A$, "F" + txtF, "")
    End If
    If InStr(A$, "B") > 0 Then
        idx = Val(ExtractStr(A$, "B", vbCrLf))
        BW = Mid$("NMW", idx + 1, 1)
        optBW(idx).Value = True
    
    
    End If
    If InStr(A$, "S") > 0 And InStr(A$, "e") = 0 Then
        LBLsig = ExtractStr(A$, "S", vbCrLf)
        Sig = Val(LBLsig)
    End If
'Wend






End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then DoEscape = True
End Sub

Private Sub Form_Load()

MacroFileName = GetSetting(App.Title, "Settings", "MacroFileName", App.Path + "\default.scm")
LoadMacros (MacroFileName)
For I = 0 To 7
        cmd(I).Caption = ChatMacro(0, I)
Next I
Cport = Val(GetSetting(App.Title, "Settings", "Cport", "1"))
Comm.CommPort = Cport
Pic1.Scale (0, 512)-(100, 0)
Pic1.Cls
End Sub

Private Sub LBLsig_Click()
    If Comm.PortOpen Then Comm.Output = "?S" + vbCrLf
End Sub

Private Sub OptBand_Click(Index As Integer)
If Index = 0 Then

Else


End If
End Sub

Private Sub optBW_Click(Index As Integer)

Select Case Index
    Case Is = 0
        A$ = "W"
    Case Is = 1
        A$ = "M"
    Case 2
        A$ = "N"
End Select
If A$ <> BW Then
    BW = A$
    Comm.Output = "B" + BW$ + vbCrLf
    cmdUpdate_Click
End If
End Sub

Private Sub txtF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Comm.Output = "F" + txtF + vbCrLf
cmdUpdate_Click
End If


End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Comm.Output = "M" + txtSend + vbCrLf
    txtSend = ""
End If
End Sub
