VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form8500 
   Caption         =   "ICOM R8500"
   ClientHeight    =   3210
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7560
   Icon            =   "FormRadio.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FormRadio.frx":0442
   ScaleHeight     =   3210
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVol 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   1680
      TabIndex        =   31
      Text            =   "0"
      Top             =   2850
      Width           =   495
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6960
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
   End
   Begin VB.TextBox TxtF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2550
      TabIndex        =   0
      Text            =   "18.000000 "
      Top             =   480
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   7110
      Top             =   2760
   End
   Begin VB.Label LBLWFM 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1110
      TabIndex        =   36
      ToolTipText     =   "Wide FM"
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label LBLFM 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1620
      TabIndex        =   35
      ToolTipText     =   "FM"
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label LBLnbDisplay 
      BackColor       =   &H0080C0FF&
      Caption         =   "NB"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2550
      TabIndex        =   33
      Top             =   840
      Width           =   465
   End
   Begin VB.Label LBLNB 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   1020
      TabIndex        =   32
      ToolTipText     =   "Noise Blanker On/Off"
      Top             =   1770
      Width           =   495
   End
   Begin VB.Label LBLmode 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   30
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label LBLincDown 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   3330
      TabIndex        =   29
      ToolTipText     =   "Decrease Freq Increment"
      Top             =   1770
      Width           =   495
   End
   Begin VB.Label LBLincUp 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   3420
      TabIndex        =   28
      ToolTipText     =   "Increase Freq. Increment"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label LBLinc 
      BackColor       =   &H0080C0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4890
      TabIndex        =   27
      Top             =   540
      Width           =   255
   End
   Begin VB.Label LBLUpFast 
      BackStyle       =   0  'Transparent
      Height          =   525
      Left            =   4830
      TabIndex        =   26
      ToolTipText     =   "Freq Up Fast"
      Top             =   2190
      Width           =   585
   End
   Begin VB.Label LBLDownFast 
      BackStyle       =   0  'Transparent
      Height          =   525
      Left            =   4080
      TabIndex        =   25
      ToolTipText     =   "Freq Down Fast"
      Top             =   2190
      Width           =   585
   End
   Begin VB.Label LBLPower 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   300
      TabIndex        =   24
      ToolTipText     =   "Power On/Off"
      Top             =   540
      Width           =   525
   End
   Begin VB.Label LBLssb 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   2640
      TabIndex        =   23
      ToolTipText     =   "SSB/CW"
      Top             =   1470
      Width           =   405
   End
   Begin VB.Label LBLam 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2130
      TabIndex        =   22
      ToolTipText     =   "AM "
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label LBLvUp 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   1290
      TabIndex        =   21
      ToolTipText     =   "Volume Up"
      Top             =   2160
      Width           =   285
   End
   Begin VB.Label LBLvDown 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   960
      TabIndex        =   20
      ToolTipText     =   "Volume Down"
      Top             =   2160
      Width           =   285
   End
   Begin VB.Label LBLV 
      BackStyle       =   0  'Transparent
      Caption         =   "Vol:"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   1380
      TabIndex        =   19
      Top             =   2850
      Width           =   285
   End
   Begin VB.Label LBLDown 
      BackStyle       =   0  'Transparent
      Height          =   765
      Left            =   4110
      TabIndex        =   18
      ToolTipText     =   "Freq Down"
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label LBLUp 
      BackStyle       =   0  'Transparent
      Height          =   705
      Left            =   4770
      TabIndex        =   17
      ToolTipText     =   "Freq Up"
      Top             =   1380
      Width           =   585
   End
   Begin VB.Label LBLF 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   5130
      TabIndex        =   16
      Top             =   2850
      Width           =   1335
   End
   Begin VB.Label LBLDot 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5700
      TabIndex        =   15
      Top             =   1080
      Width           =   315
   End
   Begin VB.Label LBLCE 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6510
      TabIndex        =   14
      ToolTipText     =   "R button Clear - L button Clear Entry"
      Top             =   1080
      Width           =   285
   End
   Begin VB.Label LBLEnter 
      BackStyle       =   0  'Transparent
      Height          =   465
      Left            =   6900
      TabIndex        =   13
      ToolTipText     =   "Enter Freq"
      Top             =   840
      Width           =   315
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   6480
      TabIndex        =   12
      ToolTipText     =   "9"
      Top             =   810
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   6060
      TabIndex        =   11
      ToolTipText     =   "8"
      Top             =   810
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   5670
      TabIndex        =   10
      ToolTipText     =   "7"
      Top             =   810
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   9
      ToolTipText     =   "6"
      Top             =   540
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   6060
      TabIndex        =   8
      ToolTipText     =   "5"
      Top             =   540
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   5670
      TabIndex        =   7
      ToolTipText     =   "4"
      Top             =   540
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   6
      ToolTipText     =   "3"
      Top             =   270
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   6060
      TabIndex        =   5
      ToolTipText     =   "2"
      Top             =   270
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   5700
      TabIndex        =   4
      ToolTipText     =   "1"
      Top             =   270
      Width           =   345
   End
   Begin VB.Label LBL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   6030
      TabIndex        =   3
      ToolTipText     =   "0"
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label LBLS 
      BackColor       =   &H00000000&
      Caption         =   "Status"
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   2220
      TabIndex        =   2
      Top             =   2850
      Width           =   2835
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "MHz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4230
      TabIndex        =   1
      Top             =   540
      Width           =   645
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   34
      Top             =   480
      Width           =   1035
   End
   Begin VB.Menu mnuControl 
      Caption         =   "&Control"
      Begin VB.Menu mnuLocal 
         Caption         =   "&Local Radio Control"
      End
      Begin VB.Menu mnuRemote 
         Caption         =   "&Remote Radio Control"
      End
   End
   Begin VB.Menu mnuComPort 
      Caption         =   "&Com Port"
   End
   Begin VB.Menu mnuModel 
      Caption         =   "&Model"
      Begin VB.Menu mnuR8500 
         Caption         =   "R8500"
      End
      Begin VB.Menu mnuR75 
         Caption         =   "R75"
      End
   End
End
Attribute VB_Name = "Form8500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public NB As Integer 'NoiseBlanker 0=off 1=on
Dim OKmsg As String
Dim NGmsg As String
Public Frequency As Long
Public R8500Busy As Boolean
Public R8500GettingMode As Boolean
Public R8500GettingResponse As Boolean
Public R8500TuneInc As Double
Public R8500ReadingF As Boolean
Public R8500mode As String
Public R8500Tuning As Boolean
Public R8500on As Boolean
Public R8500Vol As Integer
Public LocalControl As Boolean
Dim CPort As Integer 'comport used by 8500
Dim LastSetF As Double  ' last successfully set frequency or frequency read
Sub Init8500()
On Error Resume Next
LBLS = "Initializing..Wait."
LBLS.Refresh

    MSComm1.CommPort = CPort
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.PortOpen = True


Get8500Status
R8500Vol = Val(GetSetting(App.Title, "Settings", "Volume", "100"))
Set8500Vol R8500Vol
Set8500NB NB
LBLS = "Finished Initializing."
LBLS.Refresh
End Sub
Public Sub Get8500Status()
Get8500F
Get8500Mode
End Sub
Sub Get8500F()
'read F 8500 is set to
Read8500F
End Sub
Sub Get8500V()
'read Volume 8500 is set to

End Sub
Public Sub Get8500Mode()
'read mode setting of 8500
Dim C$, D$
Dim t As Date
Dim A As String
Dim buf As String
Dim p As Integer
Dim SavePortState As Boolean
If R8500GettingMode Then Exit Sub
R8500GettingMode = True
SavePortState = MSComm1.PortOpen
 'A = MSComm1.CommPort
If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
MSComm1.InBufferCount = 0
DoEvents
buf = Chr(254) + Chr(254) + RCVid + Chr(224) + Chr(4) + Chr(253)
'MSComm1.Input = ""
MSComm1.Output = buf
DoEvents
t = Now + 1 / 86400
While Now < t 'And MSComm1.InBufferCount = 0
       DoEvents
Wend
buf = Get8500Response
'Buf = Replace(Buf, Chr(253), "") 'get rid of the FD
If buf <> "P" And buf <> "T" And buf <> "B" And Len(buf) > 7 Then
        C$ = ""
        D$ = ""
        'For i = 1 To Len(Buf) - 1
        '            D$ = Hex$(Asc(Mid$(Buf, i, 1)))
        '            If Len(D$) < 2 Then D$ = "0" + D$
         '           C$ = C$ + D$ 'Hex$(Asc(Mid$(Buf, i, 1)))
       ' Next i
        C$ = DecodeResponse(buf)
        C$ = Left$(C$, Len(C$) - 2)
        If Len(C$) > 6 Then
                C$ = Right$(C$, 6)
        End If
        'is it the right command?
        If Left$(C$, 2) = "04" Then
                    Select Case Right$(C$, 4)
                    Case Is = "0001"
                        LBLmode = "LSB 2.2 kHz"
                    Case Is = "0101"
                        LBLmode = "USB 2.2 kHz"
                    Case Is = "0202"
                        LBLmode = "AM 5.5 kHz"
                    Case Is = "0203"  'wrong in docs
                        LBLmode = "AM narrow  2.2 kHz"
                    Case Is = "0201"    'wrong in docs
                        LBLmode = "AM wide  12 kHz"
                    Case Is = "0301"
                        LBLmode = "CW 2.2 kHz"
                    Case Is = "0302"
                        LBLmode = "CW narrow 0.5 kHz"
                    Case Is = "0501"
                        LBLmode = "FM 12 kHz"
                    Case Is = "0502"
                        LBLmode = "FM narrow  5.5 kHz"
                    Case Is = "0601"
                        LBLmode = "WFM 150 kHz"
              End Select
             End If
             R8500mode = LBLmode.Caption
                
End If
If SavePortState = False Then MSComm1.PortOpen = SavePortState
R8500GettingMode = False
End Sub
Sub SendToRemote(m$)
'send the message M$ to the remote

End Sub
Public Sub Set8500NB(OnOff As Integer)
'0 turn it off  1 turn it on.
Dim t As Date
If Not ControlAllowed Then Exit Sub
If Not LocalControl Then
    FormMain.SetRemote8500NB OnOff
    If OnOff = 0 Then
        LBLS = "Remote set NB off"
    Else
        LBLS = "Remote set NB on"
    End If
    Exit Sub
End If
Dim SavePortState As Boolean
SavePortState = MSComm1.PortOpen
If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
If OnOff = 0 Then
    Buffer = Chr$(254) + Chr$(254) + RCVid + Chr$(224) + Chr$(22) + Chr$(32) + Chr$(253)
Else
    Buffer = Chr$(254) + Chr$(254) + RCVid + Chr$(224) + Chr$(22) + Chr$(33) + Chr$(253)
End If
MSComm1.Output = Buffer
DoEvents
t = Now + 2 / 86400
While Now < t
       DoEvents
Wend
SaveSetting App.Title, "Settings", "NB", Str$(OnOff)
C$ = Get8500Response
'LBLS = DecodeResponse(C$)
If C$ = OKmsg Then
      If OnOff = 1 Then
        LBLnbDisplay = "NB"
      Else
        LBLnbDisplay = ""
      End If
      NB = OnOff
    
Else
    'C$ = Get8500Response
    'LBLS = C$
    'LBLS = "vol set failed"
End If

If SavePortState = False Then MSComm1.PortOpen = SavePortState
FormMain.SendNBtoClients

End Sub

Public Sub Set8500Mode(m$)
'set mode setting of 8500
Dim C$, D$
Dim t As Date
Dim A As String
Dim buf As String
Dim p As Integer
Dim SavePortState As Boolean
If Not ControlAllowed Then Exit Sub
If Not LocalControl Then
    FormMain.SetRemote8500Mode m$
    LBLS = "Remote Mode set " + m$
    Exit Sub
End If
If R8500GettingMode Then Exit Sub
R8500GettingMode = True
SavePortState = MSComm1.PortOpen
If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
MSComm1.InBufferCount = 0
DoEvents
'RCVid = Chr(74)
'RCVid = Chr(90)


buf = Chr(254) + Chr(254) + RCVid + Chr(224) + Chr(6)
Select Case m$
    Case Is = "LSB 2.2 kHz"
        C$ = "0001"
    Case Is = "USB 2.2 kHz"
        C$ = "0101"
    Case Is = "AM 5.5 kHz"
        C$ = "0202"
    Case Is = "AM narrow  2.2 kHz"
        C$ = "0203" 'wrong in docs
    Case Is = "AM wide  12 kHz"
        C$ = "0201"  'wrong in docs
    Case Is = "CW 2.2 kHz"
        C$ = "0301"
    Case Is = "CW narrow 0.5 kHz"
        C$ = "0302"
    Case Is = "FM 12 kHz"
        C$ = "0501"
    Case Is = "FM narrow  5.5 kHz"
        C$ = "0502"
    Case Is = "WFM 150 kHz"
        C$ = "0601"
End Select
buf = buf + Chr$(Val(Mid$(C$, 2, 1)))
buf = buf + Chr$(Val(Mid$(C$, 4, 1)))
buf = buf + Chr(253)
'MSComm1.Input = ""
MSComm1.Output = buf
DoEvents
t = Now + 1 / 86400
While Now < t 'And MSComm1.InBufferCount = 0
       DoEvents
Wend
buf = Get8500Response
        
If buf = OKmsg Then
   LBLmode = m$
   LBLS = "mode changed"
   LBLmode.Refresh
Else
   LBLS = "mode change failed"
End If
        
If SavePortState = False Then MSComm1.PortOpen = SavePortState
R8500GettingMode = False
FormMain.SendModeToClients

End Sub



Public Sub Set8500Vol(Vol As Integer)
Dim t As Date
If Not ControlAllowed Then Exit Sub
If Not LocalControl Then
    FormMain.SetRemote8500Vol Vol
    LBLS = "Remote Vol -> " + Str$(Vol)
    Exit Sub
End If
Dim SavePortState As Boolean
SavePortState = MSComm1.PortOpen
If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
A$ = Trim(Str$(Vol))
While Len(A$) < 4
A$ = "0" + A$
Wend
LevelA = Val("&H" + Right$(A$, 2))
LevelB = Val("&H" + Left$(A$, 2))



'Level = Vol
'One = Level - (10 * Int(Level / 10))
'ten = Int((Level - 100 * Int(Level / 100)) / 10)
'hun = Int((Level - 1000 * Int(Level / 1000)) / 100)
'LevelA = 16 * One
'LevelB = 16 * ten
BufferAudioGain = Chr$(254) + Chr$(254) + RCVid + Chr$(224) + Chr$(20) + Chr$(1) + Chr$(LevelB) + Chr$(LevelA) + Chr$(253)
MSComm1.Output = BufferAudioGain
DoEvents
t = Now + 2 / 86400
While Now < t 'And MSComm1.InBufferCount = 0
       DoEvents
Wend

C$ = Get8500Response
LBLS = DecodeResponse(C$)
If C$ = OKmsg Then
      txtVol = Str$(Vol)
      txtVol.Refresh
      R8500Vol = Vol
      SaveSetting App.Title, "Settings", "Volume", Str$(R8500Vol)
      LBLS = "vol set to " + Str$(Vol)
      FormMain.SendVOLtoClients
Else
    'C$ = Get8500Response
    'LBLS = C$
    'LBLS = "vol set failed"
End If

If SavePortState = False Then MSComm1.PortOpen = SavePortState
FormMain.SendVOLtoClients
End Sub
Function Get8500Response() As String
Dim TimedOut As Boolean
Dim rtn As Long

Dim t As Date
Dim inBuf As String
If R8500GettingResponse = True Then
        Get8500Response = "B"
        LBLS = "response routine busy"
        Exit Function
End If

If Not MSComm1.PortOpen Then
        Get8500Response = "P"
        LBLS = "Com Port Closed"
        Exit Function 'port closed error
End If
t = Now + 5 / 86400
Do
    DoEvents
    If Now > t Then TimedOut = True
    If MSComm1.InBufferCount > 0 Then inBuf = inBuf + MSComm1.Input '
    If Len(inBuf) >= 2 Then
            rtn = InStr(inBuf, Chr(&HFD))
    End If
Loop Until rtn > 0 Or TimedOut

If rtn > 0 Then
            While InStr(inBuf, Chr(&HFD)) < Len(inBuf)
                     inBuf = Right$(inBuf, Len(inBuf) - InStr(inBuf, Chr(&HFD)))
           Wend
            If InStr(inBuf, Chr(&HFD)) > 0 Then
                    Get8500Response = inBuf
            Else
                    Get8500Response = "E"
                    LBLS = "error reading 8500"
            End If
            
Else
            LBLS = "8500 Response Time Out"
            Get8500Response = "T" 'timedout error
End If

End Function


Public Function Read8500F() As Long
Dim C$, D$
Dim t As Date
Dim Fu As Double
Dim A As String
Dim buf As String
Dim p As Integer
Dim SavePortState As Boolean
If R8500ReadingF Then
        Read8500F = 0
        Exit Function
End If
'If Not LocalControl Or Not ControlAllowed Then
'   Read8500F = 0
'   Exit Function
'End If
R8500ReadingF = True
SavePortState = MSComm1.PortOpen

If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
buf = MSComm1.Input
DoEvents
'R8500
'FEFE4AE003FD
buf = Chr(254) + Chr(254) + RCVid + Chr(224) + Chr(3) + Chr(253)

'R75
'buf = Chr(254) + Chr(254) + Chr(90) + Chr(224) + Chr(3) + Chr(253)




'MSComm1.Input = ""
MSComm1.Output = buf
t = Now + 2 / 86400
While Now < t
        DoEvents
Wend
buf = Get8500Response
'Buf = Replace(Buf, Chr(253), "") 'get rid of the FD
R$ = DecodeResponse(buf)
'Debug.Print "ReadF " + R$
If buf <> "P" And buf <> "T" And buf <> "B" And Len(buf) > 10 Then
        C$ = ""
        D$ = ""
        For I = 1 To Len(buf) - 1
                    D$ = Hex$(Asc(Mid$(buf, I, 1)))
                    If Len(D$) < 2 Then D$ = "0" + D$
                    C$ = C$ + D$ 'Hex$(Asc(Mid$(Buf, i, 1)))
        Next I
        If Len(C$) > 10 Then
                C$ = Right$(C$, 10)
        Else
                LBLS = "error reading frequency"
                R8500ReadingF = False
                Exit Function
        End If
        Fu = Val(Right$(C$, 1)) * 100000000
        Fu = Fu + Val(Mid$(C$, 9, 1)) * 1000000000
        Fu = Fu + Val(Mid$(C$, 8, 1)) * 1000000
        Fu = Fu + Val(Mid$(C$, 7, 1)) * 10000000
        Fu = Fu + Val(Mid$(C$, 6, 1)) * 10000
        Fu = Fu + Val(Mid$(C$, 5, 1)) * 100000
        Fu = Fu + Val(Mid$(C$, 4, 1)) * 100
        Fu = Fu + Val(Mid$(C$, 3, 1)) * 1000
        Fu = Fu + Val(Mid$(C$, 2, 1)) * 1
        Fu = Fu + Val(Mid$(C$, 1, 1)) * 10
End If
If Fu > 1000 Then

Read8500F = Fu
Frequency = Fu
LBLF = ""
TxtF = Format(Fu / 1000000, "00.000000")
End If
If SavePortState = False Then MSComm1.PortOpen = SavePortState
R8500ReadingF = False
FormMain.Reset8500Busy
FormMain.TxtF.Text = Format(Fu / 1000000, "00.000")
FormMain.TxtF.Refresh
FormMain.CopyVscale (Fu)
FormMain.SendFtoClients
End Function
Function DecodeResponse(R$) As String
Dim C$, D$

For I = 1 To Len(R$)
                    D$ = Hex$(Asc(Mid$(R$, I, 1)))
                    If Len(D$) < 2 Then D$ = "0" + D$
                    C$ = C$ + D$ 'Hex$(Asc(Mid$(Buf, i, 1)))
Next I
DecodeResponse = C$



End Function
Public Sub Tune8500(F As Long)
'If FormMain.OpMode = 2 Then
'            FormMain.SetFreq F
 '           Exit Sub
'End If
Dim Bufferfreq As String
Dim Freq As Long
'If Not ControlAllowed Then Exit Sub
LocalControl = True
If Not LocalControl Then
    FormMain.SetFreq F
    LBLS = "Remote F set to " + Format(F / 1000000, "###.000")
    Exit Sub
End If

If R8500Tuning Then Exit Sub
FormMain.Set8500Busy
R8500Tuning = True
TxtF.Enabled = False
Dim SavePortState As Boolean
SavePortState = MSComm1.PortOpen
If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
Freq = F / 1000
OneKHz = Freq - (10 * Int(Freq / 10))
TenKHz = Int((Freq - 100 * Int(Freq / 100)) / 10)
hunKHz = Int((Freq - 1000 * Int(Freq / 1000)) / 100)
OneMHz = Int((Freq - 10000 * Int(Freq / 10000)) / 1000)
TenMHz = Int((Freq - 100000 * Int(Freq / 100000)) / 10000)
HunMhz = Int((Freq - 1000000 * Int(Freq / 1000000)) / 100000)
FreqB = 16 * OneKHz
FreqC = 16 * hunKHz + TenKHz
FreqD = 16 * TenMHz + OneMHz
'FEFE4AE00500bbccddeeFD
Bufferfreq = Chr$(254) + Chr$(254) + RCVid + Chr$(224) + Chr$(5) + Chr$(0) + Chr$(FreqB) + Chr$(FreqC) + Chr$(FreqD) + Chr$(HunMhz) + Chr$(253)
R$ = DecodeResponse(Bufferfreq)
'Debug.Print "F req " + R$
'MSComm1.Input = ""
Dim rtn As Long
Dim TimedOut As Boolean
Dim Tries As Integer
Dim inBuf As String
Tries = 0
retry:
MSComm1.InBufferCount = 0
DoEvents
MSComm1.Output = Bufferfreq
t = Now
LBLS = ""
DoEvents
'
'Do
'    DoEvents
'    If Now > T + 3 / 86400 Then TimedOut = True
'    If MSComm1.InBufferCount > 0 Then inBuf = inBuf + MSComm1.Input   '
'    If Len(inBuf) >= 2 Then
'            rtn = InStr(inBuf, Chr(&HFD))
'    End If
    'LBLS = "8500 echoed " + inBuf
'Loop Until rtn > 0 Or TimedOut

inBuf = Get8500Response
'R$ = DecodeResponse(inBuf)
'Debug.Print "Tune " + R$


'If InStr(inBuf, OKmsg) > 0 Then
'        LBLS = "OK msg received"
'End If

'If InStr(inBuf, NGmsg) > 0 Then
'        LBLS = "NG msg received"
'End If
        
        


'For i = 1 To Len(inBuf) - 1
 '                   LBLS = LBLS + Hex$(Asc(Mid$(inBuf, i, 1)))
'Next i
    
'If TimedOut Then LBLS = "Timed Out Tuning"
'If TimedOut And rtn = 0 Then
'        Tries = Tries + 1
'        TimedOut = False
'       ' If Tries < 2 Then GoTo retry
'End If

If SavePortState = False Then MSComm1.PortOpen = SavePortState
R8500Tuning = False
TxtF.Enabled = True

'Read8500F
End Sub
Private Sub Form_Load()
Dim t As String
'RCVid = Chr(74)
'RCVid = Chr(90)

RCVid = GetSetting(App.Title, "Settings", "RCVid", Chr(90))
If RCVid = Chr(74) Then
    mnuR75.Checked = False
    mnuR8500.Checked = True
    Me.Caption = "ICOM R8500"
Else
    mnuR75.Checked = True
    mnuR8500.Checked = False
    Me.Caption = "ICOM R75"
End If


'T = GetSetting(App.Title, "Settings", "R8500on", "False")
'If T = "False" Then R8500on = False Else R8500on = True
'If R85000on Then
'        LBLon.Caption = "ON"
'Else
'        LBLon.Caption = "OFF"
'End If

OKmsg = Chr(254) + Chr(254) + Chr(&HE0) + Chr(&H4A) + Chr(&HFB) + Chr(&HFD)
NGmsg = Chr(254) + Chr(254) + Chr(&HE0) + Chr(&H4A) + Chr(&HFA) + Chr(&HFD)
If GetSetting(App.Title, "Settings", "LocalControl", "False") = "True" Then
    LocalControl = True
    mnuLocal.Checked = True
Else
    LocalControl = False
    mnuRemote.Checked = True
End If

CPort = Val(GetSetting(App.Title, "Settings", "RcvrComPort", "1"))
If CPort < 1 Then CPort = 1
On Error GoTo cporterr
MSComm1.CommPort = CPort
On Error Resume Next
R8500TuneInc = Val(GetSetting(App.Title, "Settings", "TuneInc", "1"))
LBLinc = Trim(Str(R8500TuneInc))

LBLS = ""
NB = Val(GetSetting(App.Title, "Settings", "NB", "0"))
Init8500
Exit Sub
cporterr:

mnuComPort_Click
Resume Next


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Caption = Str$(x) + "    " + Str$(y)
End Sub

Private Sub LBL_Click(Index As Integer)
LBLF = LBLF + Trim(Str(Index))
End Sub

Private Sub LBLam_Click()
Get8500Mode
R8500Busy = True
LBLS = "changing mode"
LBLS.Refresh
If LBLmode = "AM 5.5 kHz" Then
        Set8500Mode "AM narrow  2.2 kHz"
ElseIf LBLmode = "AM narrow  2.2 kHz" Then
        Set8500Mode "AM wide  12 kHz"
Else   '"AM wide  12 kHz"
        Set8500Mode "AM 5.5 kHz"
End If
R8500Busy = False

End Sub

Private Sub LBLCE_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And Len(LBLF) > 0 Then
        LBLF = Left$(LBLF, Len(LBLF) - 1)
End If
If Button = 2 Then
        LBLF = ""
End If

End Sub

Private Sub LBLDot_Click()
LBLF = LBLF + "."
End Sub

Private Sub LBLDown_Click()
If R8500Tuning Then Exit Sub
'CleantxtF
'If Val(TxtF) >= 17# + R8500TuneInc / 1000 Then
            TF = Val(TxtF)
            TF = TF - R8500TuneInc / 1000
            TxtF = Format(TF, "00.000000")
            Tune8500 CLng(Val(TxtF) * 1000000)
            Read8500F
'End If



End Sub

Private Sub LBLDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Caption = Str$(x) + "    " + Str$(y)

End Sub

Private Sub LBLDownFast_Click()
If R8500Tuning Then Exit Sub
'CleantxtF
'If Val(TxtF) >= 17# + R8500TuneInc / 100 Then
            TF = Val(TxtF)
            TF = TF - R8500TuneInc / 100
            TxtF = Format(TF, "00.000000")
            Tune8500 CLng(Val(TxtF) * 1000000)
            Read8500F
'End If



End Sub

Private Sub LBLEnter_Click()
For I = 1 To Len(LBLF)
        A$ = Mid(LBLF, I, 1)
        If InStr("0123456789.", A$) = 0 Then LBLF = Replace(LBLF, A$, "")
Next I
If LBLF = "" Then Exit Sub


'If Val(LBLF) < 17# Then LBLF = "17.000000"
'If Val(LBLF) > 30# Then LBLF = "30.000000"

TxtF = Format(Val(LBLF), "00.000000")
Tune8500 Val(TxtF) * 1000000
LBLF = ""

End Sub




Private Sub LBLFM_Click()
Get8500Mode
R8500Busy = True
LBLS = "changing mode"
LBLS.Refresh
If LBLmode = "FM 12 kHz" Then
        Set8500Mode "FM narrow  5.5 kHz"
Else
        Set8500Mode "FM 12 kHz"
End If
R8500Busy = False

End Sub

Private Sub LBLincDown_Click()
If R8500TuneInc = 100 Then
        R8500TuneInc = 10
ElseIf R8500TuneInc = 10 Then
        R8500TuneInc = 1
End If

LBLinc = Trim(Str(R8500TuneInc))
SaveSetting App.Title, "Settings", "TuneInc", LBLinc.Caption
LBLinc.Refresh
        
        

End Sub

Private Sub LBLincUp_Click()
If R8500TuneInc = 10 Then
        R8500TuneInc = 100
ElseIf R8500TuneInc = 1 Then
        R8500TuneInc = 10
End If

LBLinc = Trim(Str(R8500TuneInc))
SaveSetting App.Title, "Settings", "TuneInc", LBLinc.Caption
LBLinc.Refresh

End Sub

Private Sub LBLmode_Click()
Get8500Mode
End Sub

Private Sub LBLNB_Click()
If LBLnbDisplay = "" Then
    Set8500NB 1
Else
    Set8500NB 0
End If

End Sub

Private Sub LBLPower_Click()
'R8500on = Not R8500on
'If R8500on Then
'        LBLon.Caption = "ON"
'Else
'        LBLon.Caption = "OFF"
'End If
''If R8500on Then
        SaveSetting App.Title, "Settings", "R8500on", "True"
'Else
'        SaveSetting App.Title, "Settings", "R8500on", "False"
'End If

End Sub

Private Sub LBLssb_Click()
Get8500Mode
R8500Busy = True
LBLS = "changing mode"
LBLS.Refresh
If LBLmode = "USB 2.2 kHz" Then
    Set8500Mode "CW 2.2 kHz"
ElseIf LBLmode = "CW 2.2 kHz" Then
    Set8500Mode "CW narrow 0.5 kHz"
ElseIf LBLmode = "CW narrow 0.5 kHz" Then
    Set8500Mode "LSB 2.2 kHz"
Else
    Set8500Mode "USB 2.2 kHz"
End If
R8500Busy = False


End Sub

Private Sub LBLUp_Click()
If R8500Tuning Then Exit Sub
'CleantxtF

'If Val(TxtF) <= 30 - R8500TuneInc / 1000 Then
            TF = Val(TxtF)
            TF = TF + R8500TuneInc / 1000
            TxtF = Format(TF, "00.000000")
            Tune8500 CLng(Val(TxtF) * 1000000)
            Read8500F
'End If

End Sub

Private Sub LBLUpFast_Click()
If R8500Tuning Then Exit Sub
'CleantxtF

'If Val(TxtF) <= 30 - R8500TuneInc / 100 Then
            TF = Val(TxtF)
            TF = TF + R8500TuneInc / 100
            TxtF = Format(TF, "00.000000")
            Tune8500 CLng(Val(TxtF) * 1000000)
            Read8500F
'End If

End Sub

Private Sub LBLvDown_Click()
'If Val(LBLV) < 0 Then
'        LBLV = Str$(Val(LBLV) - 1)
R8500Busy = True
If R8500Vol >= 10 Then
    'Set8500Vol 120
    Set8500Vol R8500Vol - 10
End If
R8500Busy = False

'End If
End Sub

Private Sub LBLvUp_Click()
R8500Busy = True
If R8500Vol <= 245 Then
    Set8500Vol R8500Vol + 10
End If
R8500Busy = False

End Sub
Sub CleantxtF()
        For I = 1 To Len(TxtF)
                A$ = Mid(TxtF, I, 1)
                If InStr("0123456789.", A$) = 0 Then TxtF = Replace(TxtF, A$, "")
        Next I
        If Val(TxtF) < 18# Then TxtF = "17.000000"
        If Val(TxtF) > 30# Then TxtF = "30.000000"
        TxtF = Format(Val(TxtF), "00.000000")

End Sub

Private Sub LBLWFM_Click()
R8500Busy = True
LBLS = "changing mode"
LBLS.Refresh
Set8500Mode "WFM 150 kHz"
R8500Busy = False
End Sub

Private Sub mnuComPort_Click()
On Error Resume Next
R = InputBox("Enter the Com Port for the ICOM ", "Configure Com Port", CPort)
If Val(R) = 0 Then R = "1"
CPort = Val(R)
If MSComm1.PortOpen Then
     MSComm1.PortOpen = False
End If
MSComm1.CommPort = CPort
Init8500
SaveSetting App.Title, "Settings", "RcvrComPort", Str$(CPort)

End Sub

Private Sub mnuLocal_Click()
LocalControl = Not LocalControl
mnuLocal.Checked = LocalControl
mnuRemote.Checked = Not LocalControl
If LocalControl Then
    SaveSetting App.Title, "Settings", "LocalControl", "True"
Else
    SaveSetting App.Title, "Settings", "LocalControl", "False"
End If
If LocalControl Then Init8500
End Sub


Private Sub mnuR75_Click()
RCVid = Chr(90)
mnuR75.Checked = True
mnuR8500.Checked = False
Me.Caption = "ICOM R75"
SaveSetting App.Title, "Settings", "RCVid", RCVid

End Sub

Private Sub mnuR8500_Click()
RCVid = Chr(74)
mnuR75.Checked = False
mnuR8500.Checked = True
Me.Caption = "ICOM R8500"
SaveSetting App.Title, "Settings", "RCVid", RCVid

End Sub

Private Sub mnuRemote_Click()
LocalControl = Not LocalControl
mnuLocal.Checked = LocalControl
mnuRemote.Checked = Not LocalControl
If LocalControl Then
    SaveSetting App.Title, "Settings", "LocalControl", "True"
Else
    SaveSetting App.Title, "Settings", "LocalControl", "False"
End If
End Sub

Private Sub Timer1_Timer()

Get8500Status

End Sub

Private Sub TxtF_KeyPress(KeyAscii As Integer)
If R8500Tuning Then Exit Sub
If KeyAscii = 13 Then
        Tune8500 CLng(Val(TxtF) * 1000000)
        Read8500F
End If

End Sub

Private Sub txtVol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(txtVol) >= 0 And Val(txtVol) <= 255 Then
    Set8500Vol (Val(txtVol))
Else
    txtVol = Str$(R8500Vol)
End If




End If
End Sub
