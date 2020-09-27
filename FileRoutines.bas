Attribute VB_Name = "FileRoutines"
Private Enum CSIDL_VALUES
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D
    CSIDL_FLAG_PER_USER_INIT = &H800
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_MASK = &HFF00
End Enum

Private Const SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
Private Const SHGFP_TYPE_DEFAULT = &H1

Private Const MAX_LENGTH = 260
Private Const S_OK = 0
Private Const S_FALSE = 1

Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long

Private Declare Function SHGetFolderPath Lib "shfolder.dll" _
   Alias "SHGetFolderPathA" _
  (ByVal hwndOwner As Long, _
   ByVal nFolder As Long, _
   ByVal hToken As Long, _
   ByVal dwReserved As Long, _
   ByVal lpszPath As String) As Long
   





Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer _
    As String) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.DLL" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.DLL" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.DLL" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'This function should be called by any app that
'changes anything in the shell. The shell will then
'notify each "notification registered" window of this action.
Declare Sub SHChangeNotify Lib "shell32" _
   (ByVal wEventId As SHCN_EventIDs, _
    ByVal uFlags As SHCN_ItemFlags, _
    ByVal dwItem1 As Long, _
    ByVal dwItem2 As Long)

'Shell notification event IDs
Public Enum SHCN_EventIDs
   SHCNE_RENAMEITEM = &H1          '(D) A non-folder item has been renamed.
   SHCNE_CREATE = &H2              '(D) A non-folder item has been created.
   SHCNE_DELETE = &H4              '(D) A non-folder item has been deleted.
   SHCNE_MKDIR = &H8               '(D) A folder item has been created.
   SHCNE_RMDIR = &H10              '(D) A folder item has been removed.
   SHCNE_MEDIAINSERTED = &H20      '(G) Storage media has been inserted into a drive.
   SHCNE_MEDIAREMOVED = &H40       '(G) Storage media has been removed from a drive.
   SHCNE_DRIVEREMOVED = &H80       '(G) A drive has been removed.
   SHCNE_DRIVEADD = &H100          '(G) A drive has been added.
   SHCNE_NETSHARE = &H200          'A folder on the local computer is being
                                   '    shared via the network.
   SHCNE_NETUNSHARE = &H400        'A folder on the local computer is no longer
                                   '    being shared via the network.
   SHCNE_ATTRIBUTES = &H800        '(D) The attributes of an item or folder have changed.
   SHCNE_UPDATEDIR = &H1000        '(D) The contents of an existing folder have changed,
                                   '    but the folder still exists and has not been renamed.
   SHCNE_UPDATEITEM = &H2000       '(D) An existing non-folder item has changed, but the
                                   '    item still exists and has not been renamed.
   SHCNE_SERVERDISCONNECT = &H4000 'The computer has disconnected from a server.
   SHCNE_UPDATEIMAGE = &H8000&     '(G) An image in the system image list has changed.
   SHCNE_DRIVEADDGUI = &H10000     '(G) A drive has been added and the shell should
                                   '    create a new window for the drive.
   SHCNE_RENAMEFOLDER = &H20000    '(D) The name of a folder has changed.
   SHCNE_FREESPACE = &H40000       '(G) The amount of free space on a drive has changed.

#If (WIN32_IE >= &H400) Then
   SHCNE_EXTENDED_EVENT = &H4000000 '(G) Not currently used.
#End If

  SHCNE_ASSOCCHANGED = &H8000000   '(G) A file type association has changed.
  SHCNE_DISKEVENTS = &H2381F       '(D) Specifies a combination of all of the disk
                                   '    event identifiers.
  SHCNE_GLOBALEVENTS = &HC0581E0   '(G) Specifies a combination of all of the global
                                   '    event identifiers.
  SHCNE_ALLEVENTS = &H7FFFFFFF
  SHCNE_INTERRUPT = &H80000000     'The specified event occurred as a result of a system
                                   'interrupt. It is stripped out before the clients
                                   'of SHCNNotify_ see it.
End Enum

#If (WIN32_IE >= &H400) Then
   Public Const SHCNEE_ORDERCHANGED = &H2 'dwItem2 is the pidl of the changed folder
#End If

'Notification flags
'uFlags & SHCNF_TYPE is an ID which indicates
'what dwItem1 and dwItem2 mean
Public Enum SHCN_ItemFlags
   SHCNF_IDLIST = &H0         'LPITEMIDLIST
   SHCNF_PATHA = &H1          'path name
   SHCNF_PRINTERA = &H2       'printer friendly name
   SHCNF_DWORD = &H3          'DWORD
   SHCNF_PATHW = &H5          'path name
   SHCNF_PRINTERW = &H6       'printer friendly name
   SHCNF_TYPE = &HFF
  'Flushes the system event buffer. The
  'function does not return until the system
  'is finished processing the given event.
   SHCNF_FLUSH = &H1000
  'Flushes the system event buffer. The function
  'returns immediately regardless of whether
  'the system is finished processing the given event.
   SHCNF_FLUSHNOWAIT = &H2000

#If UNICODE Then
  SHCNF_PATH = SHCNF_PATHW
  SHCNF_PRINTER = SHCNF_PRINTERW
#Else
  SHCNF_PATH = SHCNF_PATHA
  SHCNF_PRINTER = SHCNF_PRINTERA
#End If

End Enum


Public Function ShortFileName(ByVal long_name As String) As _
    String
Dim length As Long
Dim short_name As String

    short_name = Space$(1024)
    length = GetShortPathName( _
        long_name, short_name, _
        Len(short_name))
    If length < 1 Then
        MsgBox "Error converting path '" & _
            long_name & "' into a short name", _
            vbExclamation Or vbOKOnly, "Path Error"
    Else
        ShortFileName = Left$(short_name, length)
    End If
End Function


Public Function UpOneLevel(tPath) As String
'return a path up one level from that of tPath
'returns with the \ still on the path
Dim tP As String
tP = tPath
If InStr(tP, "\") = 0 Then UpOneLevel = tPath: Exit Function
Do
    tP = Left$(tP, Len(tP) - 1)
Loop Until Right$(tP, 1) = "\"

UpOneLevel = tP
End Function

Public Function GetMyDocumentsFolderPath(myForm As Form) As String
   Dim csidl As Long
   Dim result As Long
   Dim buff As String
   Dim dwFlags As Long
  
   csidl = &HD

  'fill buffer with the specified folder item
   buff = Space$(MAX_LENGTH)
   
   'dwFlags = dwFlags Or CSIDL_FLAG_PER_USER_INIT
   'dwFlags = dwFlags Or CSIDL_FLAG_NO_ALIAS
   'dwFlags = dwFlags Or CSIDL_FLAG_DONT_VERIFY
   
   result = SHGetFolderPath(myForm.hWnd, _
                      csidl, _
                      -1, _
                      SHGFP_TYPE_CURRENT, _
                      buff)
   If result = S_OK Then GetMyDocumentsFolderPath = UpOneLevel(TrimNull(buff))
       

   
End Function


Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
   
End Function

 

Public Function GetFolderPath(myForm As Form, myTitle As String) As String

    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim spath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hwndOwner = myForm.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(myTitle, "...")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
        '.pIDLRoot = "c:\windows\"
         
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        spath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, spath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(spath, vbNullChar)
        If iNull Then
            spath = Left$(spath, iNull - 1)
        End If
    End If
    GetFolderPath = spath
End Function
Function GetPath(ByVal tPath) As String
'return the path minus the file name supplied in tPath
On Error Resume Next

While Right$(tPath, 1) <> "\" And Len(tPath) > 0
    tPath = Left$(tPath, Len(tPath) - 1)
Wend

GetPath = tPath

End Function

Function StripExtension(F$) As String
If Len(F$) > 4 Then
        If InStr(F$, ".") > 0 Then
                StripExtension = Left$(F$, InStr(Len(F$) - 4, F$, ".") - 1)
                Exit Function
        End If

End If
StripExtension = F$


End Function
Function GetRootFileName(tPath) As String
Dim R As String
R = GetFileNameOnly(tPath)
GetRootFileName = StripExtension(R)

End Function
Function GetExtension(F$) As String
Dim p As Long
p = InStr(F$, ".")
If p = 0 Then
    GetExtension = ""
    Exit Function
End If
GetExtension = Right$(F$, Len(F$) - p)


End Function
Function GetFileNameOnly(ByVal tPath) As String
On Error Resume Next

While InStr(tPath, "\") > 0
     tPath = Right$(tPath, Len(tPath) - 1)
Wend

GetFileNameOnly = tPath

End Function

'Extension is three letters without the "."
'PathToExecute is full path to exe file
'Application Name is any name you want as description of
' Extension
Public Sub AssociateFileExtension(Extension As String, _
    PathToExecute As String, ApplicationName As String)
Dim sKeyName As String   'Holds Key Name in registry.
Dim sKeyValue As String  'Holds Key Value in registry.
Dim ret&           'Holds error status, if any, from API
    ' calls.
Dim lphKey&        'Holds created key handle from
    ' RegCreateKey.

    ret& = InStr(1, Extension, ".")
    If ret& <> 0 Then
        MsgBox "Extension has . in it. Remove and try " & _
            "again."
        Exit Sub
    End If

    'This creates a Root entry called 'ApplicationName'.
    sKeyName = ApplicationName
    sKeyValue = ApplicationName
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, _
        lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)

    'This creates a Root entry for the extension to be
    ' associated with 'ApplicationName'.
    sKeyName = "." & Extension
    sKeyValue = ApplicationName
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, _
        lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)

    'This sets the command line for 'ApplicationName'.
    sKeyName = ApplicationName
    sKeyValue = PathToExecute & " %1"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, _
        lphKey&)
    ret& = RegSetValue&(lphKey&, "shell\open\command", _
        REG_SZ, sKeyValue, MAX_PATH)

    'This sets the default icon
    sKeyName = ApplicationName
    sKeyValue = App.Path & "\" & App.EXEName & ".exe,0"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, _
        lphKey&)
    ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, _
        sKeyValue, MAX_PATH)

    'Force Icon Refresh
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
'Thanks to Ralf Gerstenberger <ralf.gerstenberger@arcor.de>
' for pointing out
'that WinXP seems to require the SHCNF_FLUSHNOWAIT flag in
' SHChangeNotify
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/functions/shchangenotify.asp
End Sub

Public Sub UnAssociateFileExtension(Extension As String, _
    ApplicationName As String)
Dim sKeyName As String   'Finds Key Name in registry.
Dim sKeyValue As String  'Finds Key Value in registry.
Dim ret&           'Holds error status, if any, from API
    ' calls.

    ret& = InStr(1, Extension, ".")
    If ret& <> 0 Then
        MsgBox "Extension has . in it. Remove and try " & _
            "again."
        Exit Sub
    End If

    'This deletes the default icon
    sKeyName = ApplicationName
    ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName & _
        "\DefaultIcon")

    'This deletes the command line for "ApplicationName".
    sKeyName = ApplicationName
    ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName & _
        "\shell\open\command")

    'This deletes a Root entry called "ApplicationName".
    sKeyName = ApplicationName
    ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName & _
        "\shell\open")

    'This deletes a Root entry called "ApplicationName".
    sKeyName = ApplicationName
    ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName & _
        "\shell")

    'This deletes a Root entry called "ApplicationName".
    sKeyName = ApplicationName
    ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName)

    'This deletes the Root entry for the extension to be
    ' associated with "ApplicationName".
    sKeyName = "." & Extension
    ret& = RegDeleteKey(HKEY_CLASSES_ROOT, sKeyName)

    'Force Icon Refresh
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

Public Sub FolderList(ByVal Pathname As String, FL() As String, Optional DirCount As Long)
    'Returns a array containing all files
    'at this directory level and lower.
   
    Dim ShortName As String, LongName As String
    Dim NextDir As String
    Static fFList As Collection
    Screen.MousePointer = vbHourglass
    'First time through only, create collect
    '     ion
    'to hold folders waiting to be processed
    '     .


    If fFList Is Nothing Then
        Set fFList = New Collection
        fFList.Add Pathname
        DirCount = 0
        FileCount = 0
    End If


    Do
        'Obtain next directory from list
    NextDir = fFList.Item(1)
    'Remove next directory from list
    fFList.Remove 1
    'List files in directory
    ShortName = Dir(NextDir & "\*.*", vbNormal Or _
    vbArchive Or _
    vbDirectory)


    Do While ShortName > ""


        If ShortName = "." Or ShortName = ".." Then
            'skip it
        Else
            'process it
            LongName = NextDir & "\" & ShortName


            If (GetAttr(LongName) And vbDirectory) > 0 Then
                'it's a directory - add it to the list o
                '     f directories to process
                ReDim Preserve FL(DirCount)
                FL(DirCount) = LongName
                DirCount = DirCount + 1
           ' Else
                'it's a file - add it to the list of fil
                '     es.
               
            '    FileCount = FileCount + 1
                ' FileList = FileList & LongName & vbCrLf
           '     ReDim Preserve FList(FileCount)
              '  FList(FileCount) = LongName
            '
            End If
        End If
        ShortName = Dir()
    Loop
Loop Until fFList.Count = 0
Screen.MousePointer = vbNormal

Set fFList = Nothing

End Sub
'**************************************

Private Function FixDir(ByVal this As String) As String
FixDir = Trim(this)
If Right(FixDir, 1) <> "\" Then FixDir = FixDir & "\"
End Function

Public Sub DeleteEmptyFolders(ByVal Directory As String)
DoEvents
'this is a recursive routine that deletes empty folders and all
'empty subfolders USE With Care!

Dim test As String
Dim Found As Boolean
Dim ComeBack As String

Restart:
test = Dir(Directory & "*", vbDirectory Or vbArchive Or vbHidden Or vbReadOnly Or vbSystem)

Do Until test = "" Or StopIt = True
If ComeBack <> "" Then
  If test = ComeBack Then ComeBack = ""
Else
If test <> "." And test <> ".." Then
  Found = True
  DoEvents
  If (GetAttr(Directory & test) And vbDirectory) = vbDirectory Then
    DeleteEmptyFolders FixDir(Directory & test)
    ComeBack = test
    GoTo Restart
  End If
End If
End If
test = Dir
Loop

If Found = False And StopIt = False Then
  RmDir Directory
  DoEvents
End If

End Sub


Public Sub FileList(ByVal Pathname As String, FList() As String, Optional DirCount As Long, Optional FileCount As Long, Optional MatchStr As String)
    'Returns a array containing all files
    'at this directory level and lower.
   
    Dim ShortName As String, LongName As String
    Dim NextDir As String
    Static FolderList As Collection
    Screen.MousePointer = vbHourglass
    'First time through only, create collect
    '     ion
    'to hold folders waiting to be processed
    '     .


    If FolderList Is Nothing Then
        Set FolderList = New Collection
        FolderList.Add Pathname
        DirCount = 0
        FileCount = 0
    End If


    Do
        'Obtain next directory from list
    NextDir = FolderList.Item(1)
    'Remove next directory from list
    FolderList.Remove 1
    'List files in directory
    ShortName = Dir(NextDir & "\*.*", vbNormal Or _
    vbArchive Or _
    vbDirectory)


    Do While ShortName > ""


        If ShortName = "." Or ShortName = ".." Then
            'skip it
        Else
            'process it
            LongName = NextDir & "\" & ShortName


            If (GetAttr(LongName) And vbDirectory) > 0 Then
                'it's a directory - add it to the list o
                '     f directories to process
                FolderList.Add LongName
                DirCount = DirCount + 1
            Else
                'it's a file - add it to the list of fil
                '     es.
                If IsMissing(MatchStr) Then
                    FileCount = FileCount + 1
                    ' FileList = FileList & LongName & vbCrLf
                    ReDim Preserve FList(FileCount)
                    FList(FileCount) = LongName
                Else
                    If InStr(UCase$(ShortName), UCase$(MatchStr)) > 0 Then
                        FileCount = FileCount + 1
                        ' FileList = FileList & LongName & vbCrLf
                        ReDim Preserve FList(FileCount)
                        FList(FileCount) = LongName
                    End If
                End If
            End If
        End If
        ShortName = Dir()
    Loop
Loop Until FolderList.Count = 0
Screen.MousePointer = vbNormal

Set FolderList = Nothing

End Sub

Public Function FindAssociatedProgram(ByVal Extension As _
    String) As String
Dim temp_title As String
Dim temp_path As String
Dim fnum As Integer
Dim result As String
Dim pos As Integer

    ' Get a temporary file name with this extension.
    GetTempFile Extension, temp_path, temp_title

    ' Make the file.
    fnum = FreeFile
    Open temp_path & temp_title For Output As fnum
    Close fnum

    ' Get the associated executable.
    result = Space$(1024)
    FindExecutable temp_title, temp_path, result
    pos = InStr(result, Chr$(0))
    FindAssociatedProgram = Left$(result, pos - 1)

    ' Delete the temporary file.
    Kill temp_path & temp_title
End Function

' Return a temporary file name.
Private Sub GetTempFile(ByVal Extension As String, ByRef _
    temp_path As String, ByRef temp_title As String)
Dim I As Integer

    If Left$(Extension, 1) <> "." Then Extension = "." & _
        Extension

    temp_path = Environ("TEMP")
    If Right$(temp_path, 1) <> "\" Then temp_path = _
        temp_path & "\"

    I = 0
    Do
        temp_title = "tmp" & Format$(I) & Extension
        If Len(Dir$(temp_path & temp_title)) = 0 Then Exit _
            Do
        I = I + 1
    Loop
End Sub

Public Function IsAFile(F$) As Boolean

Dim fnum As Long
On Error GoTo ferr
IsAFile = True
fnum = FreeFile
Open F$ For Input As fnum
outahere:

On Error Resume Next
Close fnum
Exit Function
ferr:
If Err.Number = 76 Or Err.Number = 53 Or Err.Number = 52 Then
        IsAFile = False
End If
Resume outahere

End Function

