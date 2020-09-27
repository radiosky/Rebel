Attribute VB_Name = "StringStuff"
Function ExtractStr(SrcStr As String, Front As String, Optional Rear) As String
'look in srcstr for something between front and rear strings and return it
'if only the front string exists it returns the remainder of the string
Dim p, p1, p2 As Long
If InStr(SrcStr, Front) = 0 Then
    ExtractStr = ""
    Exit Function
End If
If IsMissing(Rear) Then
        If Len(SrcStr) <= InStr(SrcStr, Front) + (Len(Front) - 1) Then ExtractStr = "": Exit Function
        ExtractStr = Right$(SrcStr, Len(SrcStr) - (InStr(SrcStr, Front) + (Len(Front) - 1)))
        Exit Function
Else
        If InStr(SrcStr, Rear) = 0 Then ExtractStr = "": Exit Function
End If

'rear must exist after front
If InStr(InStr(SrcStr, Front), SrcStr, Rear) = 0 Then
    ExtractStr = ""
    Exit Function
End If
p = InStr(SrcStr, Front) + Len(Front)
p1 = InStr(p, SrcStr, Rear)
'must have some length
If p1 = p Then
    ExtractStr = ""
    Exit Function
End If
ExtractStr = Mid$(SrcStr, p, p1 - p)

End Function
Function RemoveUnPrintable(m$) As String
'remove all characters with less than a chr val of 32
Dim A$
For i = 1 To Len(m$)
    If Asc(Mid$(m$, i, 1)) >= 32 Then A$ = A$ + Mid$(m$, i, 1)
Next i
RemoveUnPrintable = A$
End Function


Public Function PCase(B)
'converts a string to the first letter cap and the rest small.
Dim A
A = B
If Len(A) < 1 Then PCase = "": Exit Function

Mid$(A, 1, 1) = UCase(Mid$(A, 1, 1))
If Len(A) = 1 Then Exit Function
Mid$(A, 2, Len(A) - 1) = LCase(Mid$(A, 2, Len(A) - 1))
PCase = A

End Function


Function StripNonDigits(S$) As String
Dim Found As Boolean

Do
Found = False
For i = 1 To Len(S$)
    If Asc(Mid$(S$, i, 1)) < Asc("0") Or Asc(Mid$(S$, i, 1)) > Asc("9") Then
        S$ = Replace(S$, Mid$(S$, i, 1), "")
        Found = True
        Exit For
    End If
Next i
Loop Until Found = False
StripNonDigits = S$
End Function

Function CharacterCount(StrIn, Char)
'returns number of occurences of Char in StrIn
Dim ct
For i = 1 To Len(StrIn)
    If Mid$(StrIn, i, 1) = Char Then ct = ct + 1
Next i
CharacterCount = ct


End Function
Function ReadFormattedDate(dFmt, dt)
'returns a VB day number based upon dFmt and date string dt
Dim Y, m, D

If dFmt = "YYYY/MM/DD" Then
    If Len(dt) = 10 And Mid$(dt, 5, 1) = "/" And Mid$(dt, 8, 1) = "/" Then
        Y = Val(Left$(dt, 4))
        m = Val(Mid$(dt, 6, 2))
        D = Val(Right$(dt, 2))
    End If
ElseIf dFmt = "MM/DD/YYYY" Then
    If Len(dt) = 10 And Mid$(dt, 3, 1) = "/" And Mid$(dt, 6, 1) = "/" Then
        m = Val(Left$(dt, 2))
        D = Val(Mid$(dt, 4, 2))
        Y = Val(Right$(dt, 4))
    End If
ElseIf dFmt = "DD/MM/YYYY" Then
    If Len(dt) = 10 And Mid$(dt, 3, 1) = "/" And Mid$(dt, 6, 1) = "/" Then
        D = Val(Left$(dt, 2))
        m = Val(Mid$(dt, 4, 2))
        Y = Val(Right$(dt, 4))
    End If
End If
ReadFormattedDate = DateSerial(Y, m, D)
End Function
Function StripFront(S As String, Delim As String) As String
'returns the first part of a string up to the delim character
'string is returned trimmed of this and the first delim character
'if delim does not exist in S then just return S
'if delim is the first character return an empty string and strip delim from S
Dim p As Long
p = InStr(S, Delim)
If p > 1 Then
    StripFront = Left$(S, p - 1)
    S = Right$(S, Len(S) - p)
ElseIf p = 1 Then
    StripFront = ""
    S = Right$(S, Len(S) - p)
Else
    StripFront = S 'no delim..this is the last of the string
End If

End Function


Function StrFromCodes(ByVal m As String) As String
'given a string [x][y]... where x and y are character codes
'return a string using these character codes
'if M is not numeric then use the string itself
Dim p1 As Long
Dim p2 As Long
Dim D As Double

p1 = InStr(m, "[")
p2 = InStr(m, "]")
If p1 > p2 Or p1 = 0 Or p2 = 0 Then
    StrFromCodes = m
    Exit Function
End If
Do
    p1 = InStr(m, "[")
    p2 = InStr(m, "]")
    If p1 > p2 Or p1 = 0 Or p2 = 0 Then Exit Do
    D = Val(Mid$(m, p1 + 1, Len(m) - (p1 + 1)))
    If D <= 255 And D >= 0 Then
        StrFromCodes = StrFromCodes + Chr(D)
    Else
        Exit Do 'quit if an error
    End If
    m = Right$(m, Len(m) - p2)
Loop Until Len(m) < 3

End Function
