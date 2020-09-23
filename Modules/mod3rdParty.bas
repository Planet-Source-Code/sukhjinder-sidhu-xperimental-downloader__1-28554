Attribute VB_Name = "mod3rdParty"
'Contains 3rd Party Functions
'SOURCE:  Unknown
Public Declare Function GetTickCount Lib "kernel32" () As Long

Global strStart                             As String
Global strFinish                            As String
Global strElapsed                           As String

Dim sStart As String
Function BeginProgress()
'Call this when you begin th progressbar
sStart = Now
End Function

Function TimeRemaining(CurrentPercentage As String) As String

Dim maxTime As String
Dim sRemaining As String
Dim Mins As String

On Error GoTo Handler

'determine how much time(in seconds) has passed since the progress bar has started
maxTime = DateDiff("s", sStart, Now)

'make sure percentage is above 0%
If CurrentPercentage > 0 Then
    'calculate how many seconds until progressbar is finished
    sRemaining = Val(maxTime / CurrentPercentage) * Val(100 - CurrentPercentage)
    'convert seconds into Minutes:Seconds format
    Mins = Format(Fix(sRemaining / 60), "00")
    'set return variable to have Minutes:Seconds left and also Hours
    TimeRemaining = Format(Fix(Mins / 60), "00") & ":" & Format(Mins Mod 60, "00") & ":" & Format(sRemaining Mod 60, "00")
End If

Exit Function

Handler:
TimeRemaining = "Error"
MsgBox "Error Number" & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Error"

End Function

Function GetPercentage(ProgressBarCurrentValue As String, ProgressBarMaxValue As String) As String
'calculate Percentage completed of progressbar
GetPercentage = Format(Val(Val(ProgressBarCurrentValue / ProgressBarMaxValue) * 100), "0")
End Function
Function FormatCount(Count As Long) As String
    Dim Days As Integer, Hours As Long, Minutes As Long, Seconds As Long, Miliseconds As Long
    Dim strMinutes, strHours, strSeconds As String
    Miliseconds = Count Mod 1000
    Count = Count \ 1000
    Days = Count \ (24& * 3600&)
    If Days > 0 Then Count = Count - (24& * 3600& * Days)
    Hours = Count \ 3600&
    If Hours > 0 Then Count = Count - (3600& * Hours)
    Minutes = Count \ 60
    Seconds = Count Mod 60

    
    If Seconds < 10 Then
        strSeconds = 0 & Seconds
    Else
        strSeconds = Seconds
    End If
    
    If Hours < 10 Then
        strHours = 0 & Hours
    Else
        strHours = Hours
    End If
    
    If Minutes < 10 Then
        strMinutes = 0 & Minutes
    Else
        strMinutes = Minutes
    End If
    
        FormatCount = strHours & ":" & strMinutes & ":" & strSeconds

End Function
Public Function Base64_Encode(strSource) As String
'
Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
'
Dim strTempLine As String
Dim j As Integer
'
For j = 1 To (Len(strSource) - Len(strSource) Mod 3) Step 3
   '
   strTempLine = strTempLine + Mid(BASE64_TABLE, _
                 (Asc(Mid(strSource, j, 1)) \ 4) + 1, 1)
   '
   strTempLine = strTempLine + Mid(BASE64_TABLE, _
                 ((Asc(Mid(strSource, j, 1)) Mod 4) * 16 _
                 + Asc(Mid(strSource, j + 1, 1)) \ 16) + 1, 1)
   '
   strTempLine = strTempLine + Mid(BASE64_TABLE, _
                 ((Asc(Mid(strSource, j + 1, 1)) Mod 16) * 4 _
                 + Asc(Mid(strSource, j + 2, 1)) \ 64) + 1, 1)
   '
   strTempLine = strTempLine + Mid(BASE64_TABLE, _
                 (Asc(Mid(strSource, j + 2, 1)) Mod 64) + 1, 1)
   '
Next j
'
If Not (Len(strSource) Mod 3) = 0 Then
   '
      If (Len(strSource) Mod 3) = 2 Then
         '
         strTempLine = strTempLine + Mid(BASE64_TABLE, _
                       (Asc(Mid(strSource, j, 1)) \ 4) + 1, 1)
         '
         strTempLine = strTempLine + Mid(BASE64_TABLE, _
                       (Asc(Mid(strSource, j, 1)) Mod 4) * 16 _
                        + Asc(Mid(strSource, j + 1, 1)) \ 16 + 1, 1)
         '
         strTempLine = strTempLine + Mid(BASE64_TABLE, _
                       (Asc(Mid(strSource, j + 1, 1)) Mod 16) * 4 + 1, 1)
         '
         strTempLine = strTempLine & "="
         '
      ElseIf (Len(strSource) Mod 3) = 1 Then
         '
         strTempLine = strTempLine + Mid(BASE64_TABLE, _
                       Asc(Mid(strSource, j, 1)) \ 4 + 1, 1)
         '
         strTempLine = strTempLine + Mid(BASE64_TABLE, _
                       (Asc(Mid(strSource, j, 1)) Mod 4) * 16 + 1, 1)
         '
         strTempLine = strTempLine & "=="
         '
      End If
      '
   End If
   '
Base64_Encode = strTempLine
'
End Function


