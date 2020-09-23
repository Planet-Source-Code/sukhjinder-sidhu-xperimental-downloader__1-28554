Attribute VB_Name = "modDownload"
Option Explicit
'--------Begin Global Declarations Of Variables--------

Global strDownURL                       As String
Global strDownURLName                   As String
Global strDownFileName                  As String
Global strDownSendHead                  As String
Global strDownConnectURL                As String
Global strDownConnectPort               As String
Global strDownProxyServer               As String
Global strDownProxyServerPort           As String
Global strDownProtocol                  As String

Global strDownUsername                  As String
Global strDownPassword                  As String
Global strDownRedirectURL               As String

Global intFreeFile                      As Integer

Global lngBytesSent                     As Long
Global lngBytesRemain                   As Long
Global lngBytesTotal                    As Long
Global lngDownBefore                    As Long

Global blnDownResumeFile                As Boolean
Global blnUseProxy                      As Boolean
Global blnDeleteAll                     As Boolean
Global blnWasResumed                    As Boolean
Public Function AddressOnly(strURL As String)

If InStr(strURL, "://") = 5 Then
    strURL = Mid(strURL, 8)
    AddressOnly = Left(strURL, InStr(strURL, "/") - 1)
End If
If InStr(strURL, "://") = 4 Then
    strURL = Mid(strURL, 7)
    AddressOnly = Left(strURL, InStr(strURL, "/") - 1)
End If

End Function

Public Function GetDataHead(Data As Variant, ToRetrieve As String)
    Dim EndBYTES                        As Integer
    Dim DataHead                        As String
    Dim LengthEnd                       As Integer
    Dim Part                            As Integer
    Dim SecondPart                      As Integer
    Dim RetrieveLength                  As Integer
    
    On Error Resume Next
    If Data = "" Then Exit Function
    
    If InStr(Data, ToRetrieve) > 0 Then
        LengthEnd = Len(Data)
        Part = InStr(Data, ToRetrieve)
        RetrieveLength = Len(ToRetrieve)
        DataHead = Right(Data, LengthEnd - Part - RetrieveLength)
        LengthEnd = Len(DataHead)
        If InStr(DataHead, vbCrLf) > 0 Then
            SecondPart = InStr(DataHead, vbCrLf)
            DataHead = Left(DataHead, SecondPart - 1)
        End If
        GetDataHead = DataHead
    End If
End Function
Public Function FileCheck(Path As String) As Boolean
Dim Disregard                       As Long
On Error Resume Next

    FileCheck = True

    Disregard = FileLen(Path)
    
    If Err <> 0 Then
        FileCheck = False
    End If
    
End Function
Public Function HTTPResponseCode(HTTPHeader As String) As Integer
'Returns Response Code e.g. 404, 200, 202 etc..
  Dim strHeader As String
  
  HTTPResponseCode = 0
  
  If Len(HTTPHeader) > 10 Then
    strHeader = Left(HTTPHeader, InStr(HTTPHeader, vbCrLf))
    HTTPResponseCode = Mid(HTTPHeader, InStr(HTTPHeader, " ") + 1, 3)
  End If
End Function
Public Function NeatNumber(lngNeatNumber As Long) As String
On Error Resume Next
    
    NeatNumber = Format(lngNeatNumber, "#,###")
    
End Function

