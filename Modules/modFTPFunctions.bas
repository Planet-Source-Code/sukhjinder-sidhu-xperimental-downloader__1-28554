Attribute VB_Name = "modFTPFunctions"
Option Explicit

Global strFTPDownURL                    As String
Global strFTPLastCommand                As String
Global strFTPFile                       As String

Global blnSentUser                      As Boolean

Public Sub AddToList(strListMessage)
Dim strSectionSplit() As String
Dim intSectionMessageCounter As Integer
strSectionSplit() = Split(strListMessage, Chr(13) & Chr(10), -1)

For intSectionMessageCounter = LBound(strSectionSplit) To UBound(strSectionSplit)
    
    If strSectionSplit(intSectionMessageCounter) <> "" Then
        frmMain.lstFTP.AddItem Replace(strSectionSplit(intSectionMessageCounter), vbCrLf, "")
    End If

Next intSectionMessageCounter
End Sub
