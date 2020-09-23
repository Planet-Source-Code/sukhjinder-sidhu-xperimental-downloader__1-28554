VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XPerimental Downloader"
   ClientHeight    =   4785
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckFTPData 
      Left            =   2640
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckFTP 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstFTP 
      Height          =   1320
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   7575
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   5520
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   1440
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Begin Download"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock sckDownload 
      Left            =   1680
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraCurrentFileDown 
      Caption         =   "Current File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      Begin MSComctlLib.ProgressBar prgDownloadProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbldownper 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblSpeed 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblRemaining 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblSize 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblElapsed 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblReceive 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblPerDone 
         BackStyle       =   0  'Transparent
         Caption         =   "% Downloaded:"
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbldownelapsed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Elapsed Time:"
         Height          =   210
         Left            =   3840
         TabIndex        =   11
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lbldownremain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Remaining:"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label lbldownSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed (Kbps):"
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lblDownTotalSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Size:"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblDownRecSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Received Size:"
         Height          =   210
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.TextBox txtHeader 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1920
      Width           =   7575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileDownload 
         Caption         =   "Begin &Download"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileDashOne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpDashOne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'©2001 Sukhjinder Sidhu
'All Rights Reserved
'
Option Explicit

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Sub Form_Initialize()
'Gives XP Interface if manifest is included.
Dim X As Long
X = InitCommonControls
End Sub

Public Sub cmdDownload_Click()
On Error GoTo ErrHandler
Dim intOutCome As Integer

'Clear variables to default
    blnDownResumeFile = False
    blnWasResumed = False
    lngDownBefore = 0
    prgDownloadProgress.Value = 0
    lngBytesSent = 0
    lngBytesRemain = 0
    lngBytesTotal = 0
    lngDownBefore = 0
    blnSentUser = False
    lstFTP.Clear
    If strDownRedirectURL = "" Then
        strDownConnectPort = ""
        frmConnect.Show 1
    Else
        strDownURL = strDownRedirectURL
    End If
    
'URL For Server File
   
    If strDownURL = "" Then Exit Sub
    
'Set UserName and Password

    Dim intLengthCounter
    
    On Error Resume Next
    
    For intLengthCounter = 1 To Len(strDownURL)
        If Mid$(strDownURL, Len(strDownURL) - intLengthCounter, 1) = "/" Then
            strDownURLName = Mid$(strDownURL, Len(strDownURL) - intLengthCounter + 1)
            Exit For
        End If
    Next intLengthCounter

    If strDownURL = "" Then Exit Sub
    
    If UCase$(Mid$(strDownURL, 1, 4)) = "HTTP" Then
        strDownProtocol = "HTTP"
    Else
        If UCase$(Mid$(strDownURL, 1, 3)) = "FTP" Then
            strDownProtocol = "FTP"
        Else
            MsgBox "Invalid URL Syntax", vbCritical + vbOKOnly, "Error!"
            Exit Sub
        End If
    End If
    
    
'Filename To Save To
    'strDownFileName = "c:\documents and settings\sukhjinder sidhu\desktop\test.txt"
    'strDownFileName = "c:\documents and settings\sukhjinder sidhu\desktop\996.jpg"
    dlgMain.DialogTitle = "Save to..."
    dlgMain.FileName = strDownURLName
    dlgMain.ShowSave
    strDownFileName = dlgMain.FileName
    If strDownFileName = "" Then Exit Sub

On Error GoTo ErrHandler
    
'If no then set URL and Port replace http
    strDownConnectURL = AddressOnly(strDownURL)
    If strDownConnectPort = "" Then
        If strDownProtocol = "HTTP" Then strDownConnectPort = 80
        If strDownProtocol = "FTP" Then strDownConnectPort = 21
    End If

'Disable button(s)
    cmdDownload.Enabled = False
    mnuFileDownload.Enabled = False

'Defaults to 1 byte sent for purposes of file updating
    lngBytesSent = 1

'Check if file exists
    If FileCheck(strDownFileName) = True Then
        intOutCome = MsgBox(strDownFileName + vbCrLf + vbCrLf + "The file already exists. Do you want to resume it?", vbYesNoCancel + vbCritical, "File Exists!")
        If intOutCome = vbNo Then
            Kill strDownFileName
            blnDownResumeFile = False
        End If
        If intOutCome = vbYes Then
            blnDownResumeFile = True
        End If
        If intOutCome = vbCancel Then
            cmdDownload.Enabled = True
            mnuFileDownload.Enabled = True
            Exit Sub
        End If
    End If
        
'Close the socket
    sckDownload.Close
    sckFTP.Close
    sckFTPData.Close
        
'Connect to server and make sure the connnection is in the correct state
    Do
        DoEvents: DoEvents: DoEvents: DoEvents

    Loop Until sckDownload.State = 0 And sckFTP.State = 0 And sckFTPData.State = 0
    
    strFTPFile = Mid$(strDownURL, Len(strDownConnectURL) + 1)
    
'Percentage is 0
    lbldownper.Caption = 0
    lblElapsed.Caption = "00:00:00"
    lblRemaining.Caption = "00:00:00"
    strStart = GetTickCount
    strFinish = 0
    fraCurrentFileDown.Caption = "Current File (" & strDownURLName & ")"
    cmdStop.Enabled = True
    
'Begin Timing
    BeginProgress

'Call the TimeRemaining function and display it in label
    lblRemaining.Caption = TimeRemaining(lbldownper.Caption)

'Enable Timer
    tmrTime.Enabled = True
    If strDownProtocol = "HTTP" Then
        sckDownload.Connect strDownConnectURL, strDownConnectPort
    Else
        sckFTP.Connect strDownConnectURL, strDownConnectPort
    End If
    
Exit Sub
ErrHandler:
    MsgBox "An error has occured." & vbCrLf & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical + vbOKOnly, "Error!"
End Sub

Private Sub cmdPause_Click()
    MsgBox "Pause and resume features have yet to be implemented.", vbOKOnly + vbCritical, "To Do:"
End Sub

Private Sub cmdStop_Click()
'Disconnect from the server
    sckDownload.Close
    sckFTP.Close
    sckFTPData.Close
    tmrTime.Enabled = False
    mnuFileDownload.Enabled = True
    cmdDownload.Enabled = True
    lblRemaining.Caption = ""
    lblElapsed.Caption = ""
    lblSpeed.Caption = ""
    cmdStop.Enabled = False
    'lblSize.Caption = ""
    'lblReceive.Caption = ""
    'lbldownper.Caption = ""
    
    'fraCurrentFileDown.Caption = "Current File"
End Sub

Private Sub Form_Load()
'Sets value of boolean as false
    blnUseProxy = False
    
'Retrieve values from registry
    blnDownResumeFile = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sckDownload.Close
End Sub

Private Sub mnuFileDownload_Click()
'Go to Begin Download click event
    cmdDownload_Click
End Sub

Private Sub mnuFileExit_Click()
'Close the program
    End
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "------------------------------------------------------" & vbCrLf & _
            "Designed  for MagicInstall" & vbCrLf & _
            "" & vbCrLf & _
            "Copyright ©2001 Sukhjinder Sidhu" & vbCrLf & _
            "All Rights Reserved" & vbCrLf & _
            "" & vbCrLf & _
            "Visit: http://www.magicinstall.com/" & vbCrLf & _
            "Contact: info@magicinstall.com" & vbCrLf & vbCrLf & _
            "Version:  0.9.0 - Beta" & vbCrLf & _
            "------------------------------------------------------" & vbCrLf & _
            "", vbOKOnly + vbInformation, "About XPerimental Downloader Source Code!"
End Sub

Private Sub mnuHelpHelp_Click()
    MsgBox "Here are some instructions of how to use this downloader:" & vbCrLf & vbCrLf & _
            "   * To begin a download, press 'Begin Download'" & vbCrLf & _
            "   * To terminate a download, press 'Stop'." & vbCrLf & _
            "   * To resume it again, press 'Begin Download' and re-enter the file details." & vbCrLf & _
            "   * To terminate a download, press 'Stop'." & vbCrLf & vbCrLf & _
            "Current Development State:" & vbCrLf & vbCrLf & _
            "   * HTTP Support:" & vbCrLf & _
            "        - Redirection" & vbCrLf & _
            "        - Pause/Resume (Partial Content)" & vbCrLf & _
            "        - Base64 Username/Password" & vbCrLf & _
            "   * FTP Support:" & vbCrLf & _
            "        - Pause/Resume (Partial Content)" & vbCrLf & _
            "        - Username/Password" & vbCrLf & _
            "   * No Proxy Server Support At Present" & vbCrLf, vbOKOnly + vbInformation, "Help!"
End Sub

Private Sub sckDownload_Close()
'Enable button(s)
    mnuFileDownload.Enabled = True
    cmdDownload.Enabled = True
    tmrTime.Enabled = False
    lngDownBefore = 0
End Sub

Public Sub sckDownload_Connect()
On Error GoTo ErrHandler
'Compile file request header with file information
    strDownSendHead = "GET " + Right(strDownURL, Len(strDownURL) - Len(strDownConnectURL)) & " HTTP/1.0" & vbCrLf & _
    "Accept: *.*, */*" & vbCrLf

    If strDownUsername <> "" Then
        strDownSendHead = strDownSendHead & "Authorization: Basic " & Base64_Encode(strDownUsername & ":" & strDownPassword) & vbCrLf
    End If

'This gives the resume file size
Dim strFileLength As Long

'If the file is in resume state it sends an additional part to the header
    If blnDownResumeFile = True Then
        strFileLength = FileLen(strDownFileName)
        strDownSendHead = strDownSendHead & "Range: bytes=" & strFileLength & "-" & vbCrLf
    End If
    
'Finish off the header with information on client and referer etc..
    strDownSendHead = strDownSendHead & "User-Agent: MagicInstall (Client)" & vbCrLf & _
                      "Referer: " & strDownConnectURL & vbCrLf & _
                      "Host: " & strDownConnectURL & vbCrLf & vbCrLf
    
'Finally send the information
    sckDownload.SendData strDownSendHead

'Now begin timer to update transfer rate information

Exit Sub
ErrHandler:
    MsgBox "An error has occured." & vbCrLf & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical + vbOKOnly, "Error!"
End Sub

Public Sub sckDownload_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrHandler
Dim strReceiveData, strDataChunk, strHeader, strCurrentDataPacket As String
Dim lngDataLength, lngSplitPoint As Long
    
    'If sckDownload.State <> 7 Then Exit Sub
    
'Recieve data and put into strReceiveData
    sckDownload.GetData strReceiveData, vbString ', bytesTotal
    DoEvents ' Solves some bugs
    
'If there is nothing sent due to a bug , yet unsolved, then carry on
    If strReceiveData = "" Then Exit Sub
    
'See if this bit has the header
    If InStr(LCase(strReceiveData), "content-type:") Then
        
        Dim intHeadResponseCode
        intHeadResponseCode = HTTPResponseCode(CStr(strReceiveData))
        
        If intHeadResponseCode = 303 Or intHeadResponseCode = 301 Or intHeadResponseCode = 302 Or intHeadResponseCode = 303 Or intHeadResponseCode = 307 Then
            strDownRedirectURL = GetDataHead(strReceiveData, "Location")
        End If
        
    'See if any 206 errors occured
        If blnDownResumeFile = True Then 'check to see if its gonna resume ok or not..This is actually the worst way to check this.
            If intHeadResponseCode <> 206 Then
                MsgBox "The server you are downloading from does not allow you to resume files.", vbCritical, "Error!"
                sckDownload.Close
                Exit Sub
            End If
        End If
        
    'See if any 404 errors occured
        If intHeadResponseCode = 404 Then
            MsgBox "The file requested was not found on the server.", vbCritical, "File Not Found"
            sckDownload.Close
            Exit Sub
        End If
        
    'See if any 401 errors occured
        If intHeadResponseCode = 401 Then
            MsgBox "Access Denied!  You are not authorized to use this server.  Try using a username and password.", vbCritical, "Access Denied!"
            sckDownload.Close
            Exit Sub
        End If
        
    'Find out where the data splits from the header
        lngSplitPoint = InStr(strReceiveData, vbCrLf & vbCrLf)

    'Retrieve length of the data
        lngDataLength = Len(strReceiveData)
        
    'Get header only and ignore the rest
        strHeader = Left$(strReceiveData, lngSplitPoint - 1)
    
    'Get any data that may have been sent with the header
        strReceiveData = Right$(strReceiveData, lngDataLength - lngSplitPoint - 3)
        
    'Check if the file is in resume mode
        If blnDownResumeFile = True Then

        'Update bytes sent
            lngBytesSent = FileLen(strDownFileName)
            
        'Get actual file size
            lngBytesTotal = GetDataHead(strHeader, "Content-Length:") + lngBytesSent
            
        'Place progressbar maximum
            prgDownloadProgress.Max = lngBytesTotal
            
            If lngBytesSent = lngBytesTotal Then
                MsgBox "Unable to resume because this file appears to be complete already.", vbCritical + vbOKOnly, "Error!"
                sckDownload.Close
                Exit Sub
            End If
                        
            blnDownResumeFile = False
            blnWasResumed = True
            
        Else
        'Update bytes remaining
            If blnWasResumed = False Then
                lngBytesTotal = GetDataHead(strHeader, "Content-Length:")
            End If
        End If
    
    'Update bytes remaining & other byte information
        lngBytesRemain = lngBytesTotal - lngBytesSent
        lblSize.Caption = NeatNumber(lngBytesTotal)
        
    'Change progressbar value
        If lngBytesTotal > 0 Then
            prgDownloadProgress.Max = lngBytesTotal
        End If
        
    'Put header into text box
        txtHeader.Text = strHeader
    
    End If

    'Make sure the string size is correct
    strCurrentDataPacket = String$(Len(strReceiveData), " ")
    strCurrentDataPacket = strReceiveData
        
    'Get free file number
    intFreeFile = FreeFile
        
    'Open file and enter data
    Open strDownFileName For Binary Access Write As #intFreeFile
        Put #intFreeFile, LOF(intFreeFile) + 1, strCurrentDataPacket
        lngBytesSent = Seek(intFreeFile) - 1
    Close #intFreeFile
On Error Resume Next
    'Update captions and figures
    lblReceive.Caption = NeatNumber(lngBytesSent)
    lngBytesRemain = lngBytesTotal - lngBytesSent
    prgDownloadProgress.Value = lngBytesSent
    lbldownper.Caption = Int((prgDownloadProgress.Value / prgDownloadProgress.Max) * 100)
    
    DoEvents
            
Exit Sub
ErrHandler:
    MsgBox "An error has occured." & vbCrLf & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical + vbOKOnly, "Error!"
End Sub

Private Sub sckFTP_Close()
    cmdStop_Click
'Enable button(s)
    mnuFileDownload.Enabled = True
    cmdDownload.Enabled = True
    tmrTime.Enabled = False
    lngDownBefore = 0
End Sub

Private Sub sckFTP_Connect()
    strFTPLastCommand = "CONNECT"
End Sub

Private Sub sckFTP_DataArrival(ByVal bytesTotal As Long)
Dim strFTPRecievedData, strFTPCode As String
    sckFTP.GetData strFTPRecievedData, vbString

    strFTPCode = UCase$(Left(strFTPRecievedData, 3))
    
    If strFTPCode = "220" Then
        AddToList strFTPRecievedData
        'Send UserName
        'Do a delay loop of 200ms
        
        Dim intDelayLoop As Integer
        For intDelayLoop = 1 To 10000
            DoEvents
        Next intDelayLoop
        
        If blnSentUser = False Then
            If strDownUsername = "" Then
                sckFTP.SendData "USER Anonymous" & vbCrLf
                strFTPLastCommand = "USER"
                AddToList "USER Anonymous" & strDownUsername
            Else
                sckFTP.SendData "USER " & strDownUsername & vbCrLf
                strFTPLastCommand = "USER"
                AddToList "USER " & strDownUsername
            End If
            blnSentUser = True
        End If
    End If
    If strFTPCode = "331" Then
        AddToList strFTPRecievedData
        'Send Password
        sckFTP.SendData "PASS " & strDownPassword & vbCrLf
        Dim intFTPPassCounter As Integer
        Dim strNumberOfAs As String
        For intFTPPassCounter = 1 To Len(strDownPassword)
            strNumberOfAs = strNumberOfAs & "*"
        Next intFTPPassCounter
        
        AddToList "PASS " & strNumberOfAs
        strFTPLastCommand = "PASS"
    End If
    If strFTPCode = "230" Then
        AddToList strFTPRecievedData
        'Logged In Okay
        'Go into Passive Mode
        sckFTP.SendData "PASV" & vbCrLf
        AddToList "PASV"
        strFTPLastCommand = "PASV"
    End If
    If strFTPCode = "227" Then
        AddToList strFTPRecievedData
        'Logged In Okay
        Dim strPassiveInfo, strPassiveConnectIP, strPassiveConnectPort As String
        Dim strPassSplit() As String
        strPassiveInfo = Mid$(strFTPRecievedData, InStr(1, strFTPRecievedData, Chr(40)))
        strPassiveInfo = Replace(strPassiveInfo, "(", "")
        strPassiveInfo = Replace(strPassiveInfo, ")", "")
        strPassiveInfo = Replace(strPassiveInfo, vbCrLf, "")
  
        strPassSplit() = Split(strPassiveInfo, ",", -1)
        strPassiveConnectIP = strPassSplit(0) & "." & strPassSplit(1) & "." & strPassSplit(2) & "." & strPassSplit(3)
        strPassiveConnectPort = strPassSplit(4) * 256 + strPassSplit(5)
        sckFTPData.Close
        sckFTPData.Connect strPassiveConnectIP, strPassiveConnectPort
        
        'Set Type
        sckFTP.SendData "TYPE I" & vbCrLf
        AddToList "TYPE I"
        strFTPLastCommand = "TYPE"
        
    End If
    
    If strFTPCode = "150" Then
        AddToList strFTPRecievedData
        'File Size is sent with this command
        Dim strFTPMaxFileSize As Long
        Dim strSize As String
        strFTPMaxFileSize = InStr(1, strFTPRecievedData, "(")
        strSize = Mid$(strFTPRecievedData, strFTPMaxFileSize + 1)
        strSize = Replace(strSize, " ", "")
        strSize = Replace(UCase$(strSize), "BYTES", "")
        strFTPMaxFileSize = InStr(1, strSize, ")")
        strSize = Mid$(strSize, 1, strFTPMaxFileSize - 1)
        lngBytesTotal = CLng(strSize)
        prgDownloadProgress.Max = lngBytesTotal
        
        'Update bytes remaining & other byte information
        lngBytesRemain = lngBytesTotal - lngBytesSent
        lblSize.Caption = NeatNumber(lngBytesTotal)
        
        On Error Resume Next
        
        'Change progressbar value
        prgDownloadProgress.Max = lngBytesTotal
        
    End If
    
    If Left(strFTPRecievedData, 3) = "125" Then
        AddToList strFTPRecievedData
    End If
    
    If strFTPCode = "221" Then
        AddToList strFTPRecievedData
        cmdStop_Click
    End If
    
    If strFTPCode = "200" Then
        AddToList strFTPRecievedData
        
        If blnDownResumeFile = True Then
            'Resume
            sckFTP.SendData "REST " & FileLen(strDownFileName) & vbCrLf
            AddToList "REST " & FileLen(strDownFileName)
            strFTPLastCommand = "REST"
        Else
            'Get Files
            sckFTP.SendData "RETR " & strFTPFile & vbCrLf
            AddToList "RETR " & strFTPFile
            strFTPLastCommand = "RETR"
        End If
    End If
    
    If strFTPCode = "350" Then
        'Resume Supported
        AddToList strFTPRecievedData
                
        'Get Files
        sckFTP.SendData "RETR " & strFTPFile & vbCrLf
        AddToList "RETR " & strFTPFile
        strFTPLastCommand = "RETR"
    End If
    
    If strFTPCode = "226" Then
        AddToList strFTPRecievedData
    End If
   
    If strFTPCode = "504" Then
        AddToList strFTPRecievedData
        MsgBox "Unfortunately, this server does not allow you to resume files!  Please try again and redownload the complete file.", vbOKOnly + vbCritical, "Error!"
        cmdStop_Click
    End If
   
    If strFTPCode = "530" Then
        AddToList strFTPRecievedData
        MsgBox "Sorry, no anonymous access allowed.", vbOKOnly + vbCritical, "Error!"
        cmdStop_Click
    End If
   
    If strFTPCode = "421" Or _
        strFTPCode = "425" Or _
        strFTPCode = "426" Or _
        strFTPCode = "450" Or _
        strFTPCode = "451" Or _
        strFTPCode = "452" Or _
        strFTPCode = "500" Or _
        strFTPCode = "501" Or _
        strFTPCode = "502" Or _
        strFTPCode = "532" Or _
        strFTPCode = "550" Or _
        strFTPCode = "551" Or _
        strFTPCode = "552" Or _
        strFTPCode = "553" Or _
        strFTPCode = "202" _
    Then
        AddToList strFTPRecievedData
        MsgBox "Error trying to download server! Operation Aborted." & vbCrLf & strFTPRecievedData, vbOKOnly + vbCritical, "Error"
        cmdStop_Click
    End If
    
'AddToList "[+]" & strFTPRecievedData

End Sub

Private Sub sckFTPData_Close()
    cmdStop_Click
End Sub

Public Sub sckFTPData_DataArrival(ByVal bytesTotal As Long)
'Error occurs when the server closes connection after dataarrival has been
'called but BEFORE the data has been retrieved
On Error GoTo ErrHandler

Dim strReceiveData As String
Dim intFreeFile As Integer
       
    'If Socket is NOT connected
    If sckFTPData.State <> 7 Then
        MsgBox "Unable to download file!  There is a problem with the connection to the server.", vbCritical + vbOKOnly, "Error!"
        Debug.Print "Not Connected: " & sckFTPData.State
        cmdStop_Click
        Exit Sub
    Else
        'Get Data From Socket
        sckFTPData.GetData strReceiveData, vbString
        'DoEvents
    End If
    
On Error Resume Next
    
    'Close #1 if already open
    Close #1
   
    'Open file and enter data
    Open strDownFileName For Binary Access Write As #1
        Put #1, LOF(1) + 1, strReceiveData
        lngBytesSent = Seek(1) - 1
        DoEvents
    Close #1
        
    'Update captions and figures
    lblReceive.Caption = NeatNumber(lngBytesSent)
    lngBytesRemain = lngBytesTotal - lngBytesSent
    prgDownloadProgress.Value = lngBytesSent
    lbldownper.Caption = Int((prgDownloadProgress.Value / prgDownloadProgress.Max) * 100)
    
    DoEvents: DoEvents
    
    If lngBytesTotal = lngBytesSent Then cmdStop_Click
    
Exit Sub
ErrHandler:
    If Err.Number = "40006" Then
        MsgBox "Unable to download file!  The server closed the connection.", vbCritical + vbOKOnly, "Error!"
    Else
        MsgBox "An error has occured." & vbCrLf & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical + vbOKOnly, "Error!"
    End If
End Sub


Public Sub tmrTime_Timer()
On Error Resume Next

Dim strSpeedOutput As String

'disable timer when progressbar hits 100%
If prgDownloadProgress.Value = prgDownloadProgress.Max Then
    tmrTime.Enabled = False
    strFinish = GetTickCount
End If

'Call the TimeRemaining function and display it in label
lblRemaining.Caption = TimeRemaining(lbldownper.Caption)

'Change elapsed time
    If strFinish = 0 Then
        strElapsed = GetTickCount - strStart
    Else
        strElapsed = strFinish - strStart
    End If
    
    lblElapsed.Caption = FormatCount(CLng(strElapsed))

'Calculate speed
    If lngDownBefore = 0 Then
        lngDownBefore = lngBytesSent
    Else
        If (lngBytesSent - lngDownBefore) = 0 Then
            strSpeedOutput = "0.00"
        Else
            strSpeedOutput = Format((lngBytesSent - lngDownBefore) / 1024, "#.00")
            If strSpeedOutput > 1000 Then strSpeedOutput = Format(strSpeedOutput / 1024, "#")
        End If
        lblSpeed.Caption = strSpeedOutput
        lngDownBefore = lngBytesSent
    End If
End Sub
