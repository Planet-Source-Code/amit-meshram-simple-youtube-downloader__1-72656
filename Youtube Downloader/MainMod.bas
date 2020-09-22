Attribute VB_Name = "MainMod"
Option Explicit

Public VideoName As String
Public GFileName As String
    
Function GetVideoInfo(URL As String, InetPre As Inet)
On Error GoTo Err1:
    
    Dim var_data As Variant
    Dim RespText As String
    
    Dim VideoID As String
    
    frmMain.STBar.Panels(1).Text = "Getting Video Information.."
    
    RespText = InetPre.OpenURL(URL)
    
    Do While InetPre.StillExecuting
        DoEvents
    Loop
    
    var_data = InetPre.GetChunk(1024, icString)
    RespText = RespText & var_data
    Do
        DoEvents
        var_data = InetPre.GetChunk(1024, icString)
        If Len(var_data) = 0 Then Exit Do
        RespText = RespText & var_data
    Loop

    VideoName = GetVideoTitle(RespText)
    
    frmMain.lblVidName = VideoName
    
    If Len(VideoName) = 0 Then
        MsgBox "Failed on getting video title"
        Exit Function
    End If
    
    VideoID = GetVideoID(RespText)
    
    If Len(VideoID) = 0 Then
        MsgBox "Failed on getting video id"
        Exit Function
    End If
    
    GetVideoInfo = "http://youtube.com/get_video?" & VideoID
    Exit Function
    
Err1:
    MsgBox Err.Number & Space(2) & Err.Description & Space(2) & vbCrLf & _
           "Error Occured Please Click - Download Button Again."
End Function

Function GetVideoTitle(RespText As String) As String
    Dim pos1, pos2 As Integer
    Dim tmp1, tmp2, tmp3 As String
    
    If InStr(1, RespText, "content") Then
        pos1 = InStr(1, RespText, "content=")
        pos2 = InStr(pos1, RespText, ">")
        
        tmp1 = Mid(RespText, pos1, pos2 - pos1 - 1)
        tmp2 = Replace(tmp1, "content=", "")
        tmp3 = Replace(tmp2, Chr(&H22), "")
        Debug.Print tmp3
    End If
    GetVideoTitle = Trim(tmp3)
End Function

Function GetVideoID(strResp As String) As String
    Dim VideoID
    Dim tid
    
    VideoID = FindText(strResp, "video_id"": ""([^""]+)")
    tid = GetAnyThing(strResp, Chr(34) & ", " & Chr(34) & "t" & Chr(34) & ": " & Chr(34), Chr(34))
    
    GetVideoID = "video_id=" & VideoID & "&t=" & tid
End Function

Sub DownloadVideo(Link As String, FileName As String)
On Error GoTo Err2:
    Dim FileSize As Long
    Dim SrcSize As Double
    Dim FileData() As Byte
    Dim FileRemaining As Long
    Dim FileSizeCurrent As Long
    Dim PBValue As Integer
    
    Dim FileNumber As Long
        
    frmMain.STBar.Panels(1).Text = "Downloading Video..."
    frmMain.Inet2.Execute Trim(Link), "GET"
    
    Do While frmMain.Inet2.StillExecuting
        DoEvents
    Loop
    
    FileName = Replace(FileName, "/", "")
    FileName = Replace(FileName, "\", "")
    FileName = Replace(FileName, "*", "")
    FileName = Replace(FileName, ":", "")
    FileName = Replace(FileName, "?", "")
    FileName = Replace(FileName, "<", "")
    FileName = Replace(FileName, ">", "")
    FileName = Replace(FileName, "|", "")
    
    GFileName = FileName
    
    FileSize = frmMain.Inet2.GetHeader("Content-Length")
    SrcSize = FileSize / 1000
    
    frmMain.lblVidSize.Caption = SrcSize & " kb"
    
    FileRemaining = FileSize
    FileSizeCurrent = 0
    
    FileNumber = FreeFile
    
    Open App.Path & "/" & FileName For Binary Access Write As #FileNumber
        
        Do Until FileRemaining = 0
            If frmMain.Tag = "Cancel" Then
                frmMain.Inet2.Cancel
                frmMain.STBar.Panels(1).Text = "Stoped by user"
                Exit Sub
            End If
            
            If FileRemaining > 1024 Then
                FileData = frmMain.Inet2.GetChunk(1024, icByteArray)
                FileRemaining = FileRemaining - 1024
            Else
                FileData = frmMain.Inet2.GetChunk(FileRemaining, icByteArray)
                FileRemaining = 0
            End If
            
            FileSizeCurrent = FileSize - FileRemaining
            PBValue = CInt((100 / FileSize) * FileSizeCurrent)
            
            frmMain.lblSaved.Caption = FileSizeCurrent & " bits"
            frmMain.lblRemaining.Caption = FileSize - FileSizeCurrent & " bits"
            frmMain.lblPercent.Caption = PBValue & " %"
            frmMain.STBar.Panels(2).Text = PBValue & " %" & "Downloaded"
            
            Put #FileNumber, , FileData
        Loop
    Close #FileNumber
    MsgBox "Video Downloaded."
    Call frmMain.ResetControls
    Exit Sub
Err2:
    MsgBox Err.Number & Space(2) & Err.Description & Space(2) & vbCrLf & _
           "Error Occured while downloading the file " & vbCrLf & _
           "Please click the download button again"
End Sub

'=========================================================================
Function FindText(Text, Pattern) As String
    Dim RegEx As RegExp
    Dim Matches As Variant
    Set RegEx = New RegExp
        RegEx.Pattern = Pattern
    Set Matches = RegEx.Execute(Text)
    If Matches.Count = 0 Then
        FindText = ""
        Exit Function
    End If
    FindText = Matches(0).SubMatches(0)
End Function

Function GetAnyThing(strResp, StrFind1 As String, StrFind2 As String) As String
On Error Resume Next
    Dim pos1 As Long
    Dim pos2 As Long

    pos1 = InStr(1, strResp, StrFind1) + Len(StrFind1)
    pos2 = InStr(pos1, strResp, StrFind2)
    GetAnyThing = Mid(strResp, pos1, pos2 - pos1)
End Function

Function GetStatus(st As Integer, Inet2 As Inet)
    Select Case st
        Case icError
            GetStatus = Left$(Inet2.ResponseInfo, Len(Inet2.ResponseInfo) - 2)
        Case icResolvingHost, icRequesting, icRequestSent
            GetStatus = "Searching... "
        Case icHostResolved
            GetStatus = "Found" & GFileName
        Case icReceivingResponse, icResponseReceived
            GetStatus = "Receiving data "
        Case icResponseCompleted
            GetStatus = "Connected"
        Case icConnecting, icConnected
            GetStatus = "Connecting..."
        Case icDisconnecting
            GetStatus = "Disconnecting..."
        Case icDisconnected
            GetStatus = "Disconnected"
        Case Else
    End Select
End Function


