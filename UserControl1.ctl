VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSWinsockLib.Winsock wSocket 
      Left            =   1320
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      Picture         =   "UserControl1.ctx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_fhIn As Integer

Dim sendnumber As Integer
Dim m_IcyDataReceived As Boolean
Dim m_File As String
Dim m_freeFile As Integer
Dim m_Url As String
Dim m_IP As String
Dim m_Port As String
Dim m_TotalBytes As Long
Dim m_sBytes As String
Dim m_CurrMP3File As String
Dim m_CurrTime As String
Dim sString As String
Dim totalBytes As Long
Dim m_bStopAfterTrack As Boolean
Private Type Icy
    ServerName As String
    Genre As String
    URL As String
    Content_Type As String
    Pub As String
    MetaInt As String
    BitRate As String
    ContactICQ As String
    ContactAim As String
    StreamTitle As String
    StreamUrl As String
End Type

Dim m_Icy As Icy
Dim m_StreamIndexLeft As Long

Dim b_init As Boolean

Const sForbidden = "\/*%°^!$?=´`~':.|><"

Const ReqHeader = "" & _
                "GET / HTTP/1.0" & vbCrLf & _
                "Host: %" & vbCrLf & _
                "User-Agent: WinampMPEG/5.09" & vbCrLf & _
                "Accept: */*" & vbCrLf & _
                "Icy-MetaData: 1" & vbCrLf & _
                "Connection: close" & vbCrLf & vbCrLf

Private Type TagInfo
  TAG As String * 3
  Songname As String * 30
  artist As String * 30
  album As String * 30
  year As String * 4
  comment As String * 30
  Genre As String * 1
End Type

Dim tTagInfo As TagInfo




Public Sub Init(sURL As String)
Dim b() As String
Dim bResult As Boolean
Dim sMsg As String

m_Url = sURL
b() = Split(Replace$(sURL, "http://", ""), ":")
m_IP = b(0)
m_Port = b(1)

wSocket.RemoteHost = m_IP
wSocket.RemotePort = m_Port
b_init = True
End Sub


Private Sub UserControl_Initialize()
UserControl.Width = Picture1.Width
UserControl.Height = Picture1.Height
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Picture1.Width
UserControl.Height = Picture1.Height
End Sub

Private Sub wSocket_Connect()
wSocket.SendData (ReqHeader)
End Sub

Private Sub wSocket_DataArrival(ByVal bytesTotal As Long)

Dim iPos As Integer
Dim sStreamTitle As String
Dim lDataStart As Long

    Dim btData() As Byte
    
    Dim sData As String
    
    Select Case sendnumber
    Case 0
        
        If IcyOK = True Then
            sendnumber = sendnumber + 1
        End If
    Case 1
        Call wSocket.GetData(btData, vbByte)
        lDataStart = ReadHeader(btData)
        m_StreamIndexLeft = Val(m_Icy.MetaInt)
        m_StreamIndexLeft = m_StreamIndexLeft - (bytesTotal - lDataStart)
        sendnumber = sendnumber + 1
        m_IcyDataReceived = True
        If b_init = True Then StopRecording: Exit Sub
        Put m_freeFile, , btData()
        
    Case 2
        Dim Data() As Byte
        Dim iMetaLength As Integer
        Dim i As Integer
        Dim p As Integer
        Call wSocket.GetData(btData, vbByte)
        m_StreamIndexLeft = m_StreamIndexLeft - bytesTotal

        If m_StreamIndexLeft <= 0 Then
            sData = CStr(btData)
            btData = MidB(sData, 1, bytesTotal + m_StreamIndexLeft - 1) 'erster teil
            
            
            Put m_freeFile, , btData()
            
            
            
            btData = MidB(sData, bytesTotal + m_StreamIndexLeft) 'zweiter teil
            lDataStart = ReadMetaData(btData)
            If lDataStart < 0 Then
                m_StreamIndexLeft = Abs(lDataStart)
                Exit Sub
            End If
            
            btData = MidB(sData, bytesTotal + m_StreamIndexLeft + lDataStart)
            
            m_StreamIndexLeft = m_Icy.MetaInt
            m_StreamIndexLeft = m_StreamIndexLeft - UBound(btData) 'die bereits erhaltenen abziehen
            
        End If

        On Error GoTo ExitSub
        Put m_freeFile, , btData()
        
        
        

        

        m_sBytes = DoConv(totalBytes + bytesTotal)
        totalBytes = totalBytes + bytesTotal
        
        m_CurrTime = GetRecordTime(CLng(totalBytes / (Val(Replace$(m_Icy.BitRate, "kbps", "", , , vbTextCompare)) * 1000 / 8)))
        
    End Select
ExitSub:
End Sub

Private Function GetRecordTime(lSek As Long) As String
Dim lMin As Long
Dim lHour As Long
If lSek < 60 Then GetRecordTime = "00:00:" & Format$(lSek, "00")
If lSek >= 60 And lSek < (60& * 60&) Then lMin = lSek \ 60: lSek = lSek - (lMin * 60): GetRecordTime = "00:" & Format(lMin, "00") & ":" & Format$(lSek, "00")
If lSek >= (60& * 60&) Then lHour = (lSek \ 3600&): lMin = (lSek - (lHour * 3600&)) \ 60: lSek = lSek - (lMin * 60) - (3600& * lHour): GetRecordTime = Format$(lHour, "00") & ":" & Format(lMin, "00") & ":" & Format$(lSek, "00")
'hier könnte man noch die Tage mit einbeziehen
'If lSek >= (60& * 60&) And lSek < (60& * 60& * 24&) Then lHour = (lSek \ 3600&): lMin = (lSek - (lHour * 3600)) \ 60: lSek = lSek - (lMin * 60): GetRecordTime = Format$(lHour, "00") & ":" & Format(lMin, "00") & ":" & Format$(lSek, "00")
End Function

Public Sub StartRecording()
m_IcyDataReceived = False
wSocket.Connect
End Sub
Property Get FileName() As String
FileName = m_File
End Property
Public Sub StopRecording()
b_init = False
sendnumber = 0

wSocket.Close
PUT_ID3V1TAG
Close #m_freeFile
ResetData
End Sub
Private Function DoConv(Number As Long) As String
If Number < 1024 Then DoConv = CStr(Number) & " B"
If Number >= 1024 And Number < (1024& * 1024&) Then DoConv = CStr(Round(Number / 1024, 2)) & " KB"
If Number >= (1024& * 1024&) And Number < (1024& * 1024& * 1024&) Then DoConv = CStr(Round(Number / 1024 / 1024, 2)) & " MB"
If Number >= (1024& * 1024& * 1024&) Then DoConv = CStr(Round(Number / 1024 / 1024 / 1024, 2)) & " GB"
End Function
Property Get CurrTime() As String
CurrTime = m_CurrTime
End Property

Property Get ServerName() As String
ServerName = m_Icy.ServerName
End Property
Property Get Genre() As String
Genre = m_Icy.Genre
End Property
Property Get BitRate() As String
BitRate = m_Icy.BitRate
End Property
Property Get CurrBytes() As String
CurrBytes = m_sBytes
End Property
Property Get StreamTitle() As String
StreamTitle = m_Icy.StreamTitle
End Property

Property Get URL() As String
URL = m_Icy.URL
End Property

Private Function IcyOK() As Boolean
Dim Dta As String
Call wSocket.GetData(Dta, vbString)
If Mid$(Dta, 5, 3) = "200" Then
    IcyOK = True
    
Else
    m_Icy.ServerName = "Server momentan nicht verfügbar"
    m_IcyDataReceived = True
    Close #m_freeFile
    
    wSocket.Close
End If

End Function

Private Function ReadHeader(Data() As Byte) As Long
    'returns datastart
    Dim sData As String
    Dim b() As String
    Dim i As Integer
    Dim sURLCorrect As String
    sData = StrConv(CStr(Data), vbUnicode)
    
    Dim sString As String
    Dim sStreamTitle As String
    Dim iPos As Integer
    
    b = Split(sData, vbCrLf)
    On Error Resume Next
    For i = 0 To 8
        If UCase$(Left$(b(i), 8)) = "ICY-NAME" Then m_Icy.ServerName = Right$(b(i), Len(b(i)) - 9)
        If UCase$(Left$(b(i), 9)) = "ICY-GENRE" Then m_Icy.Genre = Right$(b(i), Len(b(i)) - 10)
        If UCase$(Left$(b(i), 7)) = "ICY-URL" Then m_Icy.URL = Right$(b(i), Len(b(i)) - 8)
        If UCase$(Left$(b(i), 12)) = "CONTENT-TYPE" Then m_Icy.Content_Type = Right$(b(i), Len(b(i)) - 13)
        If UCase$(Left$(b(i), 7)) = "ICY-PUB" Then m_Icy.Pub = Right$(b(i), Len(b(i)) - 8)
        If UCase$(Left$(b(i), 11)) = "ICY-METAINT" Then m_Icy.MetaInt = Right$(b(i), Len(b(i)) - 12)
        If UCase$(Left$(b(i), 6)) = "ICY-BR" Then m_Icy.BitRate = Right$(b(i), Len(b(i)) - 7) & "kbps"
        If UCase$(Left$(b(i), 7)) = "ICY-ICQ" Then m_Icy.ContactICQ = Right$(b(i), Len(b(i)) - 8)
        If UCase$(Left$(b(i), 7)) = "ICY-AIM" Then m_Icy.ContactAim = Right$(b(i), Len(b(i)) - 8)
    Next
    If b_init = True Then Exit Function
    On Error GoTo 0
        m_StreamIndexLeft = Val(m_Icy.MetaInt)
        m_freeFile = FreeFile
        sURLCorrect = Trim$(Replace$(m_Icy.URL, "http://", ""))
        Call CorrectStreamName(sURLCorrect)
        m_File = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "RadioStations\" & sURLCorrect & "\" & Format$(Now, "yyyymmdd") & "\"
        
        MakeSureDirectoryPathExists m_File

        m_CurrMP3File = m_File & Format$(Now, "dd-mmm-yyyy hh-nn") & ".mp3"
        
        Open m_CurrMP3File For Binary As #m_freeFile
    ReadHeader = InStr(1, sData, vbCrLf & vbCrLf)
    If ReadHeader = 0 Then
        ReadHeader = -1
    Else
        ReadHeader = ReadHeader + 4 'hinter vbcrlf
    End If
End Function

Private Function ReadMetaData(Data() As Byte) As Long
    Dim lMetaLength As Long
    Dim sMeta As String
    Dim sTemp As String
    Dim sTitle As String
    lMetaLength = Data(0) * 16
    If UBound(Data) < lMetaLength Then
        ReadMetaData = lMetaLength - UBound(Data) - 1
        Exit Function
    Else
        ReadMetaData = lMetaLength + 1
    End If
    sMeta = StrConv(CStr(MidB(Data, 2, lMetaLength)), vbUnicode)
    If InStr(1, sMeta, "StreamTitle=") > 0 Then
        sTemp = Mid(sMeta, InStr(1, sMeta, "StreamTitle=") + 12)
        sTitle = Replace$(Left(sTemp, InStr(1, sTemp, ";") - 1), "'", "")
        If m_Icy.StreamTitle <> sTitle Then
            PUT_ID3V1TAG
            Close m_freeFile

            
            totalBytes = 0
            
            If sWinAmpLocation <> "" Then
                If Dir$(sWinAmpLocation, vbNormal) <> "" Then
                    Shell sWinAmpLocation & " /ADD " & Chr$(34) & m_CurrMP3File & Chr$(34)
                End If
            End If
            If m_bStopAfterTrack = False Then
                Dim sADD As String
                CorrectStreamName sTitle
                m_Icy.StreamTitle = sTitle
                If Dir$(m_File & "\" & m_Icy.StreamTitle & ".mp3", vbNormal) <> "" Then sADD = Format$(Now, "ddmmmyyyyhhnnss") Else sADD = ""
                m_CurrMP3File = m_File & "\" & m_Icy.StreamTitle & sADD & ".mp3"
                m_freeFile = FreeFile
                Open m_CurrMP3File For Binary As #m_freeFile
            Else
                m_bStopAfterTrack = False
                StopRecording
                Exit Function
            End If
            
        End If
    End If
    If InStr(1, sMeta, "StreamUrl=") > 0 Then
        sTemp = Mid(sMeta, InStr(1, sMeta, "StreamUrl=") + 10)
        sTemp = Replace$(Left(sTemp, InStr(1, sTemp, ";") - 1), "'", "")
        
        m_Icy.StreamUrl = sTemp
    End If
    
End Function

Property Get IcyDataReceived() As Boolean
IcyDataReceived = m_IcyDataReceived
End Property
Property Let IcyDataReceived(bVal As Boolean)
m_IcyDataReceived = bVal
End Property
Private Sub ResetData()
'm_bBufferFull = False
m_File = ""
m_freeFile = 0
m_Port = ""
m_TotalBytes = 0
m_sBytes = ""
m_CurrMP3File = ""
sString = ""
totalBytes = 0
m_StreamIndexLeft = 0
sendnumber = 0
End Sub

Public Sub CorrectStreamName(ByRef sName As String)
Dim i As Integer
For i = 1 To Len(sForbidden)
    sName = Replace(sName, Mid$(sForbidden, i, 1), "_")
Next
sName = Replace$(sName, Chr$(34), "_")

End Sub
Property Let StopAfterTrack(bVal As Boolean)
m_bStopAfterTrack = True
End Property
Property Get StopAfterTrack() As Boolean
StopAfterTrack = m_bStopAfterTrack
End Property


Public Sub PUT_ID3V1TAG()
    Dim TPos&
    Dim b() As String
    On Error GoTo ERRORS
    
    With tTagInfo
        TPos = LOF(m_freeFile)
        Get #m_freeFile, TPos - 127, .TAG
        If .TAG = "TAG" Then TPos = TPos - 127
        .TAG = "TAG"
        b = Split(m_Icy.StreamTitle, "-")
        .artist = LTrim$(RTrim$(b(0)))
        .Songname = LTrim$(RTrim$(b(1)))
        .comment = m_Icy.URL
        Put #m_freeFile, TPos, tTagInfo
        
        
    End With

    Exit Sub
ERRORS:
    
End Sub
