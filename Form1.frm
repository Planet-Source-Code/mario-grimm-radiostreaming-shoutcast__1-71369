VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "Radiostreaming"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11370
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   11370
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CheckBox chckShoutcast 
      Caption         =   "Check1"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   3800
      Width           =   200
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Eintrag entfernen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   9
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdStations 
      Caption         =   "&Radiostation zur Liste hinzufügen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   9600
      Width           =   3855
   End
   Begin VB.ListBox lstStationDetails 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   0
      TabIndex        =   7
      Top             =   7920
      Width           =   8415
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Stoppen &nach Track"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Sofort stoppen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Aufnahme"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6800
      _Version        =   393217
      Indentation     =   18
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   240
   End
   Begin MSComctlLib.ListView lvwStream 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
   End
   Begin Projekt1.UserControl1 usrStreaming 
      Height          =   510
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shoutcast-Radiostations"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1725
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255
Dim sWholeData As String
Dim bDataReceived As Boolean
Dim sEndTag As String
Dim sShoutCastURL As String

Private Type Station
    StationName As String
    ID As String
    BitRate As String
    Genre As String
    CurrentTrack As String
    ListenerCount As String
End Type
Dim tStations(20000) As Station
Dim sStationFileLoc As String

Private Sub chckShoutcast_Click()
If chckShoutcast.Value = 0 Then
    Me.Height = chckShoutcast.Top + chckShoutcast.Height + 10 + 420
Else
    Me.Height = cmdStations.Top + cmdStations.Height + 420
End If
End Sub

Private Sub cmdAction_Click(Index As Integer)
With usrStreaming(lvwStream.SelectedItem.Index - 1)
Select Case Index
    Case 0
        .StartRecording
        cmdAction(Index).Enabled = False
        cmdAction(1).Enabled = True
        cmdAction(2).Enabled = True
    Case 1
        .StopRecording
        cmdAction(Index).Enabled = False
        cmdAction(2).Enabled = False
        cmdAction(0).Enabled = True
    Case 2
        .StopAfterTrack = .StopAfterTrack Xor True
    Case 3
        .StopRecording
        lvwStream.ListItems.Remove lvwStream.SelectedItem.Index
        SaveStations
End Select
End With
End Sub



Private Sub cmdStations_Click()
Dim sFile As String
Dim iPos As Long
Dim iEndPos As Long
If TreeView1.SelectedItem Is Nothing Then Exit Sub
If tStations(TreeView1.SelectedItem.Index).ID = "" Then Exit Sub
sShoutCastURL = "http://yp.shoutcast.com/sbin/tunein-station.pls?id=" & tStations(TreeView1.SelectedItem.Index).ID
sEndTag = "Title1"
Winsock1.Connect
bDataReceived = False
Do
DoEvents
Loop Until bDataReceived = True
Winsock1.Close
bDataReceived = False
iPos = InStr(1, sWholeData, "File1=") + 6
iEndPos = InStr(iPos, sWholeData, Chr$(10))
sFile = Mid$(sWholeData, iPos, iEndPos - iPos)
sFile = Replace$(sFile, "http://", "")

If lvwStream.FindItem(sFile) Is Nothing Then
    AddStation sFile
Else
    lvwStream.FindItem(sFile).EnsureVisible
    lvwStream.FindItem(sFile).Selected = True
End If

End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
WinAmpIsInstalled = IsWinampLocated


PaintLVWBckground lvwStream


sStationFileLoc = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\") & "Stations.txt"
ReadFile sStationFileLoc, lvwStream

FillTVWWithStations TreeView1

Timer1.Enabled = True
Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
With lvwStream
    .Width = Me.ScaleWidth - 2 * .Left
'    .Height = Me.ScaleHeight - 2 * .Top
End With
lstStationDetails.Width = lvwStream.Width
TreeView1.Width = lvwStream.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
For i = 0 To usrStreaming.UBound
    usrStreaming(i).StopRecording
Next
Unload Me
End
End Sub

Private Sub AddStation(sStation As String)
Dim i As Integer
Dim lv As ListItem
Dim t As Long
Dim f As Integer
Set lv = lvwStream.ListItems.Add(, , sStation)
Load usrStreaming(usrStreaming.UBound + 1)
i = usrStreaming.UBound
usrStreaming(i).Init sStation
If InternetGetConnectedStateEx(Ret, sConnType, 254, 0) = 1 Then
    usrStreaming(i).StartRecording
    t = Timer
    With usrStreaming(i)
    While .IcyDataReceived = False And Timer - t < 2
        DoEvents
    Wend
    If .ServerName = "" Then
        lv.SubItems(1) = "Kein Streamempfang"
        .StopRecording
    Else
        lv.SubItems(1) = .ServerName
        lv.SubItems(5) = .Genre
        lv.SubItems(6) = .URL
        lv.SubItems(7) = .BitRate
    End If
    End With
End If

f = FreeFile
Open sStationFileLoc For Append As #f
Print #f, sStation
Close #f

End Sub
Private Sub SaveStations()
Dim f As Integer
f = FreeFile
Open sStationFileLoc For Output As #f
For i = 1 To lvwStream.ListItems.Count
    Print #f, lvwStream.ListItems(i)
Next
Close #f
End Sub
Private Sub ReadFile(sFile As String, lvw As ListView)
Dim i As Integer
Dim sInhalt As String
Dim f As Integer
Dim b() As String
Dim lv As ListItem
Dim t As Long
f = FreeFile

On Error GoTo ErrorHandling

Open sFile For Binary As #f
sInhalt = Space$(LOF(f))
Get #f, , sInhalt
Close #f

b = Split(sInhalt, vbCrLf)

For i = 0 To UBound(b)
    If Trim$(b(i)) <> "" Then
        Set lv = lvw.ListItems.Add(, , b(i))


        If i > 0 Then Load usrStreaming(i)
        usrStreaming(i).Init b(i)
        
        If InternetGetConnectedStateEx(Ret, sConnType, 254, 0) = 1 Then
            usrStreaming(i).StartRecording
            t = Timer
            With usrStreaming(i)
            While .IcyDataReceived = False And Timer - t < 2
                DoEvents
            Wend
            If .ServerName = "" Then
                lv.SubItems(1) = "Kein Streamempfang"
                .StopRecording
            Else
                lv.SubItems(1) = .ServerName
                lv.SubItems(5) = .Genre
                lv.SubItems(6) = .URL
                lv.SubItems(7) = .BitRate
            End If
            End With
        End If
        

    End If
Next

ErrorHandling:

End Sub



Private Sub lvwStream_Click()
If lvwStream.SelectedItem Is Nothing Then Exit Sub
With usrStreaming(lvwStream.SelectedItem.Index - 1)
    If .CurrBytes = "" Then
        cmdAction(0).Enabled = True
        cmdAction(1).Enabled = False
        cmdAction(2).Enabled = False
    Else
        cmdAction(0).Enabled = False
        cmdAction(1).Enabled = True
        cmdAction(2).Enabled = True
    End If
End With

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
For i = 0 To usrStreaming.UBound
    If lvwStream.ListItems(i + 1).SubItems(2) <> usrStreaming(i).StreamTitle Then
        lvwStream.ListItems(i + 1).SubItems(2) = usrStreaming(i).StreamTitle
    End If
    
    
    If usrStreaming(i).CurrBytes = "" Then
        If Not lvwStream.ListItems(i + 1).SubItems(3) = "No Recording" Then
            lvwStream.ListItems(i + 1).SubItems(3) = "No Recording"
        End If
        lvwStream.ListItems(i + 1).SubItems(4) = ""
    Else
        lvwStream.ListItems(i + 1).SubItems(3) = usrStreaming(i).CurrBytes
        lvwStream.ListItems(i + 1).SubItems(4) = usrStreaming(i).CurrTime
    End If
    DoEvents
Next
optimalWidth lvwStream

End Sub
Private Sub PaintLVWBckground(lvw As ListView)
With lvw
    .ListItems.Add , , "Test"
    With Picture1
        .Height = 2 * lvw.ListItems(1).Height
        Picture1.Line (0, 0)-(.Width, lvw.ListItems(1).Height), RGB(245, 245, 245), BF
    End With
    
    Set .Picture = Picture1.Image
    .PictureAlignment = lvwTile
    .ListItems.Clear
End With
End Sub

Private Sub FillTVWWithStations(tvw As TreeView)
Dim iPos As Long
Dim i As Long
Dim iEndPos As Long
Dim tCount As Integer
Dim p As Long
Dim j As Long
Dim sStation As String
Dim sID As String
Dim sGenre As String
Dim sCurrTrack As String
Dim sBitRate As String
Winsock1.RemoteHost = "yp.shoutcast.com"
Winsock1.RemotePort = 80
sShoutCastURL = "http://yp.shoutcast.com/sbin/newxml.phtml?"
sEndTag = "</genrelist>"
Winsock1.Connect
Do
DoEvents
Loop Until bDataReceived = True
Winsock1.Close
bDataReceived = False
i = 1
Do
    j = InStr(i, sWholeData, "genre name", vbTextCompare)
    If j > 0 Then
        iPos = InStr(j, sWholeData, Chr$(34)) + 1
        iEndPos = InStr(iPos, sWholeData, Chr$(34))
        tvw.Nodes.Add , , "K" & Mid$(sWholeData, iPos, iEndPos - iPos), Mid$(sWholeData, iPos, iEndPos - iPos), 1
        tvw.Nodes.Add tvw.Nodes.Count, tvwChild, , "..."
        i = iEndPos
    End If
Loop Until j = 0

'tCount = tvw.Nodes.Count
'
'ReDim tStations(20000)
'
'For p = 1 To tCount
'    sWholeData = ""
'    sShoutCastURL = "http://yp.shoutcast.com/sbin/newxml.phtml?genre=" & tvw.Nodes(p)
'    sEndTag = "</stationlist>"
'    Winsock1.Connect
'    Do
'    DoEvents
'    Loop Until bDataReceived = True
'    Winsock1.Close
'    bDataReceived = False
'    Do
'        j = InStr(i, sWholeData, "station name", vbTextCompare)
'        If j > 0 Then
'            iPos = InStr(j, sWholeData, Chr$(34)) + 1
'            iEndPos = InStr(iPos, sWholeData, Chr$(34))
'            sStation = Mid$(sWholeData, iPos, iEndPos - iPos)
'            iPos = InStr(iEndPos, sWholeData, "id=") + 4
'            iEndPos = InStr(iPos, sWholeData, Chr$(34))
'            sID = Mid$(sWholeData, iPos, iEndPos - iPos)
'            iPos = InStr(iPos, sWholeData, "br=") + 4
'            iEndPos = InStr(iPos, sWholeData, Chr$(34))
'            sBitRate = Mid$(sWholeData, iPos, iEndPos - iPos)
'            iPos = InStr(iPos, sWholeData, "genre=") + 7
'            iEndPos = InStr(iPos, sWholeData, Chr$(34))
'            sGenre = Mid$(sWholeData, iPos, iEndPos - iPos)
'
'            iPos = InStr(iPos, sWholeData, "ct=") + 4
'            iEndPos = InStr(iPos, sWholeData, Chr$(34))
'            sCurrTrack = Mid$(sWholeData, iPos, iEndPos - iPos)
'
'            With tStations(tvw.Nodes.Count + 1)
'                .StationName = sStation
'                .ID = sID
'                .BitRate = sBitRate
'                .Genre = sGenre
'                .CurrentTrack = sCurrTrack
'            End With
'
'
'            tvw.Nodes.Add p, tvwChild, , sStation, 1
'            i = iEndPos
'        End If
'    Loop Until j = 0
'
'
'
'    DoEvents
'Next
'ReDim Preserve tStations(tvw.Nodes.Count)

End Sub



Private Sub TreeView1_Click()
With lstStationDetails
    .Clear
    If tStations(TreeView1.SelectedItem.Index).ID = "" Then Exit Sub
    .AddItem "Station name: " & tStations(TreeView1.SelectedItem.Index).StationName
    .AddItem "Genre: " & tStations(TreeView1.SelectedItem.Index).Genre
    .AddItem "Bitrate: " & tStations(TreeView1.SelectedItem.Index).BitRate
    .AddItem "Aktueller Titel: " & tStations(TreeView1.SelectedItem.Index).CurrentTrack
    .AddItem "Momentane Zuhörerzahl: " & tStations(TreeView1.SelectedItem.Index).ListenerCount
End With
End Sub

Private Sub TreeView1_DblClick()
'MsgBox tStations(TreeView1.SelectedItem.Index).StationName & vbNewLine & tStations(TreeView1.SelectedItem.Index).CurrentTrack

End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
Dim iPos As Long
Dim i As Long
Dim iEndPos As Long
Dim tCount As Integer
Dim p As Long
Dim j As Long
Dim sStation As String
Dim sID As String
Dim sGenre As String
Dim sCurrTrack As String
Dim sBitRate As String
Dim sLC As String
If Node.Child.Text = "..." Then
    TreeView1.Enabled = False
    Screen.MousePointer = 11
    TreeView1.Nodes.Remove Node.Child.Index
    sWholeData = ""
    sShoutCastURL = "http://yp.shoutcast.com/sbin/newxml.phtml?genre=" & Node.Text
    sEndTag = "</stationlist>"
    Winsock1.Connect
    Do
    DoEvents
    Loop Until bDataReceived = True
    Winsock1.Close
    bDataReceived = False
    Do
        j = InStr(i + 1, sWholeData, "station name", vbTextCompare)
        If j > 0 Then
            iPos = InStr(j, sWholeData, Chr$(34)) + 1
            iEndPos = InStr(iPos, sWholeData, Chr$(34))
            sStation = Mid$(sWholeData, iPos, iEndPos - iPos)
            iPos = InStr(iEndPos, sWholeData, "id=") + 4
            iEndPos = InStr(iPos, sWholeData, Chr$(34))
            sID = Mid$(sWholeData, iPos, iEndPos - iPos)
            iPos = InStr(iPos, sWholeData, "br=") + 4
            iEndPos = InStr(iPos, sWholeData, Chr$(34))
            sBitRate = Mid$(sWholeData, iPos, iEndPos - iPos)
            iPos = InStr(iPos, sWholeData, "genre=") + 7
            iEndPos = InStr(iPos, sWholeData, Chr$(34))
            sGenre = Mid$(sWholeData, iPos, iEndPos - iPos)
            iPos = InStr(iPos, sWholeData, "ct=") + 4
            iEndPos = InStr(iPos, sWholeData, Chr$(34))
            sCurrTrack = Mid$(sWholeData, iPos, iEndPos - iPos)
            iPos = InStr(iPos, sWholeData, "lc=") + 4
            iEndPos = InStr(iPos, sWholeData, Chr$(34))
            sLC = Mid$(sWholeData, iPos, iEndPos - iPos)
            
            With tStations(TreeView1.Nodes.Count + 1)
                .StationName = sStation
                .ID = sID
                .BitRate = sBitRate
                .Genre = sGenre
                .CurrentTrack = sCurrTrack
                .ListenerCount = sLC
            End With
            
            
            TreeView1.Nodes.Add Node.Index, tvwChild, , sStation, 1
            i = iEndPos
        End If
    Loop Until j = 0
    Screen.MousePointer = 0
    TreeView1.Enabled = True
End If
End Sub

Private Sub Winsock1_Connect()
  Dim Cmd$, URL$
     
    URL = sShoutCastURL
    Cmd = "GET " & URL & " HTTP/1.0" & vbCrLf & "Accept: */*" & _
          vbCrLf & "Accept: text/html" & vbCrLf & vbCrLf
          
    Winsock1.SendData Cmd

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
Winsock1.GetData sData, vbString
sWholeData = sWholeData & sData
If InStr(1, sWholeData, sEndTag, vbTextCompare) Then
    bDataReceived = True
    Winsock1.Close
End If
End Sub

