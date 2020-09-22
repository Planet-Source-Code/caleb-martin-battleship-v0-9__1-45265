VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battleship v0.9"
   ClientHeight    =   4965
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   4650
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7091
            MinWidth        =   3175
            Text            =   "Click 'New game' to begin"
            TextSave        =   "Click 'New game' to begin"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6456
            Text            =   "Enemy: Not connected"
            TextSave        =   "Enemy: Not connected"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStats 
      Caption         =   "Status"
      Height          =   1275
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   7695
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         DragIcon        =   "frmMain.frx":0442
         DragMode        =   1  'Automatic
         Height          =   420
         Left            =   3600
         Picture         =   "frmMain.frx":06FC
         ScaleHeight     =   420
         ScaleWidth      =   420
         TabIndex        =   14
         Top             =   540
         Width           =   420
      End
      Begin VB.Frame fraYou 
         Caption         =   "You"
         Height          =   735
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   3375
         Begin VB.Label lblYouSunk 
            AutoSize        =   -1  'True
            Caption         =   "Sunk: 0"
            Height          =   195
            Left            =   2280
            TabIndex        =   8
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblYouHits 
            AutoSize        =   -1  'True
            Caption         =   "Hits: 0"
            Height          =   195
            Left            =   1320
            TabIndex        =   7
            Top             =   300
            Width           =   450
         End
         Begin VB.Label lblYouShots 
            AutoSize        =   -1  'True
            Caption         =   "Shots: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   300
            Width           =   585
         End
      End
      Begin VB.Frame fraEnemy 
         Caption         =   "Enemy"
         Height          =   735
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   3375
         Begin VB.Label lblEnemySunk 
            Caption         =   "Sunk: 0"
            Height          =   255
            Left            =   2280
            TabIndex        =   11
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblEnemyHits 
            AutoSize        =   -1  'True
            Caption         =   "Hits: 0"
            Height          =   195
            Left            =   1320
            TabIndex        =   10
            Top             =   300
            Width           =   450
         End
         Begin VB.Label lblEnemyShots 
            AutoSize        =   -1  'True
            Caption         =   "Shots: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   300
            Width           =   585
         End
      End
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   0
      Picture         =   "frmMain.frx":09B6
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin VB.PictureBox picEnemy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   4800
         Picture         =   "frmMain.frx":53568
         ScaleHeight     =   151
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   151
         TabIndex        =   5
         Top             =   600
         Width           =   2265
      End
      Begin VB.PictureBox picYou 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   480
         Picture         =   "frmMain.frx":642A2
         ScaleHeight     =   151
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   151
         TabIndex        =   4
         Top             =   600
         Width           =   2265
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1234
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1508
      _Version        =   393216
      Rows            =   11
      Cols            =   11
      FormatString    =   $"frmMain.frx":74FDC
   End
   Begin VB.PictureBox picHit 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   0
      Picture         =   "frmMain.frx":74FE2
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picMiss 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   0
      Picture         =   "frmMain.frx":7528C
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************
'* Battleship v0.9                *
'* Copyright © 2003 by R³Software *
'**********************************
Option Explicit

Const SND_SYNC = &H0       ' Play synchronously (default).
Const SND_ASYNC = &H1      ' Play asynchronously (see
                                   ' note below).
Const SND_NODEFAULT = &H2  ' Do not use default sound.
Const SND_MEMORY = &H4     ' lpszSoundName points to a
                                   ' memory file.
Const SND_LOOP = &H8       ' Loop the sound until next
                                   ' sndPlaySound.
Const SND_NOSTOP = &H10    ' Do not stop any currently

Dim module As Long
Dim intShipCount As Integer
Dim boolEnemyReady As Boolean
Dim boolHost As Boolean
Dim boolYourTurn As Boolean
Dim boolGameOver As Boolean
Dim intAircraft, intBattleship, intSub, intCruiser, intDestroyer
Dim intYouShots, intYouHits, intYouSunk, intEnemyShots, intEnemyHits, intEnemySunk As Integer
'Dim Grid(1 To 10, 1 To 10) As Boolean

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Private Sub Form_Load()
  'Load frmShips
  'Call FSOUND_Init(44000, 32, FSOUND_OUTPUT_WINMM)
  'module = FMUSIC_LoadSong("titan.s3m")
  'Call FMUSIC_PlaySong(module)
  intYouShots = 0
  intYouHits = 0
  intYouSunk = 0
  intEnemyShots = 0
  intEnemyHits = 0
  intEnemySunk = 0
  intAircraft = 5
  intBattleship = 4
  intCruiser = 3
  intSub = 3
  intDestroyer = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Socket.Close
  'If frmShips.Visible = True Then
  Unload frmShips
  'End If
  Unload frmAbout
  'Call FMUSIC_FreeSong(module)
  'Call FSOUND_Close
  End
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show 1, Me
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuNewGame_Click()
  On Error Resume Next
  Unload frmShips
  Socket.Close
  If MsgBox("Do you want to host this game?", vbQuestion Or vbYesNo, "Host?") = vbYes Then
    boolHost = True
    Socket.LocalPort = 1234
    Socket.Listen
    MsgBox "Your IP address is: " & Socket.LocalIP
    Status.Panels(1).Text = "Waiting for other player to join..."
  Else
    boolHost = False
    Dim strConIP As Variant
    strConIP = InputBox("Enter the IP address of the host game to connect.", "Connect")
    Socket.LocalPort = 0
    Socket.RemoteHost = strConIP
    Socket.RemotePort = 1234
    Socket.Connect
  End If
End Sub

Private Sub picEnemy_DragDrop(Source As Control, X As Single, Y As Single)
  If boolGameOver = True Then
    Exit Sub
  End If
  
  If boolYourTurn = True Then
    Socket.SendData "*SHOOT*" & X & ":" & Y
  End If
End Sub

Private Sub picYou_DragDrop(Source As Control, X As Single, Y As Single)
  Dim intLoop As Integer
  Dim intSetX As Integer, intSetY As Integer
  Dim intShipWidth As Integer, intShipHeight As Integer
  Dim intGridX As Integer, intGridY As Integer
  Dim strShipName As String
  
  intShipWidth = (Source.Width + 2) \ 15
  intShipHeight = (Source.Height + 2) \ 15
  
  Select Case X
  Case 0 To 15
    intSetX = 1
    intGridX = 1
  Case 16 To 30
    intSetX = 16
    intGridX = 2
  Case 31 To 45
    intSetX = 31
    intGridX = 3
  Case 46 To 60
    intSetX = 46
    intGridX = 4
  Case 61 To 75
    intSetX = 61
    intGridX = 5
  Case 76 To 90
    intSetX = 76
    intGridX = 6
  Case 91 To 105
    intSetX = 91
    intGridX = 7
  Case 106 To 120
    intSetX = 106
    intGridX = 8
  Case 121 To 135
    intSetX = 121
    intGridX = 9
  Case 136 To 150
    intSetX = 136
    intGridX = 10
  Case Else
    Exit Sub
  End Select
  
  Select Case Y
  Case 0 To 15
    intSetY = 1
    intGridY = 1
  Case 16 To 30
    intSetY = 16
    intGridY = 2
  Case 31 To 45
    intSetY = 31
    intGridY = 3
  Case 46 To 60
    intSetY = 46
    intGridY = 4
  Case 61 To 75
    intSetY = 61
    intGridY = 5
  Case 76 To 90
    intSetY = 76
    intGridY = 6
  Case 91 To 105
    intSetY = 91
    intGridY = 7
  Case 106 To 120
    intSetY = 106
    intGridY = 8
  Case 121 To 135
    intSetY = 121
    intGridY = 9
  Case 136 To 150
    intSetY = 136
    intGridY = 10
  Case Else
    Exit Sub
  End Select
  
  If (intSetX + Source.Width) > picYou.Width Then
    Exit Sub
  End If
  If (intSetY + Source.Height) > picYou.Height Then
    Exit Sub
  End If
  
  If intShipWidth = 1 Then
    For intLoop = intGridY To (intGridY + (intShipHeight - 1))
      If Left$(Grid.TextMatrix(intLoop, intGridX), 4) = "SHIP" Then
        Exit Sub
      End If
    Next intLoop
  End If
  
  If intShipHeight = 1 Then
    For intLoop = intGridX To (intGridX + (intShipWidth - 1))
      If Grid.TextMatrix(intGridY, intLoop) = "SHIP" Then
        Exit Sub
      End If
    Next intLoop
  End If
  
  If Source.name = "AircraftV" Or Source.name = "AircraftH" Then
    frmShips.AircraftV.Visible = False
    frmShips.AircraftH.Visible = False
    strShipName = "aircraft carrier"
  End If
  If Source.name = "BattleshipV" Or Source.name = "BattleshipH" Then
    frmShips.BattleshipV.Visible = False
    frmShips.BattleshipH.Visible = False
    strShipName = "battleship"
  End If
  If Source.name = "CruiserV" Or Source.name = "CruiserH" Then
    frmShips.CruiserV.Visible = False
    frmShips.CruiserH.Visible = False
    strShipName = "cruiser"
  End If
  If Source.name = "SubV" Or Source.name = "SubH" Then
    frmShips.SubV.Visible = False
    frmShips.SubH.Visible = False
    strShipName = "submarine"
  End If
  If Source.name = "DestroyerV" Or Source.name = "DestroyerH" Then
    frmShips.DestroyerV.Visible = False
    frmShips.DestroyerH.Visible = False
    strShipName = "destroyer"
  End If
  
  If intShipWidth = 1 Then
    For intLoop = intGridY To (intGridY + (intShipHeight - 1))
      Grid.TextMatrix(intLoop, intGridX) = "SHIP" & strShipName
    Next intLoop
  End If
  
  If intShipHeight = 1 Then
    For intLoop = intGridX To (intGridX + (intShipWidth - 1))
      Grid.TextMatrix(intGridY, intLoop) = "SHIP"
    Next intLoop
  End If
  
  'Call SetParent(Source.hwnd, picYou.hwnd)
  'Source.DragMode = 0
  'Source.Left = intSetX
  'Source.Top = intSetY
  picYou.PaintPicture Source.Picture, intSetX, intSetY

  intShipCount = intShipCount + 1
  If intShipCount >= 5 Then
    frmShips.Visible = False
    Socket.SendData "*SETUPDONE*"
    If boolEnemyReady = True Then
      If boolHost = True Then
        boolYourTurn = True
        Status.Panels(1).Text = "Your turn..."
      Else
        Status.Panels(1).Text = "Enemy's turn..."
      End If
    Else
      Status.Panels(1).Text = "Waiting on other player..."
      Status.Panels(2).Text = "Enemy: Setting up ships."
      Do
        DoEvents
      Loop Until boolEnemyReady = True
      Status.Panels(2).Text = "Enemy: Ready"
      If boolHost = True Then
        boolYourTurn = True
        Status.Panels(1).Text = "Your turn..."
      Else
        Status.Panels(1).Text = "Enemy's turn..."
      End If
    End If
  End If
  
  'Debug.Print "Name: " & Source.Name & "     X: " & X & "     Y: " & Y
End Sub

Public Sub PlayWaveRes(vntResourceID As Variant, Optional vntFlags)
  '-----------------------------------------------------------------
  ' WARNING:  If you want to play sound files asynchronously in
  '           Win32, then you MUST change bytSound() from a local
  '           variable to a module-level or static variable. Doing
  '           this prevents your array from being destroyed before
  '           sndPlaySound is complete. If you fail to do this, you
  '           will pass an invalid memory pointer, which will cause
  '           a GPF in the Multimedia Control Interface (MCI).
  '-----------------------------------------------------------------
  Dim bytSound() As Byte ' Always store binary data in byte arrays!

  bytSound = LoadResData(vntResourceID, "WAVE")

  If IsMissing(vntFlags) Then
    vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
  End If

  If (vntFlags And SND_MEMORY) = 0 Then
    vntFlags = vntFlags Or SND_MEMORY
  End If

  sndPlaySound bytSound(0), vntFlags
End Sub

Private Sub SetupShips()
  intShipCount = 0
  Load frmShips
  frmShips.Show
  Status.Panels(1).Text = "Drag and drop your ships onto the board."
  Status.Panels(2).Text = "Enemy: Setting up ships."
End Sub

Private Sub Shot(strCoords As String, strHitMiss As String)
  Dim intX, intY, intSetX, intSetY, intGridX, intGridY As Integer
  Dim boolSunk As Boolean
  
  intX = Int(Mid$(strCoords, 1, InStr(strCoords, ":") - 1))
  intY = Int(Mid$(strCoords, InStr(strCoords, ":") + 1, Len(strCoords)))
    
  Select Case intX
  Case 0 To 15
    intSetX = 1
    intGridX = 1
  Case 16 To 30
    intSetX = 16
    intGridX = 2
  Case 31 To 45
    intSetX = 31
    intGridX = 3
  Case 46 To 60
    intSetX = 46
    intGridX = 4
  Case 61 To 75
    intSetX = 61
    intGridX = 5
  Case 76 To 90
    intSetX = 76
    intGridX = 6
  Case 91 To 105
    intSetX = 91
    intGridX = 7
  Case 106 To 120
    intSetX = 106
    intGridX = 8
  Case 121 To 135
    intSetX = 121
    intGridX = 9
  Case 136 To 150
    intSetX = 136
    intGridX = 10
  Case Else
    Exit Sub
  End Select
  
  Select Case intY
  Case 0 To 15
    intSetY = 1
    intGridY = 1
  Case 16 To 30
    intSetY = 16
    intGridY = 2
  Case 31 To 45
    intSetY = 31
    intGridY = 3
  Case 46 To 60
    intSetY = 46
    intGridY = 4
  Case 61 To 75
    intSetY = 61
    intGridY = 5
  Case 76 To 90
    intSetY = 76
    intGridY = 6
  Case 91 To 105
    intSetY = 91
    intGridY = 7
  Case 106 To 120
    intSetY = 106
    intGridY = 8
  Case 121 To 135
    intSetY = 121
    intGridY = 9
  Case 136 To 150
    intSetY = 136
    intGridY = 10
  Case Else
    Exit Sub
  End Select

  If strHitMiss = "N/A" Then
    If Left$(Grid.TextMatrix(intGridY, intGridX), 4) = "SHIP" Then
      Dim strShipHit, strShipSunk As String
      intEnemyShots = intEnemyShots + 1
      intEnemyHits = intEnemyHits + 1
      strShipHit = Mid$(Grid.TextMatrix(intGridY, intGridX), 5, Len(Grid.TextMatrix(intGridY, intGridX)))
      Select Case strShipHit
      Case "aircraft carrier"
        intAircraft = intAircraft - 1
        If intAircraft <= 0 Then
          boolSunk = True
          strShipSunk = strShipHit
        End If
      Case "battleship"
        intBattleship = intBattleship - 1
        If intBattleship <= 0 Then
          boolSunk = True
          strShipSunk = strShipHit
        End If
      Case "submarine"
        intSub = intSub - 1
        If intSub <= 0 Then
          boolSunk = True
          strShipSunk = strShipHit
        End If
      Case "destroyer"
        intDestroyer = intDestroyer - 1
        If intDestroyer <= 0 Then
          boolSunk = True
          strShipSunk = strShipHit
        End If
      Case "cruiser"
        intCruiser = intCruiser - 1
        If intCruiser <= 0 Then
          boolSunk = True
          strShipSunk = strShipHit
        End If
      End Select
      
      Grid.TextMatrix(intGridY, intGridX) = "SHOT"
      picYou.PaintPicture picHit, intSetX, intSetY
      If boolSunk = True Then
        Socket.SendData "*SUNK*" & strShipSunk & ":" & intX & ":" & intY
        intEnemySunk = intEnemySunk + 1
        'MsgBox "Your " & strShipSunk & " has been sunk."
      Else
        Socket.SendData "*HIT*" & intX & ":" & intY
      End If
      PlayWaveRes 101, SND_ASYNC
    Else
      If Grid.TextMatrix(intGridY, intGridX) <> "SHOT" Then
        intEnemyShots = intEnemyShots + 1
        Grid.TextMatrix(intGridY, intGridX) = "SHOT"
        Socket.SendData "*MISS*" & intX & ":" & intY
        picYou.PaintPicture picMiss, intSetX, intSetY
        PlayWaveRes 102, SND_ASYNC
      End If
    End If
    boolYourTurn = True
    Status.Panels(1).Text = "Your turn..."
  Else
    If strHitMiss = "HIT" Then
      picEnemy.PaintPicture picHit.Picture, intSetX, intSetY
      PlayWaveRes 101, SND_ASYNC
    ElseIf strHitMiss = "MISS" Then
      picEnemy.PaintPicture picMiss.Picture, intSetX, intSetY
      PlayWaveRes 102, SND_ASYNC
    End If
  End If
  'MsgBox intSetX & " " & intSetY & vbCrLf & intGridX & " " & intGridY
End Sub

Private Sub Socket_Connect()
  Call SetupShips
End Sub

Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
  Socket.Close
  Socket.Accept requestID
  Call SetupShips
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
  Dim Data As String
  Socket.GetData Data
  If Data = "*SETUPDONE*" Then
    boolEnemyReady = True
    Status.Panels(2).Text = "Enemy: Ready."
  End If
  If Mid$(Data, 1, 7) = "*SHOOT*" Then
    Call Shot(Mid$(Data, 8, Len(Data)), "N/A")
  End If
  If Mid$(Data, 1, 5) = "*HIT*" Then
    intYouShots = intYouShots + 1
    intYouHits = intYouHits + 1
    boolYourTurn = False
    Status.Panels(1).Text = "Enemy's turn..."
    Call Shot(Mid$(Data, 6, Len(Data)), "HIT")
  End If
  If Mid$(Data, 1, 6) = "*MISS*" Then
    intYouShots = intYouShots + 1
    boolYourTurn = False
    Status.Panels(1).Text = "Enemy's turn..."
    Call Shot(Mid$(Data, 7, Len(Data)), "MISS")
  End If
  If Mid$(Data, 1, 6) = "*SUNK*" Then
    intYouShots = intYouShots + 1
    intYouHits = intYouHits + 1
    intYouSunk = intYouSunk + 1
    boolYourTurn = False
    Status.Panels(1).Text = "Enemy's turn..."
    Call Shot(Mid$(Data, InStr(Data, ":") + 1, Len(Data)), "HIT")
    MsgBox "You have sunk the enemy's " & Mid$(Data, 7, InStr(Mid$(Data, 7, Len(Data)), ":") - 1) & "!"
  End If
  If Data = "*WON*" Then
    boolGameOver = True
    Status.Panels(1).Text = "You were defeated."
    Status.Panels(2).Text = "Enemy won."
    MsgBox "You lost."
  End If
End Sub

Private Sub UpdateLabels()
  lblYouShots.Caption = "Shots: " & intYouShots
  lblYouHits.Caption = "Hits: " & intYouHits
  lblYouSunk.Caption = "Sunk: " & intYouSunk
  lblEnemyShots.Caption = "Shots: " & intEnemyShots
  lblEnemyHits.Caption = "Hits: " & intEnemyHits
  lblEnemySunk.Caption = "Sunk: " & intEnemySunk
  
  If intYouSunk >= 5 Then
    boolGameOver = True
    Socket.SendData "*WON*"
    Status.Panels(1).Text = "You won!!"
    Status.Panels(2).Text = "Enemy defeated!!"
    MsgBox "You won!"
    Timer1.Enabled = False
  End If
End Sub

Private Sub Timer1_Timer()
  Call UpdateLabels
End Sub
