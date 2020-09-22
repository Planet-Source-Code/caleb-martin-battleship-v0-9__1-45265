VERSION 5.00
Begin VB.Form frmShips 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ships"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDirection 
      Caption         =   "Direction"
      Height          =   1035
      Left            =   0
      TabIndex        =   13
      Top             =   1860
      Width           =   1935
      Begin VB.PictureBox picWEC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1080
         Picture         =   "frmShips.frx":0000
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   17
         Top             =   180
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picNSC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1080
         Picture         =   "frmShips.frx":1DF2
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   16
         Top             =   180
         Width           =   750
      End
      Begin VB.OptionButton Option1 
         Caption         =   "N-S"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   300
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "W-E"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   660
         Width           =   675
      End
   End
   Begin VB.Frame fraShips 
      Caption         =   "Ships"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.PictureBox picWE 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1515
         Left            =   60
         ScaleHeight     =   101
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         Begin VB.PictureBox AircraftH 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            Picture         =   "frmShips.frx":3BE4
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   74
            TabIndex        =   12
            Top             =   0
            Width           =   1110
         End
         Begin VB.PictureBox BattleshipH 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            Picture         =   "frmShips.frx":4866
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   59
            TabIndex        =   11
            Top             =   300
            Width           =   885
         End
         Begin VB.PictureBox CruiserH 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            Picture         =   "frmShips.frx":5280
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   44
            TabIndex        =   10
            Top             =   600
            Width           =   660
         End
         Begin VB.PictureBox SubH 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            Picture         =   "frmShips.frx":59FA
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   44
            TabIndex        =   9
            Top             =   900
            Width           =   660
         End
         Begin VB.PictureBox DestroyerH 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            Picture         =   "frmShips.frx":6174
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   28
            TabIndex        =   8
            Top             =   1260
            Width           =   420
         End
      End
      Begin VB.PictureBox picNS 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1515
         Left            =   60
         ScaleHeight     =   101
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         Begin VB.PictureBox DestroyerV 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   1440
            Picture         =   "frmShips.frx":664E
            ScaleHeight     =   28
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   6
            Top             =   0
            Width           =   210
         End
         Begin VB.PictureBox SubV 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   1080
            Picture         =   "frmShips.frx":6B60
            ScaleHeight     =   44
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   5
            Top             =   0
            Width           =   210
         End
         Begin VB.PictureBox CruiserV 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   720
            Picture         =   "frmShips.frx":7332
            ScaleHeight     =   44
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   4
            Top             =   0
            Width           =   210
         End
         Begin VB.PictureBox BattleshipV 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   885
            Left            =   360
            Picture         =   "frmShips.frx":7B04
            ScaleHeight     =   59
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   3
            Top             =   0
            Width           =   210
         End
         Begin VB.PictureBox AircraftV 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   1110
            Left            =   0
            Picture         =   "frmShips.frx":856A
            ScaleHeight     =   74
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   2
            Top             =   0
            Width           =   210
         End
      End
   End
End
Attribute VB_Name = "frmShips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  frmShips.Top = frmMain.Top
  frmShips.Left = frmMain.Left - frmShips.Width
End Sub

Private Sub Option1_Click()
  picNS.Visible = True
  picWE.Visible = False
  picNSC.Visible = True
  picWEC.Visible = False
End Sub

Private Sub Option2_Click()
  picNS.Visible = False
  picWE.Visible = True
  picNSC.Visible = False
  picWEC.Visible = True
End Sub
