VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Burning"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdInfo 
      Caption         =   "?"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.HScrollBar hscSpeed 
      Height          =   255
      LargeChange     =   50
      Left            =   2280
      Max             =   500
      Min             =   10
      TabIndex        =   15
      Top             =   1320
      Value           =   100
      Width           =   1290
   End
   Begin VB.PictureBox picDestination 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1050
      Index           =   1
      Left            =   1680
      ScaleHeight     =   1020
      ScaleWidth      =   1860
      TabIndex        =   4
      Top             =   240
      Width           =   1890
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "6"
      Height          =   495
      Index           =   5
      Left            =   2520
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "5"
      Height          =   495
      Index           =   4
      Left            =   2040
      Style           =   1  'Grafisch
      TabIndex        =   11
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "4"
      Height          =   495
      Index           =   3
      Left            =   1560
      Style           =   1  'Grafisch
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "3"
      Height          =   495
      Index           =   2
      Left            =   1080
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   600
      Style           =   1  'Grafisch
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "[1]"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      ScaleHeight     =   825
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picArchive 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2100
      Left            =   1680
      Picture         =   "frmMain.frx":1FF2
      ScaleHeight     =   2040
      ScaleWidth      =   1860
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      Height          =   2610
      Left            =   3360
      Picture         =   "frmMain.frx":2EE2
      ScaleHeight     =   2550
      ScaleWidth      =   930
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Timer T1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Burning"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox picDestination 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   240
      ScaleHeight     =   510
      ScaleWidth      =   930
      TabIndex        =   3
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblStatic 
      AutoSize        =   -1  'True
      Caption         =   "Speed:"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   14
      Top             =   1335
      Width           =   510
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Zentriert
      Caption         =   "..."
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblStatic 
      AutoSize        =   -1  'True
      Caption         =   "Background:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Animated Picture
' ================
'
' © 2004 Stefan Ebert
'
'
'
Option Explicit

Private Type tagSpriteDefinition
  Width                       As Long
  Height                      As Long
  CntAnimations               As Long
  ActualStep                  As Long
End Type

Private tSprite               As tagSpriteDefinition
Private m_iBackground         As Integer



Private Sub Form_Load()

  ' Init...
  With tSprite
    .ActualStep = 1
    .CntAnimations = 4
    .Width = 62
    .Height = 34
  End With
  
  Show
  Refresh
  
  optBackground(1).Value = True 'Click...
  
End Sub



Private Sub Form_Unload(Cancel As Integer)

  ' Stop timer
  If (T1.Enabled) Then T1.Enabled = False
  
  Unload Me

End Sub



Private Sub cmdInfo_Click()

  MsgBox "Animated Picture" & vbCrLf & vbCrLf & "© 2004 Stefan Ebert" & Space$(10), vbInformation

End Sub



Private Sub cmdStart_Click()

  ' Setup timer
  T1.Interval = hscSpeed.Value
  T1.Enabled = True

  cmdStart.Enabled = False

End Sub



Private Sub hscSpeed_Scroll()
  
  ' Realtime changes
  hscSpeed_Change

End Sub

Private Sub hscSpeed_Change()

  ' Speed = Timer interval
  T1.Interval = hscSpeed.Value
  
End Sub



Private Sub optBackground_Click(Index As Integer)
' Change background picture

  ' Save for later (timer function)
  m_iBackground = Index

  ' Show immediately
  If Not (T1.Enabled) Then
    If (m_iBackground = 0) Then
      picDestination(0).Cls
      picDestination(1).Cls
    Else
      ' COPY background to original sized picture
      BitBlt picDestination(0).hdc, 0, 0, tSprite.Width, tSprite.Height, picBackground.hdc, 0, (m_iBackground - 1) * tSprite.Height, SRCCOPY
      ' STRETCH background to double sized picture
      StretchBlt picDestination(1).hdc, 0, 0, tSprite.Width * 2, tSprite.Height * 2, picBackground.hdc, 0, (m_iBackground - 1) * tSprite.Height, tSprite.Width, tSprite.Height, SRCCOPY
    End If
  End If

End Sub



Private Sub T1_Timer()

  ' Drawing a sprite need 4 steps:

  ' 1) create a copy of those part of the background which is
  '    part to overwrite. Copy to a temp buffer, in this case 'picBuffer'
  If (m_iBackground = 0) Then
    picBuffer.Cls
  Else
    BitBlt picBuffer.hdc, 0, 0, tSprite.Width, tSprite.Height, picBackground.hdc, 0, (m_iBackground - 1) * tSprite.Height, SRCCOPY
  End If
  
  ' 2) copy black mask to buffer
  BitBlt picBuffer.hdc, 0, 0, tSprite.Width, tSprite.Height, picArchive.hdc, tSprite.Width, tSprite.Height * (tSprite.ActualStep - 1), SRCAND
  
  ' 3) copy inverted sprite to buffer
  BitBlt picBuffer.hdc, 0, 0, tSprite.Width, tSprite.Height, picArchive.hdc, 0, tSprite.Height * (tSprite.ActualStep - 1), SRCINVERT
  
  ' 4) buffer is ready - copy buffer to original
  BitBlt picDestination(0).hdc, 0, 0, tSprite.Width, tSprite.Height, picBuffer.hdc, 0, 0, SRCCOPY
  StretchBlt picDestination(1).hdc, 0, 0, tSprite.Width * 2, tSprite.Height * 2, picBuffer.hdc, 0, 0, tSprite.Width, tSprite.Height, SRCCOPY


  ' Count-up
  lblCount.Caption = tSprite.ActualStep
  tSprite.ActualStep = tSprite.ActualStep + 1
  If (tSprite.ActualStep > tSprite.CntAnimations) Then tSprite.ActualStep = 1

End Sub
