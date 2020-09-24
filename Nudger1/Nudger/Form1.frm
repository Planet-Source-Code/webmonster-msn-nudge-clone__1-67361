VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "SHAKE FORM"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7035
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   150
      Top             =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click the picture above"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1440
      Picture         =   "Form1.frx":593B
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   0
      Picture         =   "Form1.frx":5DB5
      Top             =   6840
      Width           =   3705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Msn nudger clone by Webmonster aka James
' Shaking by Tim Miron code here http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=12924&lngWId=1

Public Flg1 As Integer
Public FTOP As Integer
Public FLEFT As Integer

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H1
Private Sub Form_Load()
Flg1 = 0
End Sub
Private Sub Image2_Click()
Timer1.Enabled = True
soundfile$ = App.Path & "\nudge.wav" ' Plays sound
wflags% = SND_ASYNC Or SND_NODEFAULT
HaHa = sndPlaySound(soundfile$, wflags%)
End Sub

Private Sub Timer1_Timer()
Select Case Flg1 ' Shaking by Tim Miron
Case 0
FTOP = Form1.Top
FLEFT = Form1.Left
Form1.Left = Form1.Left + 30
Form1.Top = Form1.Top + 30
Flg1 = Flg1 + 1
Case 1
Form1.Left = Form1.Left - 45
Form1.Top = Form1.Top - 45
Flg1 = Flg1 + 1
Case 2
Form1.Left = Form1.Left + 60
Form1.Top = Form1.Top + 60
Flg1 = Flg1 + 1
Case 3
Form1.Left = Form1.Left - 75
Form1.Top = Form1.Top - 75
Flg1 = Flg1 + 1
Case 4
Form1.Left = Form1.Left + 90
Form1.Top = Form1.Top + 90
Flg1 = Flg1 + 1
Case 5
Form1.Left = Form1.Left - 105
Form1.Top = Form1.Top - 105
Flg1 = Flg1 + 1
Case 6
Form1.Left = Form1.Left + 105
Form1.Top = Form1.Top + 105
Flg1 = Flg1 + 1
Case 7
Form1.Left = Form1.Left - 75
Form1.Top = Form1.Top - 75
Flg1 = Flg1 + 1
Case 8
Form1.Left = Form1.Left + 90
Form1.Top = Form1.Top + 90
Flg1 = Flg1 + 1
Case 9
Form1.Left = Form1.Left - 135
Form1.Top = Form1.Top - 135
Flg1 = Flg1 + 1
Case 10
Form1.Left = Form1.Left + 90
Form1.Top = Form1.Top + 90
Flg1 = Flg1 + 1
Case 11
Form1.Left = Form1.Left - 105
Form1.Top = Form1.Top - 105
Flg1 = Flg1 + 1
Case 12
Form1.Left = Form1.Left + 135
Form1.Top = Form1.Top + 135
Flg1 = Flg1 + 1
Case 13
Form1.Left = Form1.Left - 90
Form1.Top = Form1.Top - 90
Flg1 = Flg1 + 1
Case 14
Form1.Left = Form1.Left + 75
Form1.Top = Form1.Top + 75
Flg1 = Flg1 + 1
Case 15
Form1.Left = Form1.Left - 150
Form1.Top = Form1.Top - 150
Flg1 = Flg1 + 1
Case 16
Form1.Left = Form1.Left + 105
Form1.Top = Form1.Top + 105
Flg1 = Flg1 + 1
Case 17
Form1.Left = Form1.Left - 75
Form1.Top = Form1.Top - 75
Flg1 = Flg1 + 1
Case 18
Form1.Left = Form1.Left + 90
Form1.Top = Form1.Top + 90
Flg1 = Flg1 + 1
Case 19
Form1.Left = Form1.Left - 105
Form1.Top = Form1.Top - 105
Flg1 = Flg1 + 1
Case 20
Form1.Left = Form1.Left + 135
Form1.Top = Form1.Top + 135
Flg1 = Flg1 + 1
Case 21
Form1.Left = Form1.Left - 150
Form1.Top = Form1.Top - 150
Flg1 = Flg1 + 1
Case 22
Form1.Left = Form1.Left + 180
Form1.Top = Form1.Top + 180
Flg1 = Flg1 + 1
Case 23
Form1.Left = Form1.Left - 150
Form1.Top = Form1.Top - 150
Flg1 = Flg1 + 1
Case 24
Form1.Left = Form1.Left + 195
Form1.Top = Form1.Top + 195
Flg1 = Flg1 + 1
Case 25
Form1.Left = FLEFT
Form1.Top = FTOP
Flg1 = 0
Timer1.Enabled = False
End Select
End Sub
Private Sub form_unload(cancel As Integer)
stopthesoundnow = sndPlaySound(soundfile$, wflags%)
End Sub
