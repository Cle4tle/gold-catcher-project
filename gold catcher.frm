VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10665
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "gold catcher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "gold catcher.frx":0A8A
   ScaleHeight     =   7995
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_stagger 
      Interval        =   500
      Left            =   1680
      Top             =   240
   End
   Begin VB.CommandButton Cmd_mainmenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      Caption         =   "main menu"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer_drop 
      Interval        =   800
      Left            =   1200
      Top             =   240
   End
   Begin VB.Timer Timer_gametime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   240
   End
   Begin VB.CommandButton Cmd_resume 
      BackColor       =   &H00004000&
      Caption         =   "Resume"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer_bganim 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   10080
      Top             =   240
   End
   Begin VB.Timer Timer_menuanim 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton Cmd_exit 
      BackColor       =   &H00004000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Cmd_start 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   964
      Left            =   4365
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Lbl_score 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Score = 00000000"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   435
      Left            =   7080
      TabIndex        =   5
      Top             =   840
      Width           =   3480
   End
   Begin VB.Label Lbl_paused 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Caption         =   "Paused"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   3165
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Shape Shape_blocker 
      FillStyle       =   2  'Horizontal Line
      Height          =   8055
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
   End
   Begin VB.Label Lbl_Hscore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "High Score = 00000000"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.Image Img_Gbl 
      Height          =   1245
      Index           =   5
      Left            =   5640
      Picture         =   "gold catcher.frx":1B591
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Img_Gbl 
      Height          =   1245
      Index           =   4
      Left            =   4560
      Picture         =   "gold catcher.frx":1D8FB
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Img_Gbl 
      Height          =   1245
      Index           =   3
      Left            =   3480
      Picture         =   "gold catcher.frx":1FC65
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Img_Gbl 
      Height          =   1245
      Index           =   2
      Left            =   2400
      Picture         =   "gold catcher.frx":21FCF
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Img_Gbl 
      Height          =   1245
      Index           =   1
      Left            =   1320
      Picture         =   "gold catcher.frx":24339
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Img_Gbl 
      Height          =   1245
      Index           =   0
      Left            =   240
      Picture         =   "gold catcher.frx":266A3
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Img_GC 
      Height          =   960
      Index           =   6
      Left            =   6000
      Picture         =   "gold catcher.frx":28A0D
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_GC 
      Height          =   960
      Index           =   5
      Left            =   5040
      Picture         =   "gold catcher.frx":29497
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_GC 
      Height          =   960
      Index           =   4
      Left            =   4080
      Picture         =   "gold catcher.frx":29F21
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_GC 
      Height          =   960
      Index           =   3
      Left            =   3120
      Picture         =   "gold catcher.frx":2A9AB
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_GC 
      Height          =   960
      Index           =   2
      Left            =   2160
      Picture         =   "gold catcher.frx":2B435
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_GC 
      Height          =   960
      Index           =   1
      Left            =   1200
      Picture         =   "gold catcher.frx":2BEBF
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_GC 
      Height          =   960
      Index           =   0
      Left            =   240
      Picture         =   "gold catcher.frx":2C949
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_SpR 
      Height          =   1200
      Left            =   2280
      Picture         =   "gold catcher.frx":2D3D3
      Top             =   6480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Img_SpL 
      Height          =   1200
      Left            =   2280
      Picture         =   "gold catcher.frx":2E49D
      Top             =   6480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Img_bg2 
      Height          =   7965
      Left            =   0
      Picture         =   "gold catcher.frx":2F567
      Top             =   0
      Visible         =   0   'False
      Width           =   10650
   End
   Begin VB.Image Img_bg1 
      Height          =   7965
      Left            =   0
      Picture         =   "gold catcher.frx":405EF
      Top             =   0
      Visible         =   0   'False
      Width           =   10650
   End
   Begin VB.Image Img_menubg1 
      Height          =   7965
      Left            =   0
      Picture         =   "gold catcher.frx":517EC
      Top             =   0
      Width           =   10650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Retval As Long
Private Declare Function PlaySound Lib "winmm" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Dim GC As Integer
Dim Gbl As Integer
Dim gametime As Integer
Dim Hscore As Integer
Dim score As Integer
Dim x As Integer
Dim y As Integer
Dim z As Integer

Function Blocker(ShpB As Boolean, LblPv As Boolean, CmdRv As Boolean, CmdEV As Boolean, TmrBE As Boolean, CmdMV As Boolean, Optional LblPCap As String)
    Shape_blocker.Visible = ShpB
    Lbl_paused.Caption = LblPCap
    Lbl_paused.Visible = LblPv
    Cmd_mainmenu.Visible = CmdMV
    Cmd_resume.Visible = CmdRv
    Cmd_exit.Visible = CmdEV
    Timer_bganim.Enabled = TmrBE
End Function

Private Sub Form_Load()
    Img_bg1.Visible = False
    Img_bg2.Visible = False
    Img_menubg1.Visible = False
    Lbl_Hscore.BackStyle = 1
    Lbl_score.BackStyle = 1
    'Retval = PlaySound(App.Path & "\Kaeidzuka ~ Higan Retour.wav", 0, 1 Or 8)
    Timer_menuanim.Enabled = True
    Img_menubg1.Visible = True
    gametime = 0
    Randomize
    GC = Int(Rnd * 100 * 5)
    Gbl = Int(Rnd * -100 * 5)
End Sub

Private Sub Timer_menuanim_Timer()
    If Img_menubg1.Visible = True Then
        Img_menubg1.Visible = False
    ElseIf Img_menubg1.Visible = False Then
        Img_menubg1.Visible = True
    End If
End Sub

Private Sub Cmd_start_Click()
    Timer_menuanim.Enabled = False
    Img_bg1.Visible = True
    Cmd_start.Visible = False
    Cmd_exit.Visible = False
    Lbl_Hscore.BackStyle = 0
    Lbl_score.BackStyle = 0
    Timer_bganim.Enabled = True
    Img_bg2.Visible = True
    Img_SpL.Visible = True
    'Retval = PlaySound(App.Path & "\Mound of Life.wav", 0, 1 Or 8)
    Timer_gametime.Enabled = True
End Sub

Private Sub Timer_gametime_Timer()
    If gametime <= 5 Then
        gametime = gametime + 1
        score = score + 500
        Lbl_score.Caption = "Score =" & score
    ElseIf gametime >= 5 Then
        Retval = Blocker(True, True, False, True, False, True, "game over")
    End If
End Sub

Private Sub Timer_bganim_Timer()
    If Img_bg2.Visible = True Then
        Img_bg2.Visible = False
    ElseIf Img_bg2.Visible = False Then
        Img_bg2.Visible = True
    End If
End Sub

Private Sub Timer_stagger_Timer()
    x = Int(Rnd * 6)
    y = Int(Rnd * 6)

    
End Sub

Private Sub Timer_drop_Timer()
    Img_Gbl(x).Visible = True
    Img_GC(y).Visible = True
    Img_Gbl(x).Move Img_Gbl(x).Top - 200
    Img_GC(y).Move Img_GC(y).Top - 200
    
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 And Shift = 1 Then
        Img_SpL.Visible = True
        Img_SpR.Visible = False
        Img_SpL.Left = Img_SpL.Left
        If Img_SpL.Left <= 120 Then
            DoEvents
        Else
            Img_SpL.Left = Img_SpL.Left - 80
        End If
    ElseIf KeyCode = 39 And Shift = 1 Then
        Img_SpR.Left = Img_SpL.Left
        Img_SpL.Visible = False
        Img_SpR.Visible = True
        If Img_SpR.Left >= 5950 Then
            DoEvents
        Else
            Img_SpL.Left = Img_SpL.Left + 80
        End If
    ElseIf KeyCode = 37 Then
        Img_SpL.Visible = True
        Img_SpR.Visible = False
        Img_SpL.Left = Img_SpL.Left
        If Img_SpL.Left <= 120 Then
            DoEvents
        Else
            Img_SpL.Left = Img_SpL.Left - 200
        End If
    ElseIf KeyCode = 39 Then
        Img_SpR.Left = Img_SpL.Left
        Img_SpL.Visible = False
        Img_SpR.Visible = True
        If Img_SpR.Left >= 5950 Then
            DoEvents
        Else
            Img_SpL.Left = Img_SpL.Left + 200
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 And Shape_blocker.Visible = False And Cmd_start.Visible = False Then
        Retval = Blocker(True, True, True, True, False, False, "paused")
        Timer_gametime.Enabled = False
    ElseIf KeyAscii = 27 And Shape_blocker.Visible = True Then
        Retval = Blocker(False, False, False, False, False, True)
        Timer_gametime.Enabled = True
    End If
End Sub

Private Sub Cmd_mainmenu_Click()
    Retval = Blocker(False, False, False, True, False, False)
    Timer_gametime.Enabled = False
    gametime = 0
    Timer_menuanim.Enabled = True
    Img_bg1.Visible = False
    Cmd_start.Visible = True
    Lbl_Hscore.BackStyle = 1
    Lbl_score.BackStyle = 1
    Img_bg2.Visible = False
    Img_SpL.Visible = False
    Img_SpR.Visible = False
    'Retval = PlaySound(App.Path & "\Kaeidzuka ~ Higan Retour.wav", 0, 1 Or 8)
End Sub

Private Sub Cmd_resume_Click()
    Retval = Blocker(False, False, False, False, False, True)
    Timer_gametime.Enabled = False
End Sub

Private Sub Cmd_exit_Click()
    Unload Me
    Retval = PlaySound("", 0, 40 Or 2)
End Sub

Private Sub Form_Terminate()
    Retval = PlaySound("", 0, 40 Or 2)
End Sub
