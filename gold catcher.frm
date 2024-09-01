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
   Begin VB.CommandButton Cmd_cred 
      BackColor       =   &H00008000&
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Tmr_colCheck 
      Interval        =   1
      Left            =   9120
      Top             =   240
   End
   Begin VB.Timer Tmr_lblUp 
      Interval        =   100
      Left            =   9600
      Top             =   240
   End
   Begin VB.CommandButton Cmd_settings 
      BackColor       =   &H00004000&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4365
      MaskColor       =   &H00004000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Timer Tmr_gravGc 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   2160
      Top             =   240
   End
   Begin VB.Timer Tmr_rndDrop 
      Enabled         =   0   'False
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
         Italic          =   0   'False
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
   Begin VB.Timer Tmr_gravGbl 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   30
      Left            =   1200
      Top             =   240
   End
   Begin VB.Timer Tmr_gametime 
      Enabled         =   0   'False
      Interval        =   500
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
         Italic          =   0   'False
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
   Begin VB.Timer Tmr_bganim 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   10080
      Top             =   240
   End
   Begin VB.Timer Tmr_menuanim 
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
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
   Begin VB.CommandButton Cmd_lang 
      BackColor       =   &H00008000&
      Caption         =   "中文"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1920
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Lbl_crlist 
      BackColor       =   &H00004000&
      Caption         =   $"gold catcher.frx":1B591
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Index           =   2
      Left            =   480
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   9780
   End
   Begin VB.Label Lbl_crlist 
      BackColor       =   &H00004000&
      Caption         =   $"gold catcher.frx":1B6D3
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1935
      Index           =   1
      Left            =   480
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   9780
   End
   Begin VB.Label Lbl_crlist 
      BackColor       =   &H00004000&
      Caption         =   "Music : Zun - Kaeidzuka ~ Higan Retour                                                       Mound of Life"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   9780
   End
   Begin VB.Image Img_eEgg 
      Height          =   1920
      Index           =   3
      Left            =   8280
      Picture         =   "gold catcher.frx":1B848
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image Img_eEgg 
      Height          =   2280
      Index           =   2
      Left            =   6360
      Picture         =   "gold catcher.frx":52470
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Img_eEgg 
      Height          =   2160
      Index           =   1
      Left            =   1920
      Picture         =   "gold catcher.frx":7A987
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Image Img_eEgg 
      Height          =   2160
      Index           =   0
      Left            =   240
      Picture         =   "gold catcher.frx":9AD5F
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Lbl_cred 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Lbl_lang 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Language toggle"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image Img_heart 
      Height          =   960
      Index           =   2
      Left            =   9600
      Picture         =   "gold catcher.frx":A2451
      Top             =   2160
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_heart 
      Height          =   960
      Index           =   1
      Left            =   8400
      Picture         =   "gold catcher.frx":A2EDB
      Top             =   2160
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_heart 
      Height          =   960
      Index           =   0
      Left            =   7200
      Picture         =   "gold catcher.frx":A3965
      Top             =   2160
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Lbl_Pscore 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score = "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Lbl_PHscore 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "High score ="
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Img_Gbl 
      Height          =   1245
      Index           =   0
      Left            =   4320
      Picture         =   "gold catcher.frx":A43EF
      Top             =   240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Lbl_score 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Score = 0"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
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
      Top             =   1320
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
      Caption         =   "High Score = 0"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
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
      Top             =   840
      Width           =   3495
   End
   Begin VB.Image Img_GC 
      Height          =   960
      Index           =   0
      Left            =   3240
      Picture         =   "gold catcher.frx":A6759
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_SpR 
      Height          =   1200
      Left            =   3240
      Picture         =   "gold catcher.frx":A71E3
      Top             =   6480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Img_SpL 
      Height          =   1200
      Left            =   3240
      Picture         =   "gold catcher.frx":A82AD
      Top             =   6480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Img_bg2 
      Height          =   7965
      Left            =   0
      Picture         =   "gold catcher.frx":A9377
      Top             =   0
      Visible         =   0   'False
      Width           =   10650
   End
   Begin VB.Image Img_bg1 
      Height          =   7965
      Left            =   0
      Picture         =   "gold catcher.frx":BA3FF
      Top             =   0
      Visible         =   0   'False
      Width           =   10650
   End
   Begin VB.Image Img_menubg1 
      Height          =   7965
      Left            =   0
      Picture         =   "gold catcher.frx":CB5FC
      Top             =   0
      Width           =   10650
   End
   Begin VB.Label Lbl_info 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coins = 515 pts move with arrow keys"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1560
      Left            =   7200
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   3285
      WordWrap        =   -1  'True
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
Dim gc As Integer
Dim life As Integer
Dim gametime As Integer
Dim Hscore As Long
Dim score As Long
Dim i As Integer
Dim kChk As Integer ' knight collision check
Dim fChk As Integer 'floor collision check
Dim paused As String
Dim gameover As String
Dim SHscore As String
Dim Sscore As String

Function Blocker(ShpB As Boolean, CmdRv As Boolean, CmdEV As Boolean, TmrBE As Boolean, CmdMV As Boolean, TmrRD As Boolean, Optional LblPCap As String) 'function to enable/disable pause screen
    Shape_blocker.Visible = ShpB
    Shape_blocker.ZOrder (0)
    Lbl_paused.Caption = LblPCap
    Lbl_paused.Visible = ShpB
    Lbl_paused.ZOrder (0)
    Lbl_PHscore.Visible = ShpB
    Lbl_PHscore.ZOrder (0)
    Lbl_Pscore.Visible = ShpB
    Lbl_Pscore.ZOrder (0)
    Cmd_mainmenu.Visible = CmdMV
    Cmd_resume.Visible = CmdRv
    Cmd_exit.Visible = CmdEV
    Tmr_bganim.Enabled = TmrBE
    Tmr_rndDrop.Enabled = TmrRD
    For i = 0 To 20
        Tmr_gravGc(i).Enabled = TmrRD
        Tmr_gravGbl(i).Enabled = TmrRD
    Next
End Function

Private Sub Form_Load()
    Dim sFile1 As String
    Tmr_gametime.Enabled = False
    Img_bg1.Visible = False
    Img_bg2.Visible = False
    Img_menubg1.Visible = False
    Lbl_Hscore.BackStyle = 1
    Lbl_score.Visible = False
    Lbl_score.BackStyle = 1
    Retval = PlaySound(App.Path & "\Kaeidzuka ~ Higan Retour.wav", 0, 1 Or 8)
    Tmr_menuanim.Enabled = True
    Img_menubg1.Visible = True
    gametime = 0
    Hscore = 0
    score = 0
    gc = 515
    life = 3
    paused = "paused"
    gameover = "game over"
    SHscore = "High Score = "
    Sscore = "Score = "
    Img_SpR.Left = Img_SpL.Left
    Randomize
    For i = 1 To 20
        Load Img_Gbl(i)
        Load Img_GC(i)
        Load Tmr_gravGbl(i)
        Load Tmr_gravGc(i)
        Img_GC(i).Top = 240
        Img_Gbl(i).Top = 240
    Next
End Sub

Private Sub Tmr_lblUp_Timer()
    Lbl_score.Caption = Sscore & score
    Lbl_Pscore.Caption = Sscore & score
    If score > Hscore Then
        Hscore = score
        Lbl_Hscore.Caption = SHscore & Hscore
        Lbl_PHscore.Caption = SHscore & Hscore
    End If
End Sub

Private Sub Tmr_colCheck_Timer()
 'coin collision
    For kChk = 0 To 20
        If Img_GC(kChk).Top >= 5520 Then
            If Img_GC(kChk).Left + Img_GC(kChk).Width >= Img_SpL.Left + 120 Then
                If Img_GC(kChk).Left <= Img_SpL.Left + 1080 Then
                    Img_GC(kChk).Visible = False
                    Img_GC(kChk).Top = 240
                    Tmr_gravGc(kChk).Enabled = False
                    'Retval = PlaySound(App.Path & "\coin.wav", 0, 1)
                    score = score + 515
                End If
            End If
        End If
    Next
    For fChk = 0 To 20
        If Img_GC(fChk).Top >= 6840 Then
            Img_GC(fChk).Visible = False
            Tmr_gravGc(fChk).Enabled = False
            Img_GC(fChk).Top = 240
            score = score - 15
        End If
    Next
    'goblin collision
    For kChk = 0 To 20
        If Img_Gbl(kChk).Top >= 5280 Then
            If Img_Gbl(kChk).Left + Img_Gbl(kChk).Width >= Img_SpL.Left + 120 Then
                If Img_Gbl(kChk).Left <= Img_SpL.Left + 1080 Then
                    Img_Gbl(kChk).Visible = False
                    Img_Gbl(kChk).Top = 240
                    Tmr_gravGbl(kChk).Enabled = False
                    'Retval = PlaySound(App.Path & "\oof.wav", 0, 1)
                    life = life - 1
                End If
            End If
        End If
    Next
    For fChk = 0 To 20
        If Img_Gbl(fChk).Top >= 6480 Then
            Img_Gbl(fChk).Top = 240
            Img_Gbl(fChk).Visible = False
            Tmr_gravGbl(fChk).Enabled = False
        End If
    Next
End Sub

Private Sub Tmr_rndDrop_Timer()
    Dim s As Integer
    Dim r As Integer
    s = Rnd * 10
    If s <= 2 Then
        r = Rnd * 20
        Img_Gbl(r).Left = Int((Rnd * 5680) + 120)
        Img_Gbl(r).Top = 240
        Img_Gbl(r).Visible = True
        Tmr_gravGbl(r).Enabled = True
    ElseIf s > 1 Then
        r = Rnd * 20
        Img_GC(r).Left = Int((Rnd * 5880) + 240)
        Img_GC(r).Top = 240
        Img_GC(r).Visible = True
        Tmr_gravGc(r).Enabled = True
    End If
End Sub

Private Sub Tmr_gravGc_Timer(Index As Integer)
    Img_GC(Index).ZOrder (0)
    Img_GC(Index).Move Img_GC(Index).Left + 0, Img_GC(Index).Top + 100
   
End Sub
Private Sub Tmr_gravGbl_Timer(Index As Integer)
    Img_Gbl(Index).ZOrder (0)
    Img_Gbl(Index).Move Img_Gbl(Index).Left + 0, Img_Gbl(Index).Top + 100
End Sub

Private Sub Tmr_menuanim_Timer()
    If Img_menubg1.Visible = True Then
        Img_menubg1.Visible = False
    ElseIf Img_menubg1.Visible = False Then
        Img_menubg1.Visible = True
        Img_menubg1.ZOrder (1)
    End If
End Sub

Private Sub Cmd_start_Click()
    Tmr_menuanim.Enabled = False
    Img_bg1.Visible = True
    Cmd_start.Visible = False
    Cmd_exit.Visible = False
    Cmd_settings.Visible = False
    Lbl_Hscore.BackStyle = 0
    Lbl_score.Visible = True
    Lbl_score.BackStyle = 0
    Lbl_info.Visible = True
    Lbl_info.ZOrder (0)
    score = 0
    Tmr_bganim.Enabled = True
    Img_bg2.Visible = True
    Img_SpL.Visible = True
    Retval = PlaySound(App.Path & "\Mound of Life.wav", 0, 1 Or 8)
    Tmr_gametime.Enabled = True
    Tmr_rndDrop.Enabled = True
    life = 3
    For i = 0 To 20
        Img_GC(i).Top = 240
        Img_Gbl(i).Top = 240
    Next
End Sub

Private Sub Tmr_gametime_Timer()
'    If gametime <= 40 Then
'        gametime = gametime + 1
'        score = score + 10
'    ElseIf gametime >= 40 Then
'        Retval = Blocker(True, False, True, False, True, False, gameover)
'        Cmd_mainmenu.Top = 5200
'        Cmd_exit.Top = 6400
'    End If
    'life system
    If life = 0 Then
        Img_heart(2).Visible = False
        Img_heart(1).Visible = False
        Img_heart(0).Visible = False
        Retval = Blocker(True, False, True, False, True, False, gameover)
        Cmd_mainmenu.Top = 5200
        Cmd_exit.Top = 6400
    ElseIf life = 3 Then
        Img_heart(2).Visible = True
        Img_heart(1).Visible = True
        Img_heart(0).Visible = True
    ElseIf life = 2 Then
        Img_heart(2).Visible = False
        Img_heart(1).Visible = True
        Img_heart(0).Visible = True
    ElseIf life = 1 Then
        Img_heart(2).Visible = False
        Img_heart(1).Visible = False
        Img_heart(0).Visible = True
    End If
End Sub

Private Sub Tmr_bganim_Timer()
    If Img_bg2.Visible = True Then
        Img_bg2.Visible = False
    ElseIf Img_bg2.Visible = False Then
        Img_bg2.Visible = True
        Lbl_info.ZOrder (0)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) 'movement
    If KeyCode = 37 And Shift = 1 Then
        Img_SpL.Visible = True
        Img_SpR.Visible = False 'sprite faces left
        Img_SpR.Left = Img_SpL.Left 'syncs the left and right facing sprites
        If Img_SpL.Left >= 150 Then
            Img_SpL.Left = Img_SpL.Left - 80
        ElseIf Img_SpL.Left <= 150 Then
            DoEvents 'sprite stops moving at edge of frame
        End If
    ElseIf KeyCode = 39 And Shift = 1 Then
        Img_SpR.Left = Img_SpL.Left
        Img_SpL.Visible = False
        Img_SpR.Visible = True 'sprite faces right
        If Img_SpR.Left <= 5950 Then
            Img_SpL.Left = Img_SpR.Left + 80
        ElseIf Img_SpR.Left >= 5950 Then
            DoEvents
        End If
    ElseIf KeyCode = 37 Then
        Img_SpL.Visible = True
        Img_SpR.Visible = False
        Img_SpR.Left = Img_SpL.Left
        If Img_SpL.Left >= 240 Then
            Img_SpL.Left = Img_SpL.Left - 240
        ElseIf Img_SpL.Left <= 240 Then
            DoEvents
        End If
    ElseIf KeyCode = 39 Then
        Img_SpL.Visible = False
        Img_SpR.Visible = True
        Img_SpL.Left = Img_SpR.Left
        If Img_SpR.Left <= 5880 Then
            Img_SpR.Left = Img_SpR.Left + 240
        ElseIf Img_SpR.Left >= 5880 Then
            DoEvents
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer) 'esc key event
    If KeyAscii = 27 And Shape_blocker.Visible = False And Img_bg1.Visible = True Then
        Retval = Blocker(True, True, False, False, True, False, paused)
        Cmd_mainmenu.Left = 4365
        Cmd_mainmenu.Top = 5400
        Tmr_gametime.Enabled = False
    ElseIf KeyAscii = 27 And Shape_blocker.Visible = True Then
        Retval = Blocker(False, False, False, False, False, True)
        Tmr_gametime.Enabled = True
    End If
End Sub

Private Sub Cmd_mainmenu_Click()
    Retval = Blocker(False, False, True, False, False, False)
    Tmr_gametime.Enabled = False
    gametime = 0
    Tmr_menuanim.Enabled = True
    Img_bg1.Visible = False
    Cmd_start.Visible = True
    Cmd_settings.Visible = True
    Cmd_exit.Top = 6600
    Lbl_Hscore.BackStyle = 1
    Lbl_score.Visible = False
    Lbl_info.Visible = False
    Img_bg2.Visible = False
    Img_SpL.Visible = False
    Img_SpR.Visible = False
    Lbl_lang.Visible = False
    Lbl_cred.Visible = False
    Cmd_lang.Visible = False
    Cmd_cred.Visible = False
    For i = 0 To 20
        Img_GC(i).Visible = False
        Img_Gbl(i).Visible = False
    Next
    For i = 0 To 2
        Img_heart(i).Visible = False
        Lbl_crlist(i).Visible = False
    Next
    For i = 0 To 3
        Img_eEgg(i).Visible = False
    Next
    Retval = PlaySound(App.Path & "\Kaeidzuka ~ Higan Retour.wav", 0, 1 Or 8)
End Sub

Private Sub Cmd_resume_Click()
    Retval = Blocker(False, False, False, False, False, True)
    Tmr_gametime.Enabled = False
End Sub

Private Sub Cmd_settings_Click()
    Retval = Blocker(False, False, False, False, True, False)
    Cmd_mainmenu.Top = 6600
    Cmd_start.Visible = False
    Cmd_settings.Visible = False
    Lbl_lang.Visible = True
    Lbl_cred.Visible = True
    Cmd_lang.Visible = True
    Cmd_cred.Visible = True
End Sub

Private Sub Cmd_exit_Click()
    Unload Me
    Retval = PlaySound("", 0, 40 Or 2)
End Sub

Private Sub Cmd_lang_Click()
    If Cmd_lang.Caption = "中文" Then
        Cmd_lang.Caption = "English"
        Cmd_start.Caption = "开始"
        Cmd_resume.Caption = "继续"
        Cmd_mainmenu.Caption = "主页"
        Cmd_settings.Caption = "设置"
        Cmd_cred.Caption = "显示"
        Cmd_exit.Caption = "退出游戏"
        Lbl_lang.Caption = "语言切换"
        Lbl_cred.Caption = "鸣谢"
        Lbl_info.Caption = "金币 = 515 分         使用方向键移动"
        Lbl_crlist(0).Caption = "音乐：Zun - 花映塚 ~ Higan Retour                                                                 此岸の塚"
        Lbl_crlist(1).Caption = "图画： NicoleMarieProductions  - 心                                                   BizmasterStudios - 金币                                                             Master484 - 骑士                                                                                                     哥布林                                                                       ansimuz - 森林"
        Lbl_crlist(2).Caption = "组员 ：曾咏晴（5）                                                                                       周宇晴（6）                                                                                       莫丰泽（29）                                                                                     沈俊达（30）"
        SHscore = "最高纪录 = "
        Lbl_Hscore.Caption = "最高纪录 = 0"
        Sscore = "积分 = "
        paused = "暂停"
        gameover = "游戏结束"
    ElseIf Cmd_lang.Caption = "English" Then
        Cmd_lang.Caption = "中文"
        Cmd_start.Caption = "start"
        Cmd_resume.Caption = "resume"
        Cmd_mainmenu.Caption = "main menu"
        Cmd_settings.Caption = "settings"
        Cmd_cred.Caption = "show"
        Cmd_exit.Caption = "exit"
        Lbl_lang.Caption = "Language toggle"
        Lbl_cred.Caption = "credits"
        Lbl_info.Caption = "Coins = 515 pts move with arrow keys"
        Lbl_crlist(0).Caption = "Music : Zun - Kaeidzuka ~ Higan Retour                                                       Mound of Life"
        Lbl_crlist(1).Caption = "Graphics : NicoleMarieProductions - heart                                           BizmasterStudios - gold coin                                                  Master484 - Knight                                                                                                Goblin                                                                      ansimuz - Forest "
        
        Lbl_crlist(2).Caption = "Groupmates : 曾咏晴（5）                                                                                       周宇晴（6）                                                                                       莫丰泽（29）                                                                                     沈俊达（30）"
        SHscore = "High Score = "
        Lbl_Hscore.Caption = "High Score = 0"
        Sscore = "Score = "
        paused = "paused"
        gameover = "game over"
    End If
End Sub

Private Sub Cmd_cred_Click()
    For i = 0 To 3
        Img_eEgg(i).Visible = True
    Next
    Lbl_lang.Visible = False
    Lbl_cred.Visible = False
    Cmd_lang.Visible = False
    Cmd_cred.Visible = False
    For i = 0 To 2
        Lbl_crlist(i).Visible = True
    Next
End Sub

Private Sub Form_Terminate()
    Retval = PlaySound("", 0, 40 Or 2)
End Sub

