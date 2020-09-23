VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cardz"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic_board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      Height          =   8055
      Left            =   0
      ScaleHeight     =   533
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   597
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin MSComctlLib.ProgressBar pro_shre 
         Height          =   375
         Left            =   0
         TabIndex        =   54
         Top             =   7260
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   51
         Scrolling       =   1
      End
      Begin VB.CommandButton cmd_shre 
         Caption         =   "Shuffle and Re-rack"
         Height          =   375
         Left            =   0
         TabIndex        =   53
         Top             =   7620
         Width           =   1635
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   51
         Left            =   1860
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   52
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   50
         Left            =   1680
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   51
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   49
         Left            =   1500
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   50
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   48
         Left            =   1320
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   49
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   47
         Left            =   1140
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   48
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   46
         Left            =   960
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   47
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   45
         Left            =   780
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   46
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   44
         Left            =   600
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   45
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   43
         Left            =   420
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   44
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   42
         Left            =   240
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   43
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   41
         Left            =   7620
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   42
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   40
         Left            =   7440
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   41
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   39
         Left            =   7260
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   40
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   38
         Left            =   7080
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   39
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   37
         Left            =   6900
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   38
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   36
         Left            =   6720
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   37
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   35
         Left            =   6540
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   34
         Left            =   6360
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   35
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   33
         Left            =   6180
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   32
         Left            =   6000
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   33
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   31
         Left            =   5820
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   30
         Left            =   5640
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   29
         Left            =   5460
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   28
         Left            =   5280
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   27
         Left            =   5100
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   26
         Left            =   4920
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   25
         Left            =   4740
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   24
         Left            =   4560
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   23
         Left            =   4380
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   22
         Left            =   4200
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   21
         Left            =   4020
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   20
         Left            =   3840
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   19
         Left            =   3660
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   18
         Left            =   3480
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   17
         Left            =   3300
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   16
         Left            =   3120
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   15
         Left            =   2940
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   14
         Left            =   2760
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   13
         Left            =   2580
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   12
         Left            =   2400
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   11
         Left            =   2220
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   10
         Left            =   2040
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   9
         Left            =   1860
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   8
         Left            =   1680
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   7
         Left            =   1500
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   6
         Left            =   1320
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   5
         Left            =   1140
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   4
         Left            =   960
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   3
         Left            =   780
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   2
         Left            =   600
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   1
         Left            =   420
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pic_card 
         AutoSize        =   -1  'True
         Height          =   1335
         Index           =   0
         Left            =   240
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   7995
         Left            =   0
         Picture         =   "frm_main.frx":1CCA
         Top             =   0
         Width           =   8955
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MOVEMENT HANDLING
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim facepath As String
Dim backpath As String

'SOUND API
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub cmd_shre_Click()
    shre 0, 0
End Sub

Private Sub shre(left, top)
    Dim play As Integer
    Dim shre_card As Integer
    For shre_card = 0 To 51
        play = play + 1
        If play = 10 Then
            play = 1
            sndPlaySound App.Path & "\Sounds\pickup.wav", &H1 Or &H2
        End If
        pro_shre = shre_card
        With pic_card(shre_card)
            If .left <> left Then .left = left
            If .top <> top Then .top = top
            DoEvents
            toggle_card shre_card, False
            Randomize
            pic_card(Int((51) * Rnd)).ZOrder 0
        End With
    Next shre_card
    pro_shre = 0
End Sub

Private Sub Form_Paint()
    Static loaded As Boolean
    If Not loaded Then
        DoEvents
        facepath = App.Path & "\graphics\cards\faces"
        backpath = App.Path & "\graphics\cards\backs\eqypt.gif"
        load_deck facepath, backpath
        loaded = True
    End If
End Sub

Private Sub Form_Resize()
    'RESIZE THE BOARD TO FILL THE FORM
    pic_board.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub pic_card_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        pic_card(index).ZOrder 0
        ReleaseCapture
        sndPlaySound App.Path & "\Sounds\pickup.wav", &H1 Or &H2
        SendMessage pic_card(index).hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
        sndPlaySound App.Path & "\Sounds\putdown.wav", &H1 Or &H2
    Else
        sndPlaySound App.Path & "\Sounds\putdown.wav", &H1 Or &H2
        toggle_card index, , True
    End If
End Sub

Private Sub toggle_card(index As Integer, Optional showhide As Boolean, Optional toggle As Boolean)
    Static hidden(51) As Boolean                                    'CARDS HIDEN FLAGS
    If toggle Then showhide = Not hidden(index)
    If showhide Then                                                'EVALUATE CARDS HIDDEN FLAG
        pic_card(index).Picture = LoadPicture(pic_card(index).Tag)  'SHOW CARD
        hidden(index) = True                                        'TURN HIDDEN FLAG ON
    Else
        pic_card(index).Picture = LoadPicture(backpath)             'HIDE CARD
        hidden(index) = False                                       'TURN HIDDEN FLAG OFF
    End If
End Sub

'LOADS THE DECK FROM A DIRECTORY
Public Sub load_deck(facepath As String, backpath As String)
    Dim fs, f, fc, f1                                   'FILE SYSTEM VARIABLES
    Dim file_count As Integer                           'FILE COUNTER
    Set fs = CreateObject("Scripting.FileSystemObject") 'SETUP THE FILE SYSTEM OBJECT
    Set f = fs.getfolder(facepath)                      'REFERENCE F TO THE FILE SYSTEM PATH
    Set fc = f.Files                                    'REFERENCE FC TO ALL FILES IN F
    For Each f1 In fc
        pic_card(file_count).Picture = LoadPicture(facepath & "\" & f1.Name)
        pic_card(file_count).Tag = facepath & "\" & f1.Name
        file_count = file_count + 1
        If file_count = 53 Then Exit Sub
    Next
    shre 0, 0
End Sub
