VERSION 5.00
Begin VB.Form frmMidi 
   BackColor       =   &H00788C9E&
   BorderStyle     =   0  'None
   ClientHeight    =   8685
   ClientLeft      =   -3885
   ClientTop       =   990
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMidi.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmMidi.frx":08E2
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H0003008E&
      HasDC           =   0   'False
      Height          =   240
      Left            =   405
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   712
      TabIndex        =   21
      Top             =   5100
      Width           =   10680
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   1335
      Left            =   10110
      Picture         =   "frmMidi.frx":77F66
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   7
      Top             =   6345
      Width           =   1425
      Begin VB.Image Image1 
         Height          =   195
         Left            =   645
         Picture         =   "frmMidi.frx":7A5C6
         Top             =   225
         Width           =   165
      End
      Begin VB.Image Image2 
         Height          =   195
         Left            =   645
         Picture         =   "frmMidi.frx":7A706
         Top             =   675
         Width           =   165
      End
      Begin VB.Image Image3 
         Height          =   195
         Left            =   90
         Picture         =   "frmMidi.frx":7A846
         Top             =   1140
         Width           =   165
      End
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   840
      Left            =   3555
      Picture         =   "frmMidi.frx":7A986
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   6
      Top             =   6840
      Width           =   375
      Begin VB.Image Image5 
         Height          =   105
         Left            =   15
         Picture         =   "frmMidi.frx":7B3E6
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00788C9E&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   810
      IntegralHeight  =   0   'False
      Left            =   480
      TabIndex        =   2
      Top             =   6855
      Width           =   3390
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   15
      Left            =   11115
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   14
      Left            =   10380
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   13
      Left            =   9645
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   12
      Left            =   8910
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   11
      Left            =   8175
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   10
      Left            =   7440
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   9
      Left            =   6705
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   8
      Left            =   5970
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   7
      Left            =   5235
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   6
      Left            =   4500
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   5
      Left            =   3765
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   4
      Left            =   3030
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   3
      Left            =   2295
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   2
      Left            =   1560
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   1
      Left            =   825
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin MidiPlayer.chCtl chCtl1 
      Height          =   4515
      Index           =   0
      Left            =   90
      Top             =   90
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   7964
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H0000C000&
      HasDC           =   0   'False
      Height          =   330
      Left            =   405
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   715
      TabIndex        =   1
      Top             =   4770
      Width           =   10725
   End
   Begin MidiPlayer.VuBar VuBar1 
      Height          =   2310
      Left            =   4320
      TabIndex        =   0
      Top             =   5625
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   4075
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   1320
      Left            =   1935
      Picture         =   "frmMidi.frx":7B8B2
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   5
      Top             =   2295
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   1890
      Left            =   630
      Picture         =   "frmMidi.frx":7E3AE
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   4
      Top             =   1170
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      Height          =   240
      Left            =   135
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   23
      Top             =   3690
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblOption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   240
      Index           =   0
      Left            =   465
      TabIndex        =   22
      Top             =   7710
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   11130
      Picture         =   "frmMidi.frx":84E62
      Top             =   4905
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   9735
      TabIndex        =   20
      Top             =   5895
      Width           =   420
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Time Length  00:00:00"
      ForeColor       =   &H00FFFF00&
      Height          =   165
      Left            =   8460
      TabIndex        =   19
      Top             =   5625
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FF8080&
      Height          =   165
      Index           =   5
      Left            =   11265
      TabIndex        =   18
      Top             =   5760
      Width           =   210
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FF8080&
      Height          =   165
      Index           =   4
      Left            =   11265
      TabIndex        =   17
      Top             =   5940
      Width           =   210
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FF8080&
      Height          =   165
      Index           =   3
      Left            =   11265
      TabIndex        =   16
      Top             =   5580
      Width           =   210
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Velocity"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   165
      Index           =   2
      Left            =   10710
      TabIndex        =   15
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0001B3FE&
      Height          =   195
      Left            =   9030
      TabIndex        =   14
      Top             =   7980
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   8550
      Shape           =   3  'Circle
      Top             =   5940
      Width           =   120
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Transpose"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   165
      Index           =   0
      Left            =   10530
      TabIndex        =   13
      Top             =   5580
      Width           =   675
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   165
      Index           =   1
      Left            =   10725
      TabIndex        =   12
      Top             =   5940
      Width           =   480
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   "Stopped"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0001B3FE&
      Height          =   195
      Left            =   8730
      TabIndex        =   11
      Top             =   5895
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   1
      Left            =   450
      TabIndex        =   10
      Top             =   5843
      Width           =   3795
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   2
      Left            =   450
      TabIndex        =   9
      Top             =   6075
      Width           =   3795
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   0
      Left            =   450
      TabIndex        =   8
      Top             =   5610
      Width           =   3795
   End
   Begin VB.Image Image4 
      Height          =   180
      Left            =   3930
      Picture         =   "frmMidi.frx":85746
      Top             =   8115
      Width           =   225
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0001B3FE&
      Height          =   195
      Left            =   2835
      TabIndex        =   3
      Top             =   7980
      Width           =   930
   End
End
Attribute VB_Name = "frmMidi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_READONLY = &H1
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const MIDI_STOPPED = &H0
Private Const MIDI_PLAYING = &H1
Private Const MIDI_PAUSED = &H2
Private Const MIDI_INVALIDFILE = &H3
Private Const MIDI_CLOSED = &H4
Private Const MIDI_FILEOPEN = &H5

Dim WithEvents midi As MidiDll.clsMidi
Attribute midi.VB_VarHelpID = -1
Dim ControlEnable(14) As Boolean, CurrentControl As Integer, LastControl As Integer
Dim X1 As Integer, UndValue As Integer, MouseDown As Boolean, CurrentChannel As Integer
Private Sub OpenFileOnStartup(ByVal MidiFile As String)

    MousePointer = 11
    DoEvents
    Picture1.Cls
    Picture6.Cls
    For i = 0 To 15
        chCtl1(i).Celeste = 0
        chCtl1(i).Chorus = 0
        chCtl1(i).Phaser = 0
        chCtl1(i).Tremulo = 0
    Next
    midi.OpenMidiFile MidiFile
    For i = 0 To 2
        lblTitle(i) = midi.Titles(i)
    Next
    List1.ListIndex = midi.Instrument(0)
    CurrentChannel = 0
    Picture7.PaintPicture LoadResPicture(129, 0), 0, 0, 16, 16, 0, 0, 16, 16
    Picture7.Move 31, 514
    Picture7.Visible = True
    SetPatch 0
    MousePointer = 0
    midi.PlayMidi

End Sub

Sub SetPatch(ByVal Channel As Integer)
    
    On Error Resume Next
    List1.ListIndex = midi.Instrument(Channel)

End Sub
Sub Control3()
    
    If List1.TopIndex > 0 Then List1.TopIndex = List1.TopIndex - 1
    
End Sub
Sub Control4()
    
    If List1.TopIndex < List1.ListCount - 1 Then List1.TopIndex = List1.TopIndex + 1

End Sub
Private Sub DrawButtonUp(ByVal ButtonId As Integer)
    
    Select Case ButtonId
        Case 3
            Picture4.PaintPicture LoadResPicture(107, 0), 0, 0
            If ControlEnable(ButtonId) Then Control3
        Case 4
            Picture4.PaintPicture LoadResPicture(108, 0), 0, 41
            If ControlEnable(ButtonId) Then Control4
        Case 5
            PaintPicture LoadResPicture(109, 0), 561, 463
            If ControlEnable(ButtonId) Then Control5
        Case 6
            PaintPicture LoadResPicture(111, 0), 561, 427
            If ControlEnable(ButtonId) Then midi.PlayMidi
        Case 7
            PaintPicture LoadResPicture(113, 0), 588, 449
            If ControlEnable(ButtonId) Then midi.Pause
        Case 8
            PaintPicture LoadResPicture(115, 0), 605, 480
            If ControlEnable(ButtonId) Then midi.StopMidi
        Case 9
            PaintPicture LoadResPicture(117, 0), 645, 522
            If ControlEnable(ButtonId) Then Control9
        Case 10
            PaintPicture LoadResPicture(119, 0), 690, 522
            If ControlEnable(ButtonId) Then Control10
    End Select
 
End Sub
Private Sub DrawButtonDown(ByVal ButtonId As Integer)
    
    Select Case ButtonId
        Case 3
            Picture4.PaintPicture LoadResPicture(105, 0), 0, 0
        Case 4
            Picture4.PaintPicture LoadResPicture(106, 0), 0, 41
        Case 5
            PaintPicture LoadResPicture(110, 0), 561, 463
        Case 6
            PaintPicture LoadResPicture(112, 0), 561, 426
        Case 7
            PaintPicture LoadResPicture(114, 0), 588, 448
        Case 8
            PaintPicture LoadResPicture(116, 0), 605, 480
        Case 9
            PaintPicture LoadResPicture(118, 0), 645, 522
        Case 10
            PaintPicture LoadResPicture(120, 0), 690, 522
    End Select
    If ButtonId > 0 Then LastControl = ButtonId

End Sub
Private Sub GetControl(ByVal X As Single, ByVal Y As Single)
    
    On Error Resume Next
    CurrentControl = 0
    If ((Y >= 454) And (Y <= 550)) And ((X >= 128) And (X <= 261)) Then
        Select Case Picture3.Point(X - 137, Y - 457)
            Case RGB(255, 0, 0)
                CurrentControl = 3
            Case RGB(255, 255, 0)
                CurrentControl = 4
        End Select
    End If
    
    If ((Y >= 422) And (Y <= 517)) And ((X >= 558) And (X <= 643)) Then
        Select Case Picture2.Point(X - 561, Y - 426)
            Case RGB(0, 255, 255)
                CurrentControl = 5
            Case RGB(255, 255, 0)
                CurrentControl = 6
            Case RGB(255, 0, 255)
                CurrentControl = 7
            Case RGB(255, 0, 0)
                CurrentControl = 8
        End Select
    End If
    
    If ((Y >= 518) And (Y <= 553)) And ((X >= 644) And (X <= 776)) Then
        Select Case Picture2.Point(X - 561, Y - 426)
            Case RGB(0, 255, 0)
                CurrentControl = 9
            Case RGB(0, 0, 255)
                CurrentControl = 10
        End Select
    End If

End Sub
Private Function MouseOver(ByVal X As Single, ByVal Y As Single) As Boolean
    
    Select Case CurrentControl
        Case 3
            MouseOver = (Picture3.Point(X - 137, Y - 457) = RGB(255, 0, 0))
        Case 4
            MouseOver = (Picture3.Point(X - 137, Y - 457) = RGB(255, 255, 0))
        Case 5
            MouseOver = (Picture2.Point(X - 561, Y - 426) = RGB(0, 255, 255))
        Case 6
            MouseOver = (Picture2.Point(X - 561, Y - 426) = RGB(255, 255, 0))
        Case 7
            MouseOver = (Picture2.Point(X - 561, Y - 426) = RGB(255, 0, 255))
        Case 8
            MouseOver = (Picture2.Point(X - 561, Y - 426) = RGB(255, 0, 0))
        Case 9
            MouseOver = (Picture2.Point(X - 561, Y - 426) = RGB(0, 255, 0))
        Case 10
            MouseOver = (Picture2.Point(X - 561, Y - 426) = RGB(0, 0, 255))
    End Select

End Function
Sub RepaintControls()

    Select Case LastControl
        Case 3
            Picture4.PaintPicture LoadResPicture(107, 0), 0, 0
        Case 4
            Picture4.PaintPicture LoadResPicture(108, 0), 0, 41
        Case 5
            PaintPicture LoadResPicture(109, 0), 561, 463
        Case 6
            PaintPicture LoadResPicture(111, 0), 561, 427
        Case 7
            PaintPicture LoadResPicture(113, 0), 588, 449
        Case 8
            PaintPicture LoadResPicture(115, 0), 605, 480
        Case 9
            PaintPicture LoadResPicture(117, 0), 645, 522
        Case 10
            PaintPicture LoadResPicture(119, 0), 690, 522
    End Select

End Sub
Public Function ShowOpen(ByVal hForm As Long, ByVal Filter As String, ByVal Title As String, ByVal InitDir As String) As String
 
    Dim ofn As OPENFILENAME
    Dim a As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hForm
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    a = GetOpenFileName(ofn)
    If (a) Then
        ShowOpen = Trim$(ofn.lpstrFile)
    Else
        ShowOpen = ""
    End If

End Function
Sub FillListPatches()
    
    List1.AddItem "Acoustic Grand"
    List1.AddItem "Bright Acoustic"
    List1.AddItem "Electric Grand"
    List1.AddItem "Honky-Tonk"
    List1.AddItem "Electric Piano 1"
    List1.AddItem "Electric Piano 2"
    List1.AddItem "Harpsichord"
    List1.AddItem "Clavinet"
    List1.AddItem "Celesta"
    List1.AddItem "Glockenspiel"
    List1.AddItem "Music Box"
    List1.AddItem "Vibraphone"
    List1.AddItem "Marimba"
    List1.AddItem "Xylophone"
    List1.AddItem "Tubular Bells"
    List1.AddItem "Dulcimer"
    List1.AddItem "Drawbar Organ"
    List1.AddItem "Percussive Organ"
    List1.AddItem "Rock Organ"
    List1.AddItem "Church Organ"
    List1.AddItem "Reed Organ"
    List1.AddItem "Accoridan"
    List1.AddItem "Harmonica"
    List1.AddItem "Tango Accordian"
    List1.AddItem "Nylon String Guitar"
    List1.AddItem "Steel String Guitar"
    List1.AddItem "Electric Jazz Guitar"
    List1.AddItem "Electric Clean Guitar"
    List1.AddItem "Electric Muted Guitar"
    List1.AddItem "Overdriven Guitar"
    List1.AddItem "Distortion Guitar"
    List1.AddItem "Guitar Harmonics"
    List1.AddItem "Acoustic Bass"
    List1.AddItem "Electric Bass(finger)"
    List1.AddItem "Electric Bass(pick)"
    List1.AddItem "Fretless Bass"
    List1.AddItem "Slap Bass 1"
    List1.AddItem "Slap Bass 2"
    List1.AddItem "Synth Bass 1"
    List1.AddItem "Synth Bass 2"
    List1.AddItem "Violin"
    List1.AddItem "Viola"
    List1.AddItem "Cello"
    List1.AddItem "Contrabass"
    List1.AddItem "Tremolo Strings"
    List1.AddItem "Pizzicato Strings"
    List1.AddItem "Orchestral Strings"
    List1.AddItem "Timpani"
    List1.AddItem "String Ensemble 1"
    List1.AddItem "String Ensemble 2"
    List1.AddItem "SynthStrings 1"
    List1.AddItem "SynthStrings 2"
    List1.AddItem "Choir Aahs"
    List1.AddItem "Voice Oohs"
    List1.AddItem "Synth Voice"
    List1.AddItem "Orchestra Hit"
    List1.AddItem "Trumpet"
    List1.AddItem "Trombone"
    List1.AddItem "Tuba"
    List1.AddItem "Muted Trumpet"
    List1.AddItem "French Horn"
    List1.AddItem "Brass Section"
    List1.AddItem "SynthBrass 1"
    List1.AddItem "SynthBrass 2"
    List1.AddItem "Soprano Sax"
    List1.AddItem "Alto Sax"
    List1.AddItem "Tenor Sax"
    List1.AddItem "Baritone Sax"
    List1.AddItem "Oboe"
    List1.AddItem "English Horn"
    List1.AddItem "Bassoon"
    List1.AddItem "Clarinet"
    List1.AddItem "Piccolo"
    List1.AddItem "Flute"
    List1.AddItem "Recorder"
    List1.AddItem "Pan Flute"
    List1.AddItem "Blown Bottle"
    List1.AddItem "Skakuhachi"
    List1.AddItem "Whistle"
    List1.AddItem "Ocarina"
    List1.AddItem "Lead 1 (square)"
    List1.AddItem "Lead 2 (sawtooth)"
    List1.AddItem "Lead 3 (calliope)"
    List1.AddItem "Lead 4 (chiff)"
    List1.AddItem "Lead 5 (charang)"
    List1.AddItem "Lead 6 (voice)"
    List1.AddItem "Lead 7 (fifths)"
    List1.AddItem "Lead 8 (bass+lead)"
    List1.AddItem "Pad 1 (new age)"
    List1.AddItem "Pad 2 (warm)"
    List1.AddItem "Pad 3 (polysynth)"
    List1.AddItem "Pad 4 (choir)"
    List1.AddItem "Pad 5 (bowed)"
    List1.AddItem "Pad 6 (metallic)"
    List1.AddItem "Pad 7 (halo)"
    List1.AddItem "Pad 8 (sweep)"
    List1.AddItem "FX 1 (rain)"
    List1.AddItem "FX 2 (soundtrack)"
    List1.AddItem "FX 3 (crystal)"
    List1.AddItem "FX 4 (atmosphere)"
    List1.AddItem "FX 5 (brightness)"
    List1.AddItem "FX 6 (goblins)"
    List1.AddItem "FX 7 (echoes)"
    List1.AddItem "FX 8 (sci-fi)"
    List1.AddItem "Sitar"
    List1.AddItem "Banjo"
    List1.AddItem "Shamisen"
    List1.AddItem "Koto"
    List1.AddItem "Kalimba"
    List1.AddItem "Bagpipe"
    List1.AddItem "Fiddle"
    List1.AddItem "Shanai"
    List1.AddItem "Tinkle Bell"
    List1.AddItem "Agogo"
    List1.AddItem "Steel Drums"
    List1.AddItem "Woodblock"
    List1.AddItem "Taiko Drum"
    List1.AddItem "Melodic Tom"
    List1.AddItem "Synth Drum"
    List1.AddItem "Reverse Cymbal"
    List1.AddItem "Guitar Fret Noise"
    List1.AddItem "Breath Noise"
    List1.AddItem "Seashore"
    List1.AddItem "Bird Tweet"
    List1.AddItem "Telephone Ring"
    List1.AddItem "Helicopter"
    List1.AddItem "Applause"
    List1.AddItem "Gunshot"

End Sub
Private Sub chCtl1_BalanceChange(Index As Integer, ByVal Value As Integer)

    midi.Balance(Index) = Value

End Sub
Private Sub chCtl1_CelesteChange(Index As Integer, ByVal Value As Integer)

    midi.Celeste(Index) = Value

End Sub
Private Sub chCtl1_ChorusChange(Index As Integer, ByVal Value As Integer)

    midi.Chorus(Index) = Value

End Sub
Private Sub chCtl1_MuteChange(Index As Integer, ByVal Value As Boolean)
    
    midi.MuteChannel(Index) = Value
    If Value Then
        chCtl1(Index).NoteEvent = ""
        chCtl1(Index).Led = False
    End If

End Sub
Private Sub chCtl1_PhaserChange(Index As Integer, ByVal Value As Integer)

    midi.Phaser(Index) = Value

End Sub
Private Sub chCtl1_TremuloChange(Index As Integer, ByVal Value As Integer)

    midi.Tremulo(Index) = Value

End Sub
Private Sub chCtl1_VolumeChange(Index As Integer, ByVal Value As Integer)

    midi.Volume(Index) = Value

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    X1 = ScaleX(X, 1, 3)

End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CurrentControl = 11
    If Button = 1 Then
        Picture5_MouseMove Button, Shift, Image1.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 15
    End If

End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
            
    midi.Transpose = UndValue

End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    X1 = ScaleX(X, 1, 3)

End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    CurrentControl = 12
    If Button = 1 Then
        Picture5_MouseMove Button, Shift, Image2.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 45
    End If

End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    midi.Velocity = UndValue * -1
    
End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    X1 = ScaleX(X, 1, 3)

End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    CurrentControl = 13
    If Button = 1 Then
        Picture5_MouseMove Button, Shift, Image3.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 76
    End If

End Sub
Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not ControlEnable(14) Then Exit Sub
    midi.Pause
    X1 = ScaleX(X, 1, 3)
    CurrentControl = 14

End Sub
Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Not ControlEnable(14) Then Exit Sub
    If Button = 1 Then
        Form_MouseMove Button, Shift, Image4.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 541
    End If

End Sub
Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Not ControlEnable(14) Then Exit Sub
    Value = Int(((Image4.Left - 261) * 100) / 320)
    midi.SetCurrentPosition = Int((Value * midi.TimeLength) / 100)
    midi.PlayMidi
    CurrentControl = 0
    
End Sub
Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseDown = True
    X1 = ScaleX(Y, 1, 3)
    CurrentControl = 15

End Sub
Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        Picture4_MouseMove Button, Shift, 1, Image5.Top + ScaleY(Y, 1, 3)
    End If

End Sub
Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseDown = False
    CurrentControl = 0
    
End Sub
Private Sub Image6_Click()

    midi.StopMidi
    Load frmText
    frmText.Text1 = midi.LyricText
    frmText.Label1 = midi.Titles(0)
    frmText.Show 1
    Unload frmText
    Set frmText = Nothing

End Sub
Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image6 = LoadResPicture(122, 0)

End Sub
Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6 = LoadResPicture(121, 0)

End Sub
Private Sub List1_Click()
    
    midi.ChangePatch List1.ListIndex, CurrentChannel
    If Not MouseDown Then Image5.Top = Int(List1.TopIndex / 6.88) + 16

End Sub
Private Sub Control9()

    midi.StopMidi
    Load frmDevices
    On Error Resume Next
    Dim clt As New Collection, nDevs As Integer
    Set clt = midi.ListDevices(nDevs)
    For Each dName In clt
        frmDevices.List1.AddItem dName
    Next
    If frmDevices.List1.ListCount > 0 Then
        lIndex = Val(GetSetting("midiPlayer", "Devices", "DeviceIndex", "0"))
        frmDevices.List1.ListIndex = lIndex
        midi.SetDevice = lIndex
    End If
    frmDevices.Show 1
    Unload frmDevices
    Set frmDevices = Nothing
    midi.SetDevice = Val(GetSetting("midiPlayer", "Devices", "DeviceIndex", "0"))
    
End Sub
Private Sub Control5()

    If midi.Status <> MIDI_STOPPED Then midi.StopMidi
    IniDir = GetSetting("midiPlayer", "Files", "LastDir", App.Path)
    MidiFile = ShowOpen(hWnd, "Midi Files (*.mid *.kar)|*.mid;*.kar", Caption, IniDir)
    If Trim(MidiFile) = "" Then Exit Sub
    SaveSetting "midiPlayer", "Files", "LastDir", MidiFile
    DoEvents
    Picture1.Cls
    Picture6.Cls
    For i = 0 To 15
        chCtl1(i).Celeste = 0
        chCtl1(i).Chorus = 0
        chCtl1(i).Phaser = 0
        chCtl1(i).Tremulo = 0
    Next
    midi.OpenMidiFile MidiFile
    For i = 0 To 2
        lblTitle(i) = midi.Titles(i)
    Next
    List1.ListIndex = midi.Instrument(0)
    CurrentChannel = 0
    Picture7.PaintPicture LoadResPicture(129, 0), 0, 0, 16, 16, 0, 0, 16, 16
    Picture7.Move 31, 514
    Picture7.Visible = True
    SetPatch 0
    midi.PlayMidi
    

End Sub
Private Sub Control10()

    midi.StopMidi
    DoEvents
    End
    
End Sub
Private Sub Form_Load()
    
    On Error Resume Next
    Move (Screen.Width / 2) - ((Screen.TwipsPerPixelX * 800) / 2), (Screen.Height / 2) - ((Screen.TwipsPerPixelY * 600) / 2), Screen.TwipsPerPixelX * 800, Screen.TwipsPerPixelY * 600
    Dim lIndex As Integer
    Set midi = New MidiDll.clsMidi
    FillListPatches
    For t = 1 To 16
        chCtl1(t - 1).Channel = t
    Next
    List1.ListIndex = 0
    ControlEnable(3) = True
    ControlEnable(4) = True
    ControlEnable(5) = True
    ControlEnable(6) = False
    ControlEnable(7) = False
    ControlEnable(8) = False
    ControlEnable(9) = True
    ControlEnable(10) = True
    ControlEnable(11) = True
    ControlEnable(12) = True
    ControlEnable(13) = True
    ControlEnable(14) = False
    
    If Trim(Command$) <> "" Then
        DoEvents
        FileToOpen = Command$
        FileToOpen = IIf(Left(FileToOpen, 1) = Chr(34), Right(FileToOpen, Len(FileToOpen) - 1), FileToOpen)
        FileToOpen = IIf(Right(FileToOpen, 1) = Chr(34), Left(FileToOpen, Len(FileToOpen) - 1), FileToOpen)
        Open FileToOpen For Input As #1
        Line Input #1, a$
        Close #1
        Signature = Left(a$, 4)
        If Trim(Signature) = "MThd" Then
            Me.Visible = True
            OpenFileOnStartup FileToOpen
        Else
            MsgBox "Invalid File Format.", vbInformation, "MidiPlayer"
        End If
    End If
        
    
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    If ((Y >= 514) And (Y <= 530)) And ((X >= 31) And (X <= 159)) Then
        Vl = Int((X - 31) / 16) + 1
        If Vl > 8 Then Exit Sub
        CurrentChannel = Vl - 1
        Picture7.PaintPicture LoadResPicture(129, 0), 0, 0, 16, 16, ((Vl * 16) - 16), 0, 16, 16
        Picture7.Move ((Vl * 16) - 16) + 31, 514
        Picture7.Visible = True
        SetPatch CurrentChannel
        Exit Sub
    End If
    If ((Y >= 531) And (Y <= 547)) And ((X >= 31) And (X <= 159)) Then
        Vl = Int((X - 31) / 16) + 1
        If Vl > 8 Then Exit Sub
        CurrentChannel = (Vl - 1) + 8
        Picture7.PaintPicture LoadResPicture(129, 0), 0, 0, 16, 16, ((Vl * 16) - 16), 17, 16, 16
        Picture7.Move ((Vl * 16) - 16) + 31, 531
        Picture7.Visible = True
        SetPatch CurrentChannel
        Exit Sub
    End If
    
    GetControl X, Y
    If Not ControlEnable(CurrentControl) Then
        CurrentControl = 0
        Exit Sub
    End If
    If CurrentControl > 0 Then
        DrawButtonDown CurrentControl
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    If ((Y >= 514) And (Y <= 547)) And ((X >= 31) And (X <= 159)) Then Exit Sub
    
    If Not ControlEnable(CurrentControl) Then
        CurrentControl = 0
        Exit Sub
    End If
    
    If ((Y >= 541) And (Y <= 553)) And ((X >= 262) And (X <= 606)) Then
        If Button = 1 Then
            If CurrentControl = 14 Then
                px = X - X1
                px = IIf(px < 262, 262, px)
                px = IIf(px > 582, 582, px)
                Image4.Move px - 1, 541
            End If
        End If
    Else
        If Button = 1 Then
            If MouseOver(X, Y) Then
                DrawButtonDown CurrentControl
            Else
                RepaintControls
            End If
        End If
    End If
    
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    If ((Y >= 514) And (Y <= 547)) And ((X >= 31) And (X <= 159)) Then Exit Sub
    
    If Not ControlEnable(CurrentControl) Then
        CurrentControl = 0
        Exit Sub
    End If
    If MouseOver(X, Y) Then
        DrawButtonUp CurrentControl
    End If
    CurrentControl = 0
        
End Sub
Private Sub List1_Scroll()
    
    If Not MouseDown Then Image5.Top = Int(List1.TopIndex / 6.88) + 16

End Sub
Private Sub Midi_BalanceChange(ByVal Channel As Integer, ByVal Value As Integer)

    chCtl1(Channel).Balance = Value

End Sub

Private Sub Midi_ChannelMuteChange(ByVal Channel As Integer, ByVal bMute As Boolean)

    chCtl1(Channel).Mute = IIf(bMute, 1, 0)

End Sub
Private Sub midi_MasterVolumeChange(ByVal Value As Integer)

    lblInfo(4) = Value
    Image3.Left = Int(Value * 7.5) + 6

End Sub
Private Sub midi_MidiInfo(ByVal TimeLength As String, ByVal TimeSignature As String)

    lblTime = "Time Length  " & TimeLength
    Label1 = TimeSignature

End Sub
Private Sub midi_MidiLevel(ByVal Channel As Integer, ByVal Level As Integer)

    VuBar1.Value(Channel) = Level

End Sub
Private Sub midi_NotesPlaying(ByVal Channel As Integer, ByVal NoteValue As Integer, ByVal NoteText As String, ByVal DrumsValue As Integer, ByVal DrumsText As String, ByVal NoteOn As Boolean)
    
    chCtl1(Channel).Led = NoteOn
    If NoteOn Then
        chCtl1(Channel).NoteEvent = NoteText
    Else
        chCtl1(Channel).NoteEvent = ""
    End If

End Sub
Private Sub midi_Progress(ByVal midiTicks As Long, ByVal midiTime As String)
    
    On Error Resume Next
    Label2 = midiTime
    Static mTicks
    If midiTicks <> mTicks Then
        mTicks = midiTicks
        lblPercent = Int((midiTicks * 100) / midi.TimeLength) & " %"
        Lft = IIf(Int((319 * ((midiTicks * 100) / midi.TimeLength)) / 100) + 262 < 262, 262, Int((319 * ((midiTicks * 100) / midi.TimeLength)) / 100) + 262)
        Image4.Left = Lft
    End If
    
End Sub
Private Sub midi_Status(ByVal Status As Integer, ByVal StatusText As String)
    
    lblStatus = StatusText
    Shape1.FillColor = &HFF
    Select Case Status
        Case MIDI_FILEOPEN
            ControlEnable(6) = True
            ControlEnable(7) = False
            ControlEnable(8) = False
            ControlEnable(14) = True
        Case MIDI_PLAYING
            ControlEnable(6) = False
            ControlEnable(7) = True
            ControlEnable(8) = True
            ControlEnable(14) = True
            Shape1.FillColor = &HFF00&
        Case MIDI_STOPPED
            ControlEnable(6) = True
            ControlEnable(7) = False
            ControlEnable(8) = False
            ControlEnable(14) = True
        Case MIDI_PAUSED
            ControlEnable(6) = True
            ControlEnable(7) = False
            ControlEnable(8) = False
            ControlEnable(14) = True
            Shape1.FillColor = &HFFFF&
        Case MIDI_INVALIDFILE
            ControlEnable(6) = False
            ControlEnable(7) = False
            ControlEnable(8) = False
            ControlEnable(14) = False
        Case Else
            ControlEnable(6) = False
            ControlEnable(7) = False
            ControlEnable(8) = False
            ControlEnable(14) = False
    End Select

End Sub
Private Sub midi_TextKar(ByVal TxtString As String, ByVal TextStart As Integer, ByVal AtualPhrase As String, ByVal NextPhrase As String)
    
    On Error Resume Next
    Static Phrase1 As String, Phrase2 As String
    If Phrase1 <> AtualPhrase Then
        Picture1.Cls
        Phrase1 = AtualPhrase
    End If
    If Phrase2 <> NextPhrase Then
        Picture6.Cls
        Phrase2 = AtualPhrase
    End If
    Picture1.CurrentX = (Picture1.ScaleWidth / 2) - (Picture1.TextWidth(AtualPhrase) / 2)
    Picture1.CurrentY = 0
    Picture1.ForeColor = &HFF&
    
    If TextStart <> -1 Then
        Picture1.Print Left(AtualPhrase, TextStart);
        Picture1.ForeColor = &HFFFFFF
        Picture1.Print TxtString;
        Picture1.ForeColor = &HFF&
        Picture1.Print Right(AtualPhrase, Len(AtualPhrase) - (TextStart + Len(TxtString)))
    Else
        Picture1.Cls
    End If
    
    Picture6.CurrentY = 0
    Picture6.CurrentX = (Picture6.ScaleWidth / 2) - (Picture6.TextWidth(NextPhrase) / 2)
    If Trim(NextPhrase) <> "" Then
        Picture6.Print NextPhrase
    Else
        Picture6.Cls
    End If

End Sub
Private Sub midi_TransposeChange(ByVal Value As Integer)

    lblInfo(3) = Value
    Image1.Left = Int((6.25 * Value) + 43.5)

End Sub
Private Sub midi_VelocityChange(ByVal Value As Integer)

    lblInfo(5) = Value * -1
    Image2.Left = Int((6.25 * (Value * -1)) + 43.5)

End Sub
Private Sub Midi_VolumeChange(ByVal Channel As Integer, ByVal Value As Integer)

    chCtl1(Channel).Volume = Value
    If Value > 0 Then chCtl1(Channel).chActive = True Else chCtl1(Channel).chActive = False
    
End Sub
Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Form_MouseDown Button, Shift, Picture4.Left + X, Picture4.Top + Y

End Sub
Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    If (Y >= 16) And (Y <= 40) Then
        If Button <> 1 Then Exit Sub
        If CurrentControl <> 15 Then Exit Sub
        px = Y - X1
        px = IIf(px < 16, 16, px)
        px = IIf(px > 34, 34, px)
        Image5.Move 1, px
        List1.TopIndex = Int((px - 16) * 6.89)
        Exit Sub
    End If
    Form_MouseMove Button, Shift, Picture4.Left + X, Picture4.Top + Y

End Sub
Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Form_MouseUp Button, Shift, Picture4.Left + X, Picture4.Top + Y

End Sub
Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    Select Case Y
        Case 15 To 28
            If Button <> 1 Then Exit Sub
            If CurrentControl <> 11 Then Exit Sub
            px = X - X1
            px = IIf(px < 6, 6, px)
            px = IIf(px > 81, 81, px)
            UndValue = Int((px - 6) / 6.25) - 6
            px = Int((6.25 * UndValue) + 43.5)
            Image1.Move px, 15
            lblInfo(3) = UndValue
            Exit Sub
        Case 45 To 58
            If Button <> 1 Then Exit Sub
            If CurrentControl <> 12 Then Exit Sub
            px = X - X1
            px = IIf(px < 6, 6, px)
            px = IIf(px > 81, 81, px)
            UndValue = Int((px - 6) / 6.25) - 6
            px = Int((6.25 * UndValue) + 43.5)
            Image2.Move px, 45
            lblInfo(5) = UndValue
            Exit Sub
        Case 75 To 88
            If Button <> 1 Then Exit Sub
            If CurrentControl <> 13 Then Exit Sub
            px = X - X1
            px = IIf(px < 6, 6, px)
            px = IIf(px > 81, 81, px)
            Und = Int((px - 6) / 7.5)
            px = Int(Und * 7.5) + 6
            Image3.Move px, 76
            lblInfo(4) = Und
            midi.MasterVolume = Und
            Exit Sub
    End Select

End Sub
