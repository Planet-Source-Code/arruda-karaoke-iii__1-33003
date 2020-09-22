VERSION 5.00
Begin VB.UserControl chCtl 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   CanGetFocus     =   0   'False
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   765
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   Picture         =   "ctlChannel.ctx":0000
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   51
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   60
      Picture         =   "ctlChannel.ctx":419C
      Top             =   495
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Bt6 
      Height          =   195
      Left            =   90
      Picture         =   "ctlChannel.ctx":46D4
      Top             =   1035
      Width           =   150
   End
   Begin VB.Image Bt5 
      Height          =   195
      Left            =   90
      Picture         =   "ctlChannel.ctx":4890
      Top             =   1500
      Width           =   150
   End
   Begin VB.Image Bt4 
      Height          =   195
      Left            =   90
      Picture         =   "ctlChannel.ctx":4A4C
      Top             =   1935
      Width           =   150
   End
   Begin VB.Image Bt3 
      Height          =   195
      Left            =   90
      Picture         =   "ctlChannel.ctx":4C08
      Top             =   2370
      Width           =   150
   End
   Begin VB.Image Bt2 
      Height          =   195
      Left            =   285
      Picture         =   "ctlChannel.ctx":4DC4
      Top             =   2850
      Width           =   150
   End
   Begin VB.Image Bt1 
      Height          =   165
      Left            =   195
      Picture         =   "ctlChannel.ctx":4F94
      Top             =   4185
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   75
      Top             =   330
      Width           =   195
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   90
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   540
      TabIndex        =   0
      Top             =   135
      Width           =   75
   End
End
Attribute VB_Name = "chCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim y1 As Integer, X1 As Integer
Dim CurrentControl As Integer
Const m_def_Channel = 1
Const m_def_NoteEvent = ""
Const m_def_Led = 0
Const m_def_Volume = 0
Const m_def_Balance = 64
Const m_def_Phaser = 0
Const m_def_Celeste = 0
Const m_def_Chorus = 0
Const m_def_Tremulo = 0
Const m_def_Mute = 1
Dim m_Channel As Integer
Dim m_NoteEvent As String
Dim m_Led As Boolean
Dim m_Volume As Integer
Dim m_Balance As Integer
Dim m_Phaser As Integer
Dim m_Celeste As Integer
Dim m_Chorus As Integer
Dim m_Tremulo As Integer
Dim m_Mute As Boolean
Dim m_chActive As Boolean
Event VolumeChange(ByVal Value As Integer)
Event BalanceChange(ByVal Value As Integer)
Event ChorusChange(ByVal Value As Integer)
Event CelesteChange(ByVal Value As Integer)
Event PhaserChange(ByVal Value As Integer)
Event TremuloChange(ByVal Value As Integer)
Event MuteChange(ByVal Value As Boolean)
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    
    Enabled = UserControl.Enabled

End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"

End Property
Private Function GetValue(ByVal t1 As Integer, ByVal t2 As Integer, ByVal t3 As Integer) As Integer

    GetValue = (127 / (t2 - t1)) * (t3 - t1)

End Function
Private Function SetValue(ByVal t1 As Integer, ByVal t2 As Integer, ByVal t3 As Integer)

    SetValue = (((t2 - t1) / 127) * t3) + t1

End Function
'MemberInfo=7,0,0,0
Public Property Get Volume() As Integer
    Volume = m_Volume
End Property
Public Property Let Volume(ByVal New_Volume As Integer)
    
    m_Volume = New_Volume
    PropertyChanged "Volume"
    Bt1.Top = 212 + (279 - (SetValue(212, 279, m_Volume)))
    
End Property
'MemberInfo=7,0,0,64
Public Property Get Balance() As Integer
    
    Balance = m_Balance
    
End Property
Public Property Let Balance(ByVal New_Balance As Integer)
    
    m_Balance = New_Balance
    PropertyChanged "Balance"
    Bt2.Left = SetValue(6, 33, m_Balance)
    
End Property
'MemberInfo=7,0,0,0
Public Property Get Phaser() As Integer
    
    Phaser = m_Phaser

End Property
Public Property Let Phaser(ByVal New_Phaser As Integer)
    
    m_Phaser = New_Phaser
    PropertyChanged "Phaser"
    Bt3.Left = SetValue(6, 33, m_Phaser)

End Property
'MemberInfo=7,0,0,0
Public Property Get Celeste() As Integer
    
    Celeste = m_Celeste

End Property
Public Property Let Celeste(ByVal New_Celeste As Integer)
    
    m_Celeste = New_Celeste
    PropertyChanged "Celeste"
    Bt4.Left = SetValue(6, 33, m_Celeste)

End Property
'MemberInfo=7,0,0,0
Public Property Get Chorus() As Integer
    
    Chorus = m_Chorus

End Property
Public Property Let Chorus(ByVal New_Chorus As Integer)
    
    m_Chorus = New_Chorus
    PropertyChanged "Chorus"
    Bt5.Left = SetValue(6, 33, m_Chorus)

End Property
'MemberInfo=7,0,0,0
Public Property Get Tremulo() As Integer
    
    Tremulo = m_Tremulo

End Property
Public Property Let Tremulo(ByVal New_Tremulo As Integer)
    
    m_Tremulo = New_Tremulo
    PropertyChanged "Tremulo"
    Bt6.Left = SetValue(6, 33, m_Tremulo)

End Property
'MemberInfo=0,0,0,1
Public Property Get Mute() As Boolean
    
    Mute = m_Mute

End Property
Public Property Let Mute(ByVal New_Mute As Boolean)
    
    m_Mute = New_Mute
    PropertyChanged "Mute"
    If m_Mute Then Image1.Visible = True Else Image1.Visible = False

End Property
Private Sub Bt1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    y1 = ScaleY(Y, 1, 3)
    CurrentControl = 1
    
End Sub
Private Sub Bt1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        UserControl_MouseMove Button, Shift, ScaleX(X, 1, 3) + 13, Bt1.Top + ScaleY(Y, 1, 3)
    End If

End Sub
Private Sub Bt2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    X1 = ScaleX(X, 1, 3)
    CurrentControl = 2

End Sub
Private Sub Bt2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        UserControl_MouseMove Button, Shift, Bt2.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 190
    End If

End Sub
Private Sub Bt3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    X1 = ScaleX(X, 1, 3)
    CurrentControl = 3

End Sub
Private Sub Bt3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        UserControl_MouseMove Button, Shift, Bt3.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 158
    End If

End Sub
Private Sub Bt4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    X1 = ScaleX(X, 1, 3)
    CurrentControl = 4

End Sub
Private Sub Bt4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        UserControl_MouseMove Button, Shift, Bt4.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 129
    End If

End Sub
Private Sub Bt5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    X1 = ScaleX(X, 1, 3)
    CurrentControl = 5

End Sub
Private Sub Bt5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        UserControl_MouseMove Button, Shift, Bt5.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 100
    End If

End Sub
Private Sub Bt6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    X1 = ScaleX(X, 1, 3)
    CurrentControl = 6

End Sub
Private Sub Bt6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        UserControl_MouseMove Button, Shift, Bt6.Left + ScaleX(X, 1, 3), ScaleY(Y, 1, 3) + 69
    End If

End Sub
Private Sub UserControl_InitProperties()
    
    m_Volume = m_def_Volume
    m_Balance = m_def_Balance
    m_Phaser = m_def_Phaser
    m_Celeste = m_def_Celeste
    m_Chorus = m_def_Chorus
    m_Tremulo = m_def_Tremulo
    m_Mute = m_def_Mute
    m_Channel = m_def_Channel
    m_NoteEvent = m_def_NoteEvent
    m_Led = m_def_Led

End Sub
'MemberInfo=7,0,0,1
Public Property Get Channel() As Integer
    
    Channel = m_Channel

End Property
Public Property Let Channel(ByVal New_Channel As Integer)
    
    m_Channel = New_Channel
    PropertyChanged "Channel"
    Label1 = m_Channel

End Property
'MemberInfo=13,0,0,
Public Property Get NoteEvent() As String
    
    NoteEvent = m_NoteEvent

End Property
Public Property Let NoteEvent(ByVal New_NoteEvent As String)
    
    m_NoteEvent = New_NoteEvent
    PropertyChanged "NoteEvent"
    Label4 = m_NoteEvent

End Property
'MemberInfo=0,0,0,0
Public Property Get Led() As Boolean
    
    Led = m_Led

End Property
Public Property Get chActive() As Boolean
    
    chActive = m_chActive

End Property
Public Property Let Led(ByVal New_Led As Boolean)
    
    m_Led = New_Led
    PropertyChanged "Led"
    If m_Led Then
        Shape1.FillColor = &HFF&
    Else
        Shape1.FillColor = &H80&
    End If
    
End Property
Public Property Let chActive(ByVal New_chActive As Boolean)
    
    m_chActive = New_chActive
    PropertyChanged "chActive"
    If m_chActive Then
        Label1.ForeColor = &HFF00&
    Else
        Label1.ForeColor = &H8000&
    End If
    
End Property
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (Y > 32) And (Y < 51) Then
        If (X > 3) And (X < 59) Then
            If Image1.Visible Then
                Image1.Visible = False
                m_Mute = False
            Else
                Image1.Visible = True
                m_Mute = True
            End If
            RaiseEvent MuteChange(m_Mute)
        End If
    End If

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Select Case Y
            Case 210 To 289
                If CurrentControl <> 1 Then Exit Sub
                If (X > 11) And (X < 33) Then
                    py = Y - y1
                    py = IIf(py < 212, 212, py)
                    py = IIf(py > 279, 279, py)
                    Bt1.Move 13, py
                    m_Volume = 127 - GetValue(212, 279, py)
                    RaiseEvent VolumeChange(m_Volume)
                End If
            Case 190 To 203
                If CurrentControl <> 2 Then Exit Sub
                If (X > 5) And (X < 43) Then
                    px = X - X1
                    px = IIf(px < 6, 6, px)
                    px = IIf(px > 33, 33, px)
                    Bt2.Move px, 190
                    m_Balance = GetValue(6, 33, px)
                    RaiseEvent BalanceChange(m_Balance)
                End If
            Case 158 To 171
                If CurrentControl <> 3 Then Exit Sub
                If (X > 5) And (X < 43) Then
                    px = X - X1
                    px = IIf(px < 6, 6, px)
                    px = IIf(px > 33, 33, px)
                    Bt3.Move px, 158
                    m_Phaser = GetValue(6, 33, px)
                    RaiseEvent PhaserChange(m_Phaser)
                End If
            Case 129 To 142
                If CurrentControl <> 4 Then Exit Sub
                If (X > 5) And (X < 43) Then
                    px = X - X1
                    px = IIf(px < 6, 6, px)
                    px = IIf(px > 33, 33, px)
                    Bt4.Move px, 129
                    m_Celeste = GetValue(6, 33, px)
                    RaiseEvent CelesteChange(m_Celeste)
                End If
            Case 100 To 113
                If CurrentControl <> 5 Then Exit Sub
                If (X > 5) And (X < 43) Then
                    px = X - X1
                    px = IIf(px < 6, 6, px)
                    px = IIf(px > 33, 33, px)
                    Bt5.Move px, 100
                    m_Chorus = GetValue(6, 33, px)
                    RaiseEvent ChorusChange(m_Chorus)
                End If
            Case 69 To 82
                If CurrentControl <> 6 Then Exit Sub
                If (X > 5) And (X < 43) Then
                    px = X - X1
                    px = IIf(px < 6, 6, px)
                    px = IIf(px > 33, 33, px)
                    Bt6.Move px, 69
                    m_Tremulo = GetValue(6, 33, px)
                    RaiseEvent TremuloChange(m_Tremulo)
                End If
        End Select
    End If

End Sub
