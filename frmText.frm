VERSION 5.00
Begin VB.Form frmText 
   BackColor       =   &H00758799&
   BorderStyle     =   0  'None
   ClientHeight    =   5175
   ClientLeft      =   2265
   ClientTop       =   1830
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmText.frx":0000
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   675
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   5535
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   795
      Left            =   2385
      Picture         =   "frmText.frx":26000
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   5715
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00758799&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   6555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Teste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   6540
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   5940
      Picture         =   "frmText.frx":26B14
      Top             =   4380
      Width           =   750
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Public Sub Drag(ByVal hObj As Long, ByVal Button As Integer)
    
    Dim Ret As Long
    If Button = 1 Then
        ReleaseCapture
        Ret = SendMessage(hObj, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
    
End Sub
Private Sub Form_Load()

    CenterForm Me

End Sub
Public Sub CenterForm(ByVal Frm As Form)
        
    Frm.Top = (Screen.Height / 2) - (Frm.Height / 2)
    Frm.Left = (Screen.Width / 2) - (Frm.Width / 2)

End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Picture1.Point(ScaleX(X, 1, 3), ScaleY(Y, 1, 3)) = 0 Then
        Image1.Picture = LoadResPicture(124, 0)
    End If

End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Image1.Picture = LoadResPicture(123, 0)
    If Picture1.Point(ScaleX(X, 1, 3), ScaleY(Y, 1, 3)) = 0 Then Unload Me

End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Drag hWnd, Button

End Sub
