VERSION 5.00
Begin VB.Form frmDevices 
   BackColor       =   &H00788C9E&
   BorderStyle     =   0  'None
   ClientHeight    =   2460
   ClientLeft      =   3615
   ClientTop       =   1605
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDevices.frx":0000
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1200
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   3435
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   2655
      Picture         =   "frmDevices.frx":A5B0
      Top             =   1965
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   1665
      Picture         =   "frmDevices.frx":B0B0
      Top             =   1965
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Devices"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   3570
   End
End
Attribute VB_Name = "frmDevices"
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
Public Sub CenterForm(ByVal Frm As Form)
        
    Frm.Top = (Screen.Height / 2) - (Frm.Height / 2)
    Frm.Left = (Screen.Width / 2) - (Frm.Width / 2)

End Sub
Private Sub Form_Load()

    CenterForm Me

End Sub
Private Sub Image1_Click()
    
    SaveSetting "midiPlayer", "Devices", "DeviceIndex", List1.ListIndex
    Me.Hide

End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    Image1.Picture = LoadResPicture(126, 0)

End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Image1.Picture = LoadResPicture(125, 0)

End Sub
Private Sub Image2_Click()
    
    Me.Hide

End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image2.Picture = LoadResPicture(128, 0)

End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image2.Picture = LoadResPicture(127, 0)

End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Drag hWnd, Button

End Sub

Private Sub List1_DblClick()

    Image1_Click

End Sub


