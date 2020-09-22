VERSION 5.00
Begin VB.UserControl VuBar 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Enabled         =   0   'False
   HitBehavior     =   0  'None
   PaletteMode     =   4  'None
   Picture         =   "VuBar.ctx":0000
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      Height          =   1875
      Left            =   4185
      Picture         =   "VuBar.ctx":98D8
      ScaleHeight     =   1875
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4410
      Top             =   2655
   End
End
Attribute VB_Name = "VuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim Pic As StdPicture
Const m_def_Value = 0
Dim m_Value(0 To 15) As Integer
'MemberInfo=7,0,2,0
Public Property Get Value(ByVal Channel As Integer) As Integer
Attribute Value.VB_MemberFlags = "400"
    
    Value = m_Value(Channel)

End Property
Public Property Let Value(ByVal Channel As Integer, ByVal New_Value As Integer)
    
    On Error Resume Next
    m_Value(Channel) = New_Value
    m_Value(Channel) = 7 * Int(m_Value(Channel) / 7)
    If Not Ambient.UserMode Then
        Timer1.Enabled = False
    Else
        Timer1.Enabled = True
    End If
    SetVu Channel, m_Value(Channel)
    
End Property
Private Sub Timer1_Timer()

    For i = 0 To 15
        X = Value(i) - 5
        If X < 1 Then X = 1
        Value(i) = X
        DoEvents
    Next

End Sub
Private Sub UserControl_InitProperties()
    
    For i = 0 To 15
        m_Value(i) = m_def_Value
    Next
    
End Sub
Private Sub SetVu(ByVal Channel As Integer, ByVal nValue As Integer)
    
    On Error Resume Next
    l = ((Channel) * 15) + 7
    BitBlt UserControl.hDC, l, (127 - nValue) + 5, 12, nValue, Picture1.hDC, 0, 127 - nValue, vbSrcCopy
    BitBlt UserControl.hDC, l, 5, 12, 127 - nValue, Picture1.hDC, 12, 0, vbSrcCopy
    DoEvents

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    For i = 0 To 15
        m_Value(i) = 0
    Next
    
End Sub
