VERSION 5.00
Begin VB.Form frmGetColor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   DrawStyle       =   2  'Dot
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmGetColor.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPixel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   75
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   134
      TabIndex        =   0
      Top             =   105
      Width           =   2040
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Web Color: N/A"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   15
         TabIndex        =   1
         Top             =   375
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmGetColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strColor As String

Private Sub Form_Load()
  'Capture desktop and make it this forms background picture
  Dim DeskhWnd As Long, DeskDC As Long
  Me.WindowState = vbMaximized
  DeskhWnd& = GetDesktopWindow()
  DeskDC& = GetDC(DeskhWnd&)
  BitBlt Me.hDC, 0&, 0&, Screen.Width, Screen.Height, DeskDC&, 0&, 0&, SRCCOPY
  Me.Picture = Me.Image
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'User hit escape so unload
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Used press the mouse button, set clipboard to color selected
  Clipboard.Clear
  Clipboard.SetText strColor
  Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  With picPixel
    'Get left position to show picture box that follows the mouse
    If ScaleWidth - x < (150) Then
      .Left = x - 154
    Else
      .Left = x + 12
    End If
    
    'Get top position to show picture box that follows the mouse
    If ScaleHeight - y < (50) Then
      .Top = y - 54
    Else
      .Top = y + 12
    End If
  End With
    
  Dim lColor As Long
  lColor = Me.Point(x, y) 'Get pixel color under mouse
  picPixel.BackColor = lColor 'Set out pictuebox to the same color
  strColor = MakeHexLong(lColor)  'Get our hex value from the color
  lblColor = " Web Color: #" & strColor 'Set our lable to show current color
End Sub
