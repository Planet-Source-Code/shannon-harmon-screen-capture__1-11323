VERSION 5.00
Begin VB.Form frmCapture 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   DrawMode        =   6  'Mask Pen Not
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmCapture.frx":0000
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   90
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   108
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnCapturing As Boolean
Dim X1 As Single
Dim Y1 As Single

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
  'User pressed escape so unload
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not blnCapturing Then  'Start capture
    MousePointer = 99 'Change our mousepointer to custom
    X1 = x: Y1 = y  'Set our starting x & y
    blnCapturing = True 'Turn capturing bit on
  ElseIf blnCapturing = True Then 'Done capturing
    If Button = vbRightButton Then  'User clicked right mouse so cancel but stay capturing
      blnCapturing = False  'Turn capturing bit off
      MousePointer = vbNormal 'Set our mousepointer back to normal
      Cls 'Clear anything we drew to the form
    ElseIf Button = vbLeftButton Then 'User clicked left mouse button so capture
      CaptureIt X1, x, Y1, y  'Do the capture
      Unload Me 'Unload form
    End If
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If blnCapturing Then  'If we are capturing then draw box and dimensions
    Cls 'Clear the form
    Line (X1, Y1)-(x, y), , B 'Draw our box where the mouse selection is
    
    'Get left, right, top and bottom regarldess of where they started and ended
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim lWidth As Long, lHeight As Long
    Left = IIf(X1 > x, x, X1)
    Right = IIf(X1 < x, x, X1)
    Top = IIf(Y1 > y, y, Y1)
    Bottom = IIf(Y1 < y, y, Y1)
    lWidth = (Right - Left)
    lHeight = (Bottom - Top)

    Dim strOut As String
    strOut = lWidth & "x" & lHeight 'Setup our dimensions string

    'If the text will fit in our selection then draw to screen
    If lWidth > TextWidth(strOut) And lHeight > TextHeight(strOut) Then

      Dim tX As Single, tY As Single
      Dim cX As Single, cY As Single
      cX = Right - (lWidth / 2) 'Get our center of our rectangle's x position
      cY = Bottom - (lHeight / 2) 'Get our center of our rectangle's y position
      tX = cX - TextWidth(strOut) / 2 'Get our offset from x center with text width
      tY = cY - TextHeight(strOut) / 2  'Get our offset from y center with text height

      If Me.Point(cX, cY) < 62255 / 2 Then  'Center of selection color was darker
        ForeColor = vbWhite 'Set font color to white
      Else  'Color was lighter
        ForeColor = vbBlack 'Set font color to black
      End If

      TextOut Me.hDC, tX, tY, strOut, Len(strOut) 'Draw our dimensions text on the form

    End If
  End If
End Sub

Private Sub CaptureIt(xStart As Single, xEnd As Single, yStart As Single, yEnd As Single)
  Dim Left As Long, Top As Long, Right As Long, Bottom As Long
  Dim lWidth As Long, lHeight As Long

  blnCapturing = False

  'Get left, right, top and bottom regarldess of where they started and ended
  Left = IIf(xStart > xEnd, xEnd, xStart)
  Right = IIf(xStart < xEnd, xEnd, xStart)
  Top = IIf(yStart > yEnd, yEnd, yStart)
  Bottom = IIf(yStart < yEnd, yEnd, yStart)
  lWidth = (Right - Left)
  lHeight = (Bottom - Top)
  
  If lWidth <= 0 Or lHeight <= 0 Then GoTo PROC_TOOSMALL  'Nothing to capture
  
  With picTemp
    .Cls  'Clear our picture box that holds the image till copied to clipboar
    .Width = lWidth 'Set it's hight and width
    .Height = lHeight
  End With
  
  Me.Cls  'Clear screen so we don't get the box and dimensions
  BitBlt picTemp.hDC, 0, 0, lWidth, lHeight, Me.hDC, Left, Top, SRCCOPY 'Copy screen to picture box
  
  Clipboard.Clear 'Clear clipboard
  Clipboard.SetData picTemp.Image 'Copy image to clipboard
  
PROC_EXIT:
  Exit Sub
  
PROC_TOOSMALL:
  GoTo PROC_EXIT
End Sub
