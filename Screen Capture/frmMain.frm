VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Capture"
   ClientHeight    =   1410
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3585
   ClipControls    =   0   'False
   DrawMode        =   6  'Mask Pen Not
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Type declarations
Private Type POINTAPI
  x As Long
  y As Long
End Type
    
Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
    
'Menu API constants for popup menu
Private Const MF_STRING As Long = &H0&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_BYCOMMAND As Long = &H0&
Private Const TPM_RETURNCMD As Long = &H100&

'Internal menu constants for popup menu
Private Const ID_CANCEL As Long = &H6000&
Private Const ID_SEPERATOR As Long = &H6001
Private Const ID_CAPIMAGE As Long = &H6002&
Private Const ID_CAPWEBCOLOR As Long = &H6003&
Private Const ID_EXIT As Long = &H6004&

'Popup menu variable
Private m_hPopup As Long

'Functions used for popup menu
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu&) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu&, ByVal wFlags&, ByVal wIDNewItem&, ByVal lpNewItem$) As Long
Private Declare Function ClientToScreen& Lib "user32" (ByVal hwnd&, lpPoint As POINTAPI)
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu&, ByVal wFlags&, ByVal x&, ByVal y&, ByVal nReserved&, ByVal hwnd&, ByVal lpRect&) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)

'System tray icon constants
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLICK = &H203
Private Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = 0&
Private Const NIM_DELETE = 2&
Private Const NIM_MODIFY = 1&
Private Const NIF_ICON = 2&
Private Const NIF_TIP = &H4
Private Const NIF_MESSAGE = 1&

'System tray variable
Private Notify As NOTIFYICONDATA

'Function used for system tray
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Sub Form_Load()
  'Add icon to system tray
  With Notify
    .cbSize = Len(Notify)
    .hwnd = Me.hwnd
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "Double click to capture image, right click for options!" & vbNullChar
  End With
  Dim lResult As Long
  lResult = Shell_NotifyIcon(NIM_ADD, Notify)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

  'Destroy popup menu if it was created
  If m_hPopup Then Call DestroyMenu(m_hPopup)
  
  'Remove system tray icon
  Dim lResult As Long
  lResult = Shell_NotifyIcon(NIM_DELETE, Notify)
  
  'Unload all forms
  Dim oForm As Form
  For Each oForm In Forms
    If oForm.hwnd <> Me.hwnd Then Unload oForm
  Next oForm
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Get system tray mouse message
  Static Message As Long
  Message = x / Screen.TwipsPerPixelX

  Select Case Message
    Case WM_RBUTTONUP 'Show popup menu
      If m_hPopup = 0 Then CreateMenu 'Create the popup menu if it doesn't exists yet
      
      Dim ptAPI As POINTAPI 'Get our mouse position
      GetCursorPos ptAPI
      Call ClientToScreen(Me.hwnd, ptAPI)
      
      'Select popup menu clicked
      Select Case TrackPopupMenu(m_hPopup, TPM_RETURNCMD, ptAPI.x, ptAPI.y, 0&, Me.hwnd, 0&)
        Case ID_CANCEL
          'Do Nothing
        Case ID_EXIT
          Unload frmMain
        Case ID_CAPWEBCOLOR
          Screen.MousePointer = vbHourglass
          frmGetColor.Show
          Screen.MousePointer = vbNormal
        Case ID_CAPIMAGE
          Screen.MousePointer = vbHourglass
          frmCapture.Show
          Screen.MousePointer = vbNormal
      End Select
      
    Case WM_LBUTTONDBLCLICK 'Capture image
      Screen.MousePointer = vbHourglass
      frmCapture.Show
      Screen.MousePointer = vbNormal
  End Select
End Sub

Private Sub CreateMenu()
  'Create popup menu
  m_hPopup = CreatePopupMenu()
  Call AppendMenu(m_hPopup, MF_STRING, ID_CANCEL, ByVal "Cancel")
  Call AppendMenu(m_hPopup, MF_STRING Or MF_SEPARATOR, ID_SEPERATOR, ByVal vbNullString)
  Call AppendMenu(m_hPopup, MF_STRING, ID_CAPIMAGE, ByVal "Capture Image")
  Call AppendMenu(m_hPopup, MF_STRING, ID_CAPWEBCOLOR, ByVal "Capture Web Color")
  Call AppendMenu(m_hPopup, MF_STRING Or MF_SEPARATOR, ID_SEPERATOR, ByVal vbNullString)
  Call AppendMenu(m_hPopup, MF_STRING, ID_EXIT, ByVal "Exit")
  'Bold the first menu item (cancel)
  Call SetMenuDefaultItem(m_hPopup, 0, 1&)
End Sub

