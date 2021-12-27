VERSION 5.00
Begin VB.Form MonitorBox 
   Caption         =   "Image Monitor"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "Monitor.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   13.313
   ScaleMode       =   4  'Character
   ScaleWidth      =   39
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu MainMenu 
      Caption         =   "&Main"
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
      End
      Begin VB.Menu OpenDirectoryMenu 
         Caption         =   "&Open Directory"
      End
      Begin VB.Menu CloseMonitorMenu 
         Caption         =   "&Close Monitor"
      End
   End
End
Attribute VB_Name = "MonitorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit

'The Microsoft Windows API constants, functions, and structures used by this program:
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeoutVersion As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Private Const NIF_ICON As Long = &H2&
Private Const NIF_INFO As Long = &H10&
Private Const NIF_MESSAGE As Long = &H1&
Private Const NIF_TIP As Long = &H4&
Private Const NIIF_INFO As Long = &H1&
Private Const NIM_ADD As Long = &H0&
Private Const NIM_DELETE As Long = &H2&
Private Const NIM_MODIFY As Long = &H1&
Private Const NIM_SETFOCUS As Long = &H3&
Private Const NIM_SETVERSION  As Long = &H4&
Private Const NOTIFYICON_VERSION As Long = &H3&
Private Const WM_MOUSEMOVE As Long = &H200&

Private Declare Function SetForegroundWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "Shell32.dll" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long

'The variables used by this program:
Private Destination As String           'The directore where the captured images are saved.
Private StopMonitor As Boolean          'Indicates whether the image monitor should stop.
Private TrayIconData As NOTIFYICONDATA  'The system tray icon information.


'This procedure closes this window.
Private Sub CloseMonitorMenu_Click()
On Error Resume Next
Shell_NotifyIconA NIM_SETFOCUS, TrayIconData
Unload Me
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
ChDrive Left$(App.Path, InStr(App.Path, ":"))
ChDir App.Path

Me.Width = Screen.Width / 2
Me.Height = Screen.Height / 2
Me.Visible = False
  
If App.PrevInstance Then
   MsgBox App.Title & " is already running.", vbInformation
   Unload Me
End If
  
Destination = Empty
StopMonitor = False

With TrayIconData
   .cbSize = Len(TrayIconData)
   .dwInfoFlags = NIIF_INFO
   .dwState = 0
   .dwStateMask = 0
   .hIcon = Me.Icon
   .hwnd = Me.hwnd
   .szInfo = App.Title & " has started." & vbNullChar
   .szInfoTitle = App.Title & vbNullChar
   .szTip = App.Title & vbNullChar
   .uCallbackMessage = WM_MOUSEMOVE
   .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
   .uID = vbNull
   .uTimeoutVersion = 3000
End With

Shell_NotifyIconA NIM_ADD, TrayIconData
TrayIconData.uTimeoutVersion = NOTIFYICON_VERSION
Shell_NotifyIconA NIM_SETVERSION, TrayIconData

StartMonitor
EndRoutine:
Exit Sub

ErrorTrap:
HandleError
Resume EndRoutine
End Sub

'This procedure displays the system tray icon menu.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
If Button = vbRightButton Then
   SetForegroundWindow Me.hwnd
   PopupMenu MainMenu
End If
EndRoutine:
Exit Sub

ErrorTrap:
HandleError
Resume EndRoutine
End Sub

'This procedure closes this program.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
EndRoutine:
StopMonitor = True
Shell_NotifyIconA NIM_DELETE, TrayIconData
End
Exit Sub

ErrorTrap:
HandleError
Resume EndRoutine
End Sub

'This procedure handles any errors that occur.
Private Sub HandleError()
Dim ErrorCode As Long
Dim Description As String

Description = Err.Description
ErrorCode = Err.Number
Err.Clear

On Error Resume Next
MsgBox "Error code: " & CStr(ErrorCode) & vbCr & Description & ".", vbExclamation
End Sub


'This procedur displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
With App
   MsgBox .Comments, vbInformation, .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
End With

Shell_NotifyIconA NIM_SETFOCUS, TrayIconData
EndRoutine:
Exit Sub

ErrorTrap:
HandleError
Resume EndRoutine
End Sub


'This procedure opens a Microsoft Windows Explorer window displaying the image destination directory.
Private Sub OpenDirectoryMenu_Click()
On Error GoTo ErrorTrap
Shell "Explorer.exe " & Destination, vbNormalFocus
Shell_NotifyIconA NIM_SETFOCUS, TrayIconData
EndRoutine:
Exit Sub

ErrorTrap:
HandleError
Resume EndRoutine
End Sub

'This procedure requests the user to specify a destination directory and starts monitoring the clipboard.
Private Sub StartMonitor()
On Error GoTo ErrorTrap
Dim FileName As String
Dim Index As Long
 
Destination = InputBox$("Specify a directory to save any captured images in.", , CurDir$())
If Destination = Empty Then Unload Me
If Not Right$(Destination, 1) = "\" Then Destination = Destination & "\"

FileName = Dir$(Destination & "*.bmp", vbNormal Or vbHidden Or vbSystem)
Index = -1
Do Until FileName = Empty
   FileName = Left$(FileName, Len(FileName) - 4)
   If CStr(CLng(Val(FileName))) = FileName Then If Val(FileName) > Index Then Index = Val(FileName)
   FileName = Dir$()
Loop

Do
   Clipboard.Clear
   Index = Index + 1
   Do Until (Clipboard.GetFormat(vbCFBitmap) Or StopMonitor)
      DoEvents
   Loop
   
   If StopMonitor Then Exit Do
   
   SavePicture Clipboard.GetData(vbCFBitmap), Destination & Trim$(CStr(Index)) + ".bmp"
   TrayIconData.szInfo = "Saved: """ & Destination & Trim$(CStr(Index)) + ".bmp""" & vbNullChar
   Shell_NotifyIconA NIM_MODIFY, TrayIconData
Loop
EndRoutine:
Exit Sub

ErrorTrap:
HandleError
Resume EndRoutine
End Sub


