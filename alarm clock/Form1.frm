VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timMain 
      Interval        =   800
      Left            =   1740
      Top             =   1380
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1380
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   285
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   1080
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "User32" (ByVal _
    hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
    ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()


Sub SetTopmostWindow(ByVal hWnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hWnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
End Sub
Private Sub Form_Load()
On Local Error GoTo Error
    Me.Left = Screen.Width / 2
    Me.Top = -10
    picMain.Left = 0
    picMain.Top = 0
    Me.Width = picMain.Width
    Me.Height = picMain.Height
    App.Title = "Fosters' Alarm Clock"
    
    Me.Show
    SetTopmostWindow Me.hWnd
    AlarmTime = GetSetting(App.Title, "Settings", "AlarmTime")
    If GetSetting(App.Title, "Settings", "AlarmSet") = 0 Then
        Form2.mnuAlarmEnabled.Checked = False
    Else
        Form2.mnuAlarmEnabled.Checked = True
    End If
Error:
    SaveSetting App.Title, "Settings", "AlarmSet", 0
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Form2.mnuPop
    End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub timMain_Timer()
Dim dNow As Date
    dNow = Now

    If Form2.mnuAlarmEnabled.Checked = True And AlarmTime < dNow Then
        DoTheAlarmThing
        timMain.Enabled = False
        Form2.mnuAlarmEnabled.Checked = False
        SaveSetting App.Title, "Settings", "AlarmSet", 0
        
    End If
    If Form2.mnuAlarmEnabled.Checked = True Then
        picMain.ToolTipText = "Alarm set for " & Format(AlarmTime, "dd/mm/yyyy hh:mm")
    Else
        picMain.ToolTipText = "Alarm is not set"
    End If
End Sub
Sub DoTheAlarmThing()
On Local Error GoTo Failed
Dim rc As Integer

    If GetSetting(App.Title, "Settings", "EXEEnabled") = 1 Then
        rc = Shell(GetSetting(App.Title, "Settings", "EXELocation"), vbNormalFocus)
    End If
    If GetSetting(App.Title, "Settings", "WAVEnabled") = 1 Then
        sndPlaySound GetSetting(App.Title, "Settings", "WAVLocation"), 2
    End If
    If GetSetting(App.Title, "Settings", "POPEnabled") = 1 Then
        MsgBox GetSetting(App.Title, "Settings", "POPmsg"), vbExclamation, App.Title
    End If
Failed:
End Sub
