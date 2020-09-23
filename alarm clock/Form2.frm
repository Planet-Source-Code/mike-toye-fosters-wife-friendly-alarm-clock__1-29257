VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuSetAlarm 
         Caption         =   "Set alarm"
      End
      Begin VB.Menu mnuWakeUp 
         Caption         =   "Wake up settings"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlarmEnabled 
         Caption         =   "Alarm enabled"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAlarmEnabled_Click()
    If mnuAlarmEnabled.Checked = True Then
        mnuAlarmEnabled.Checked = False
        SaveSetting App.Title, "Settings", "AlarmSet", 0
    Else
        mnuAlarmEnabled.Checked = True
        SaveSetting App.Title, "Settings", "AlarmSet", 1
    End If
    
End Sub

Private Sub mnuExit_Click()
On Error GoTo EndAnyway
    Unload Form1
    Unload Form2
EndAnyway:
    End
End Sub

Private Sub mnuSetAlarm_Click()
    Form3.Show 1
End Sub

Private Sub mnuWakeUp_Click()
    Form4.Show 1
End Sub
