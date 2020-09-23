VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm Clock"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "1 day"
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton Command8 
      Caption         =   "15 minutes"
      Height          =   315
      Left            =   2520
      TabIndex        =   11
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton Command7 
      Caption         =   "1 hour"
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton Command6 
      Caption         =   "30 minutes"
      Height          =   315
      Left            =   600
      TabIndex        =   9
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton Command5 
      Caption         =   "10 minutes"
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton Command4 
      Caption         =   "5 minutes"
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3660
      TabIndex        =   5
      Top             =   540
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set alarm"
      Height          =   315
      Left            =   3660
      TabIndex        =   4
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "99:99"
      Top             =   540
      Width           =   555
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "99/99/9999"
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Set alarm for now plus..."
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2835
      Left            =   4320
      Top             =   0
      Width           =   315
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   300
      TabIndex        =   6
      Top             =   960
      Width           =   2715
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Time (HH:MM)"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date (YYYY/MM/DD)"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Not IsDate(txtDate) Or Not IsDate(txtTime) Then
        Label3 = "You cannot set a date/time, there is an error"
        If Not IsDate(txtDate) Then
            txtDate.SetFocus
        Else
            txtTime.SetFocus
        End If
        Exit Sub
    End If
    AlarmTime = txtDate & " " & txtTime
    SaveSetting App.Title, "Settings", "AlarmTime", AlarmTime
    Form2.mnuAlarmEnabled.Checked = True
    SaveSetting App.Title, "Settings", "AlarmSet", 1
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    AlarmTime = DateAdd("d", 1, Now)
    SetTXTDateTime
End Sub
Sub SetTXTDateTime()
        txtDate = Format(AlarmTime, "yyyy/mm/dd")
        txtTime = Format(AlarmTime, "hh:mm")
End Sub

Private Sub Command4_Click()
    AlarmTime = DateAdd("n", 5, Now)
    SetTXTDateTime
End Sub

Private Sub Command5_Click()
    AlarmTime = DateAdd("n", 10, Now)
    SetTXTDateTime
End Sub

Private Sub Command6_Click()
    AlarmTime = DateAdd("n", 30, Now)
    SetTXTDateTime
End Sub

Private Sub Command7_Click()
    AlarmTime = DateAdd("h", 1, Now)
    SetTXTDateTime
End Sub

Private Sub Command8_Click()
    AlarmTime = DateAdd("n", 15, Now)
    SetTXTDateTime
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title
    
    If AlarmTime > 0 Then
        SetTXTDateTime
    Else
        txtDate = Format(Now, "yyyy//mm/dd")
        txtTime = Format(Now, "hh:mm")
    End If
    
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0: txtDate.SelLength = Len(txtDate)
End Sub

Private Sub txtDate_LostFocus()
If Not IsDate(txtDate) Then
    Label3 = "Date is not a valid format"
Else
    Label3 = ""
End If
End Sub

Private Sub txtTime_GotFocus()
    txtTime.SelStart = 0: txtTime.SelLength = Len(txtTime)
End Sub

Private Sub txtTime_LostFocus()
If Not IsDate(txtTime) Then
    Label3 = "Time is not a valid format"
Else
    Label3 = ""
End If
End Sub
