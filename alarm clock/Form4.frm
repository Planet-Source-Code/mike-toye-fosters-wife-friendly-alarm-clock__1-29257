VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wake up settings"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Execute an application"
      Height          =   2235
      Left            =   120
      TabIndex        =   12
      Top             =   2820
      Width           =   5715
      Begin VB.Frame Frame4 
         Caption         =   "Application run focus"
         Height          =   975
         Left            =   180
         TabIndex        =   17
         Top             =   1020
         Width           =   2415
         Begin VB.OptionButton OptRUN 
            Caption         =   "Hidden"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton OptRUN 
            Caption         =   "Minimised"
            Height          =   195
            Index           =   1
            Left            =   1140
            TabIndex        =   20
            Top             =   300
            Width           =   1035
         End
         Begin VB.OptionButton OptRUN 
            Caption         =   "Normal"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   19
            Top             =   600
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptRUN 
            Caption         =   "Maximised"
            Height          =   195
            Index           =   3
            Left            =   1140
            TabIndex        =   18
            Top             =   600
            Width           =   1035
         End
      End
      Begin VB.CheckBox chkEXE 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         Height          =   195
         Left            =   4560
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   540
         Width           =   4995
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   285
         Left            =   5220
         TabIndex        =   13
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "File"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   1800
   End
   Begin VB.PictureBox picPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2220
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   3660
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Play WAV file"
      Height          =   1515
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   5715
      Begin VB.CheckBox chkWAV 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         Height          =   195
         Left            =   4560
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdPlayWAV 
         Enabled         =   0   'False
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   435
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   285
         Left            =   5220
         TabIndex        =   8
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   4995
      End
      Begin VB.Label Label1 
         Caption         =   "File"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pop up message"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5715
      Begin VB.CheckBox chkPOP 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Text            =   "Time's up!!"
         Top             =   300
         Width           =   5355
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6060
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6060
      TabIndex        =   0
      Top             =   180
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   5355
      Left            =   6900
      Top             =   0
      Width           =   315
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPlayWAV_Click()
    sndPlaySound Text2, 2
End Sub

Private Sub Command1_Click()
    Timer1.Enabled = False
    SaveSetting App.Title, "Settings", "EXEEnabled", IIf(chkEXE.Value = vbChecked, 1, 0)
    SaveSetting App.Title, "Settings", "POPEnabled", IIf(chkPOP.Value = vbChecked, 1, 0)
    SaveSetting App.Title, "Settings", "WAVEnabled", IIf(chkWAV.Value = vbChecked, 1, 0)
    SaveSetting App.Title, "Settings", "POPmsg", Text1
Dim x As Integer
    For x = 0 To 3
        If OptRUN(x).Value = True Then
            SaveSetting App.Title, "Settings", "EXEopt", x
            Exit For
        End If
    Next x
    Unload Me
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub Command3_Click()
    If Text2 > "" Then
        sIO_filename = Text2
    Else
        sIO_filename = "c:\"
    End If
    Form6.Show 1
    If sIO_filename > "" Then
        Text2 = sIO_filename
        SaveSetting App.Title, "Settings", "WAVLocation", sIO_filename
    End If
End Sub

Private Sub Command4_Click()
    If Text3 > "" Then
        sIO_filename = Text3
    Else
        sIO_filename = "c:\"
    End If
    Form6.Show 1
    If sIO_filename > "" Then
        Text3 = sIO_filename
        SaveSetting App.Title, "Settings", "EXELocation", sIO_filename
    End If
End Sub

Private Sub Form_Load()
On Local Error GoTo Bugger
    cmdPlayWAV.Picture = picPlay.Picture
    Timer1.Enabled = True
    Text2 = GetSetting(App.Title, "Settings", "WAVLocation")
    Text3 = GetSetting(App.Title, "Settings", "EXELocation")
    If GetSetting(App.Title, "Settings", "EXEEnabled") Then
        chkEXE.Value = vbChecked
    Else
        chkEXE.Value = vbUnchecked
    End If
    If GetSetting(App.Title, "Settings", "WAVEnabled") Then
        chkWAV.Value = vbChecked
    Else
        chkWAV.Value = vbUnchecked
    End If
    If GetSetting(App.Title, "Settings", "POPEnabled") Then
        chkPOP.Value = vbChecked
    Else
        chkPOP.Value = vbUnchecked
    End If
    Text1 = GetSetting(App.Title, "Settings", "POPmsg")
Dim x As Integer
    x = GetSetting(App.Title, "Settings", "EXEopt")
    OptRUN(x).Value = True
    
Bugger:
    
End Sub

Private Sub Timer1_Timer()
    If FileExists(Text2) Then
        cmdPlayWAV.Enabled = True
    Else
        cmdPlayWAV.Enabled = False
    End If
End Sub
