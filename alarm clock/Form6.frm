VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open File"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   180
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   8
      Top             =   180
      Width           =   510
   End
   Begin VB.PictureBox picOpen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "Form6.frx":0E12
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   240
      Picture         =   "Form6.frx":1C24
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4140
      TabIndex        =   5
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   5460
      TabIndex        =   4
      Top             =   3060
      Width           =   1215
   End
   Begin VB.FileListBox FIL 
      Height          =   2235
      Left            =   4140
      TabIndex        =   3
      Top             =   180
      Width           =   2535
   End
   Begin VB.DirListBox DIR 
      Height          =   1890
      Left            =   960
      TabIndex        =   2
      Top             =   180
      Width           =   3135
   End
   Begin VB.DriveListBox DRV 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   2100
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   3120
      Width           =   6975
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   5655
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    sIO_filename = FIL.Path & "\" & FIL.FileName
    sIO_filename = Replace(sIO_filename, "\\", "\")
    
    Unload Me
End Sub

Private Sub Command2_Click()
    sIO_filename = ""
    Unload Me
End Sub



Private Sub DIR_Change()
    FIL.Path = DIR.Path
    SetlblPath
End Sub

Private Sub DRV_change()
    DIR.Path = DRV.Drive
    SetlblPath
End Sub
Sub SetlblPath()
    lblPath = GetShortenedFileName(FIL.Path & IIf(FIL.FileName > "", "\" & FIL.FileName, ""), 70)
    lblPath = UCase(Left(lblPath, 1)) & Mid(lblPath, 2)
End Sub
Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

Function GetFilePath(FileName As String) As String
    Dim i As Long
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case ":"
                ' colons are always included in the result
                GetFilePath = Left$(FileName, i)
                Exit For
            Case "\"
                ' backslash aren't included in the result
                GetFilePath = Left$(FileName, i - 1)
                Exit For
        End Select
    Next
End Function

Function GetShortenedFileName(ByVal strFilePath As String, _
    ByVal maxLength As Long) As String
    Dim astrTemp() As String
    Dim lngCount As Long
    Dim strTemp As String
    Dim index As Long
    
    ' if the path is shorter than the max allowed length, just return it
    If Len(strFilePath) <= maxLength Then
        GetShortenedFileName = strFilePath
    Else
        ' split the path in its constituent dirs
        astrTemp() = Split(strFilePath, "\")
        lngCount = UBound(astrTemp)
        
        ' lets replace each part with ellipsis, until the length is OK
        ' but never substitute drive and file name
        For index = 1 To lngCount - 1
            astrTemp(index) = "..."
            ' rebuild the result
            GetShortenedFileName = Join(astrTemp, "\")
            If Len(GetShortenedFileName) <= maxLength Then Exit For
        Next
    End If
    
End Function



Private Sub FIL_Click()
    SetlblPath
End Sub


Private Sub Form_Load()
    If FileExists(sIO_filename) Then
        sIO_filename = GetFilePath(sIO_filename)
    End If
    DRV.Drive = sIO_filename
    DIR.Path = sIO_filename
    FIL.Path = sIO_filename
    
    
    SetlblPath
End Sub

