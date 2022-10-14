VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M3U-Maker-Extended"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13455
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "FrmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   678
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   897
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrSrcDir 
      Interval        =   1000
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton CmdSelAdd 
      Caption         =   "Hinzuf�gen"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ListBox LstFilesSrc 
      DragIcon        =   "FrmMain.frx":0442
      Height          =   4155
      Left            =   6360
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   120
      Width           =   6975
   End
   Begin VB.CheckBox ChkIncSubdir 
      Caption         =   "Unterverzeichnisse einbinden"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CdgMain 
      Left            =   5400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSelInv 
      Caption         =   "Markierung umkehren"
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton CmdSelNone 
      Caption         =   "Markierung aufheben"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton CmdSelAll 
      Caption         =   "Alles markieren"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox TxtPattern 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "*.mp3;*.wav;*.wma;*.wmv"
      ToolTipText     =   "Patterns separated by "";""."
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CheckBox ChkApplyPattern 
      Caption         =   "Pattern anwenden"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.OptionButton OptPathRel 
      Caption         =   "Relativ"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   4680
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton OptPathAbs 
      Caption         =   "Absolut"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   4680
      Width           =   855
   End
   Begin VB.DirListBox DirSrc 
      Height          =   2790
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6135
   End
   Begin VB.DriveListBox DlbSrc 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton CmdM3uSave 
      Caption         =   "Liste speichern"
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton CmdSelRem 
      Caption         =   "Markierte Elemente l�schen"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton CmdM3uOpen 
      Caption         =   "M3U-Liste �ffnen"
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.ListBox LstFilesDst 
      Height          =   4935
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   5040
      Width           =   13215
   End
   Begin VB.Label LblPathRelAbs 
      Caption         =   "Pfade:"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   4680
      Width           =   495
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetLogicalDrives Lib "kernel32" () As Long

Dim PathAbs As Boolean
Dim SrcDriveOri As String
Dim LogicalDrives As Long
    
Private Sub ChkApplyPattern_Click()
    ListFiles
End Sub

Private Sub ChkIncSubdir_Click()
    ListFiles
End Sub

Private Sub CmdSelAdd_Click()
    Dim i As Integer
    
    ' Append all selected files to output list
    For i = 0 To LstFilesSrc.ListCount - 1
        If LstFilesSrc.Selected(i) = True Then
            LstFilesDst.AddItem (LstFilesSrc.List(i))
        End If
    Next i
End Sub

Private Sub CmdM3uOpen_Click()
    Dim Filename As String
    
    ' Try to open file
    On Error GoTo ShowOpenError
    With CdgMain
        .Filter = "M3U Files (*.m3u)|*.m3u|All Files (*.*)|*.*"
        .CancelError = True
        .ShowOpen
        Filename = .Filename
    End With
    On Error GoTo 0
    
    If FileExists(Filename) Then
        Dim Line As String
        Dim FileNum As Integer
        
        FileNum = FreeFile
        LstFilesDst.Clear
        
        Open Filename For Input As FileNum
        While Not EOF(FileNum)
            Line Input #FileNum, Line
            
            ' TODO Theoretically there might be files that start with that string
            ' If we use Extended we should handle it correctly or better not at all as only writing "EXTM3U" on top is not valid: https://de.wikipedia.org/wiki/M3U, https://en.wikipedia.org/wiki/M3U
            If InStr(1, Line, "#EXT") <> 1 And Trim(Line) <> "" Then
                LstFilesDst.AddItem Line
            End If
        Wend
        Close FileNum
    Else
        MsgBox "File """ & Filename & """ does not exist", vbInformation, "File not found"
    End If

ShowOpenError:
End Sub

Private Sub CmdSelRem_Click()
    Dim i As Integer
    
    ' Remove selected files from output list
    For i = LstFilesDst.ListCount - 1 To 0 Step -1
        If LstFilesDst.Selected(i) = True Then LstFilesDst.RemoveItem (i)
    Next i
End Sub

Private Sub CmdM3uSave_Click()
    Dim Filename As String
    
    ' Try to save file
    On Error GoTo ShowSaveError
    With CdgMain
        .Flags = .Flags Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
        .Filter = "M3U Files (*.m3u)|*.m3u|All Files (*.*)|*.*"
        .CancelError = True
        .ShowSave
        Filename = .Filename
    End With
    On Error GoTo 0
    
    Dim FileNum As Integer
    Dim Line As Variant
    Dim i As Integer
    
    FileNum = FreeFile
    Open Filename For Output As FileNum

    ' Write file content
    Print #FileNum, "#EXTM3U"
    For i = 0 To LstFilesDst.ListCount
        Print #FileNum, LstFilesDst.List(i)
    Next i
    
    Close FileNum
    
ShowSaveError:
End Sub

Private Sub CmdSelAll_Click()
    Dim i As Integer
    
    For i = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(i) = True
    Next i
End Sub

Private Sub CmdSelNone_Click()
    Dim i As Integer
    
    For i = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(i) = False
    Next i
End Sub

Private Sub CmdSelInv_Click()
    Dim i As Integer
    
    For i = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(i) = Not LstFilesSrc.Selected(i)
    Next i
End Sub

Private Sub ListFiles(Optional Overwrite As Boolean = False)
    Dim Files() As String
    Dim File As String
    Dim FilePath As String
    Dim i As Integer

    ' First get the array of files
    Files = GetFiles(DirSrc.Path, IIf(ChkIncSubdir.Value = 0, 0, 1))
    
    If Not IsArrayEmpty(Files) Then
        If Not Overwrite Then LstFilesSrc.Clear
        
        For i = 0 To UBound(Files)
            File = Files(i)
            
            ' Add files to file list only if their name matches one of the patterns
            If FileMatchPattern(CStr(File), IIf(ChkApplyPattern.Value = 1, TxtPattern.Text, "*")) Then
                FilePath = IIf(PathAbs, DirSrc.Path & "\" & File, File)
                
                If Overwrite Then
                    LstFilesSrc.List(i) = FilePath ' Change item, but bor rebuild list, to keep selection and scroll position
                Else
                    LstFilesSrc.AddItem (FilePath)
                End If
            End If
        Next i
    End If
End Sub

Private Sub DirSrc_Change()
    ListFiles
End Sub

Private Sub DlbSrc_Change()
    On Error GoTo ErrorHandler
    
    ' Set new drive
    DirSrc.Path = DlbSrc.Drive
    
    ' If successfully changed the drive, save the new one as fallback for the next drive change
    SrcDriveOri = DlbSrc.Drive
    Exit Sub

ErrorHandler:
    MsgBox "Couldn't change to drive " & UCase(Mid$(DlbSrc.Drive, 1, 1)) & ": " & Error, vbExclamation, "Change Drive"
    DlbSrc.Drive = SrcDriveOri
End Sub

Private Sub Form_Load()
    PathAbs = False
    SrcDriveOri = DlbSrc.Drive
    ListFiles ' Initialize lists
End Sub

Private Sub LstFilesDst_DragDrop(Source As Control, x As Single, Y As Single)
    Dim i As Integer
    
    ' Append all selected files to output list
    If Source = LstFilesSrc Then
        For i = 0 To LstFilesSrc.ListCount - 1
            If LstFilesSrc.Selected(i) = True Then LstFilesDst.AddItem (LstFilesSrc.List(i))
        Next i
    End If
End Sub

Private Sub LstFilesSrc_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then LstFilesSrc.Drag vbBeginDrag
End Sub

Private Sub ToggleAbsRelPath()
    ' Update list with/without absolute paths
    ListFiles True
End Sub

Private Sub OptPathAbs_Click()
    PathAbs = True
    ToggleAbsRelPath
End Sub

Private Sub OptPathRel_Click()
    PathAbs = False
    ToggleAbsRelPath
End Sub

Private Function Basename(ByVal Path As String) As String
    Dim i As Integer
    
    ' Check beginning at the end where the first occurrence of "\" is -> take the string after that position
    For i = 0 To Len(Path) - 1
        If Mid(Path, Len(Path) - i, 1) = "\" Then
            Basename = Mid(Path, Len(Path) - i + 1, i)
            Exit Function
        End If
    Next i
End Function

Private Sub TmrSrcDir_Timer()
    Dim Drives As Long
    
    Drives = GetLogicalDrives
    If Drives <> LogicalDrives Then
        DlbSrc.Refresh
        LogicalDrives = Drives
    End If
End Sub

Private Sub TxtPattern_Change()
    If ChkApplyPattern.Value = 1 Then ListFiles
End Sub