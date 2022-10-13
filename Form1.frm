VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M3U-Maker-Extended"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   678
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   897
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   1680
      Width           =   495
   End
   Begin VB.Timer TmrSrcDir 
      Interval        =   500
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton CmdSelAdd 
      Caption         =   "Hinzufügen"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ListBox LstFilesSrc 
      DragIcon        =   "Form1.frx":0442
      Height          =   4155
      Left            =   6360
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   3615
   End
   Begin VB.CheckBox ChkIncSubdir 
      Caption         =   "Unterverzeichnisse einbinden"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   4680
      Width           =   2055
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
      TabIndex        =   13
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton CmdSelNone 
      Caption         =   "Markierung aufheben"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton CmdSelAll 
      Caption         =   "Alles markieren"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox TxtPattern 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "*.mp3;*.wav;*.wma;*.wmv"
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CheckBox ChkApplyPattern 
      Caption         =   "Pattern anwenden"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.OptionButton OptPathRel 
      Caption         =   "Relativ"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   4680
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton OptPathAbs 
      Caption         =   "Absolut"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   4680
      Width           =   855
   End
   Begin VB.FileListBox FilSrc 
      DragIcon        =   "Form1.frx":0884
      Height          =   4185
      Left            =   10080
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   120
      Width           =   3255
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
      Caption         =   "Markierte Elemente löschen"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton CmdM3uOpen 
      Caption         =   "M3U-Liste öffnen"
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
   Begin VB.Label Label1 
      Caption         =   "Pfade:"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
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

Dim PathAbs As Boolean
Dim SrcDriveOri As String
    
Private Sub ChkApplyPattern_Click()
    If ChkApplyPattern.Value = 1 Then
        FilSrc.Pattern = TxtPattern.Text
    Else
        FilSrc.Pattern = "*.*"
    End If
    
    Call DirSrc_Change
End Sub

Private Sub ChkIncSubdir_Click()
    Call DirSrc_Change ' XXX
End Sub

Private Sub CmdSelAdd_Click()
    Dim I As Integer
    Dim Pfad As String
    
    For I = 0 To LstFilesSrc.ListCount - 1
        Pfad = DirSrc.List(DirSrc.ListIndex) & "\" & LstFilesSrc.List(I)
        If LstFilesSrc.Selected(I) = True Then
            If PathAbs Then
                LstFilesDst.AddItem (Pfad)
            Else
                LstFilesDst.AddItem (Basename(Pfad))
            End If
        End If
    Next I
End Sub

Private Sub CmdM3uOpen_Click()
    On Error Resume Next
    LstFilesDst.Clear
    CdgMain.ShowOpen
    Dim Datei As String
    Dim Fnr As Long
    Datei = CdgMain.Filename
    Fnr = FreeFile
    Open Datei For Input As Fnr
    While Not EOF(Fnr)
        Line Input #Fnr, Zeile
        LstFilesDst.AddItem Zeile
        DoEvents
    Wend
    Close Fnr
    LstFilesDst.RemoveItem (0)
End Sub

Private Sub CmdSelRem_Click()
    Dim I As Integer
    
    For I = LstFilesDst.ListCount - 1 To 0 Step -1
        If LstFilesDst.Selected(I) = True Then LstFilesDst.RemoveItem (I)
    Next I
End Sub

Private Sub CmdM3uSave_Click()
    On Error Resume Next
    
    Dim Liste As String
    
    Liste = "#EXTM3U"
    For I = 0 To LstFilesDst.ListCount - 1
        Liste = Liste & vbNewLine & LstFilesDst.List(I)
    Next I
    CdgMain.ShowSave
    Dim A
    A = FreeFile
    Open CdgMain.Filename & ".m3u" For Output As A
    Print #A, Liste
    Close A
End Sub

Private Sub CmdSelAll_Click()
    Dim I As Integer
    
    For I = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(I) = True
    Next I
End Sub

Private Sub CmdSelNone_Click()
    Dim I As Integer
    
    For I = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(I) = False
    Next I
End Sub

Private Sub CmdSelInv_Click()
    Dim I As Integer
    
    For I = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(I) = Not LstFilesSrc.Selected(I)
    Next I
End Sub




Public Sub ListenX(ByVal Pfad As String, Unterverzeichnisse As Boolean)
    Dim I%, Anzahl%
    I = 0
    
    If Unterverzeichnisse = True Then
        Auflisten Pfad
        While I < FrmMain.List2.ListCount
            If (GetAttr(FrmMain.List2.List(I)) And vbDirectory) = vbDirectory Then
            Auflisten FrmMain.List2.List(I)
            End If
            I = I + 1
        Wend
    End If
End Sub

Public Sub AuflistenX(ByVal Pfad As String)
    On Error GoTo Fehler
    Dim Pfad1 As String, Name As String
    
    Pfad1 = Pfad
    'Add Pfad1
    Name = Dir$(Pfad1, vbDirectory) ' Ersten Eintrag abrufen.
    Do While Name <> ""  ' Schleife beginnen.
      ' Aktuelles und übergeordnetes Verzeichnis ignorieren.
      If Name <> "." And Name <> ".." Then
        ' Mit bit-weisem Vergleich sicherstellen, daß Name1 ein
    
    ' Verzeichnis ist.
        If (GetAttr(Pfad1 & Name) And vbDirectory) = vbDirectory Then
          FrmMain.List2.AddItem Pfad1 & Name & "\"
        End If  ' um ein Verzeichnis handelt.
      End If
      Name = Dir ' Nächsten Eintrag abrufen.
    Loop

Fehler:
End Sub

Private Function GetFilesRecursive(BasePath As String, RelPath As String, RecursionDepth As Integer, MaxRecursionDepth As Integer) As String()
    Dim Path As String
    Dim File As String
    Dim Files() As String
    Dim FileCount As Integer
    Dim Folder As String
    Dim Folders() As String
    Dim FolderCount As Integer
    Dim FilesSubdir() As String
    Dim I As Integer
    Dim TmpFolder As Variant
    
    ' Prevent recursive algortihm to go deeper than allowed
    If RecursionDepth = MaxRecursionDepth Then GoTo Finish
    
    ' Add trailing "\" if not already exisiting as this is necessary for Dir()
    If Len(BasePath) > 0 Then
        If Mid(BasePath, Len(BasePath), 1) <> "\" Then BasePath = BasePath & "\"
    End If
    
    If Len(RelPath) > 0 Then
        If Mid(RelPath, Len(RelPath), 1) <> "\" Then RelPath = RelPath & "\"
    End If
    
    ' Assemble full path
    Path = BasePath & RelPath
    
    ' Determine number of files in advance to prepare result array
    File = Dir$(Path)
    Do Until File = ""
        FileCount = FileCount + 1
        File = Dir$()
    Loop
    
    ' Fill array with the files
    If FileCount > 0 Then
        ReDim Files(0 To FileCount - 1)
        File = Dir$(Path)
        I = 0
        Do Until File = ""
            Files(I) = File 'RelPath & File
            I = I + 1
            File = Dir$()
        Loop
    End If
    
    ' We need to store the folders first as we cannot use Dir() recursively as described in
    ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dir-function
    
    ' Determine number of folders in advance to prepare result array
    Folder = Dir$(Path, vbDirectory)
    Do Until Folder = ""
        If Folder <> "." And Folder <> ".." And (GetAttr(Path & Folder) And vbDirectory) = vbDirectory Then FolderCount = FolderCount + 1
        Folder = Dir$()
    Loop
    
    If FolderCount > 0 Then
        ' Fill array with the folders
        ReDim Folders(0 To FolderCount - 1)
        Folder = Dir$(Path, vbDirectory)
        I = 0
        Do Until Folder = ""
            If Folder <> "." And Folder <> ".." And (GetAttr(Path & Folder) And vbDirectory) = vbDirectory Then
                Folders(I) = Folder
                I = I + 1
            End If
            Folder = Dir$()
        Loop
        
        ' Get all files recursively from subdirectories
        For Each TmpFolder In Folders
            ' Get files from subdir
            FilesSubdir = GetFilesRecursive(BasePath, RelPath & TmpFolder, RecursionDepth + 1, MaxRecursionDepth)
            
            If Not IsArrayEmpty(FilesSubdir) Then
                FileCount = UBound(Files) + 1
                
                ' Expand array
                ReDim Preserve Files(0 To FileCount + UBound(FilesSubdir))
                
                ' Store new values
                For I = 0 To UBound(FilesSubdir)
                    Files(FileCount + I) = RelPath & TmpFolder & "\" & FilesSubdir(I)
                Next I
            End If
        Next TmpFolder
    End If
    
Finish:
    GetFilesRecursive = Files
End Function

Private Sub GetFiles(BasePath As String, Optional IncludeSubDir As Boolean = False)
    Dim Result() As String

    Result = GetFilesRecursive(BasePath, "", 0, 2)
    
    Dim I As Integer
    For I = 0 To UBound(Result)
        Debug.Print Result(I)
    Next I
End Sub


Private Sub Command1_Click()
    Call GetFiles("C:\Users\ahoess\Downloads\__ABZULEGEN")
    
'    Dim File As Variant
'
'    For Each File In GetFiles("C:\Users\ahoess\Downloads\__ABZULEGEN")
'        Debug.Print File
'    Next File
End Sub

Private Sub DirSrc_Change()
    Dim I As Integer, A As Integer
    
    List2.Clear
    If ChkIncSubdir.Value = 0 Then
        If Mid(DirSrc.Path, Len(DirSrc.Path), 1) = "\" Then
            Listen DirSrc.Path, False
        Else
            Listen DirSrc.Path & "\", False
        End If
    Else
        If Mid(DirSrc.Path, Len(DirSrc.Path), 1) = "\" Then
            Listen DirSrc.Path, True
        Else
            Listen DirSrc.Path & "\", True
        End If
    End If
    
    ' Einträge auf gleicher Ebene wie Pfad
    LstFilesSrc.Clear
    FilSrc.Path = DirSrc.Path
    For I = 0 To FilSrc.ListCount - 1
        If FilSrc.Path Like "?:\" = True Then
            LstFilesSrc.AddItem (FilSrc.Path & FilSrc.List(I))
        Else
            LstFilesSrc.AddItem (FilSrc.Path & "\" & FilSrc.List(I))
        End If
    Next I
    
    ' Einträge auf tieferer Ebene wie Pfad
    For I = 0 To List2.ListCount - 1
        FilSrc.Path = List2.List(I)
        For A = 0 To FilSrc.ListCount - 1
            If FilSrc.Path Like "?:\" = True Then
                LstFilesSrc.AddItem (FilSrc.Path & FilSrc.List(A))
            Else
                LstFilesSrc.AddItem (FilSrc.Path & "\" & FilSrc.List(A))
            End If
        Next A
    Next I
    
    ' Wenn PathAbs = True, dann aus relativem einen absoluten Pfad machen
    If PathAbs = False Then
        For I = 0 To LstFilesSrc.ListCount - 1
            If DirSrc.Path Like "?:\" = True Then
                LstFilesSrc.List(I) = Mid(LstFilesSrc.List(I), Len(DirSrc.Path) + 1)
            Else
                LstFilesSrc.List(I) = Mid(LstFilesSrc.List(I), Len(DirSrc.Path) + 2)
            End If
        Next I
    End If
End Sub

Private Sub DlbSrc_Change()
    On Error GoTo ErrorHandler
    
    ' Set new drive
    DirSrc.Path = DlbSrc.Drive
    
    ' If successfully changed the drive, save the new one as fallback for the next drive change
    SrcDriveOri = DlbSrc.Drive
    
    Exit Sub

ErrorHandler:
    MsgBox "Couldn't change drive: " & Error, vbExclamation, "Change Drive"
    DlbSrc.Drive = SrcDriveOri
End Sub

Private Sub Form_Load()
    LstFilesDst.Clear
    PathAbs = False
    FilSrc.Pattern = TxtPattern.Text
    FilSrc.Refresh
    
    SrcDriveOri = DlbSrc.Drive
End Sub

Private Sub LstFilesDst_DragDrop(Source As Control, X As Single, Y As Single) ' XXX D&D checken
    Dim I As Integer
    'If Source = LstFilesSrc Then 'XXX
        For I = 0 To LstFilesSrc.ListCount - 1
            If LstFilesSrc.Selected(I) = True Then LstFilesDst.AddItem (LstFilesSrc.List(I))
        Next I
    'End If
End Sub

Private Sub LstFilesSrc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then FilSrc.Drag vbBeginDrag
End Sub

Private Sub ToggleAbsRelPath()
    Dim I As Integer
    Dim Sel() As Boolean
    ReDim Sel(0 To LstFilesSrc.ListCount - 1)
    
    ' Store selection (which is possible as the entries itself will not change)
    For I = 0 To DirSrc.ListCount - 1
        If LstFilesSrc.Selected(I) Then Sel(I) = True
    Next I
    
    ' Update list with/without absolute paths
    Call DirSrc_Change
    
    ' Restore selection
    For I = 0 To DirSrc.ListCount - 1
        If Sel(I) Then LstFilesSrc.Selected(I) = True
    Next I
End Sub

Private Sub OptPathAbs_Click()
    PathAbs = True
    Call ToggleAbsRelPath
End Sub

Private Sub OptPathRel_Click()
    PathAbs = False
    Call ToggleAbsRelPath
End Sub

Private Function Basename(ByVal Path As String) As String
    Dim I As Integer
    
    ' Check beginning at the end where the first occurrence of "\" is -> take the string after that position
    For I = 0 To Len(Path) - 1
        If Mid(Path, Len(Path) - I, 1) = "\" Then
            Basename = Mid(Path, Len(Path) - I + 1, I)
            Exit Function
        End If
    Next I
End Function

Private Sub TmrSrcDir_Timer()
    DlbSrc.Refresh
End Sub

Private Sub TxtPattern_Change()
    FilSrc.Pattern = TxtPattern.Text
    If ChkApplyPattern.Value = 1 Then Call DirSrc_Change
End Sub
