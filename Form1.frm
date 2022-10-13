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
   Begin VB.Timer TmrSrcDir 
      Interval        =   500
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton CmdSelAdd 
      Caption         =   "Hinzuf�gen"
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
      Width           =   6975
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
   Begin VB.FileListBox File1 
      DragIcon        =   "Form1.frx":0884
      Height          =   2040
      Left            =   4200
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   4320
      Width           =   3855
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
'Option Explicit

Dim Liste As String, I As Integer, Pfad As String, L�schcount As Integer, DName As String, PathAbs As Boolean
Dim SrcDriveOri As String
    
Private Sub ChkApplyPattern_Click()
    If ChkApplyPattern.Value = 1 Then
        File1.Pattern = TxtPattern.Text
    Else
        File1.Pattern = "*.*"
    End If
    
    Call DirSrc_Change
End Sub

Private Sub ChkIncSubdir_Click()
    Call DirSrc_Change
End Sub

Private Sub CmdSelAdd_Click()
    If LstFilesSrc.ListIndex < 0 Then Exit Sub
    L�schcount = LstFilesSrc.ListCount
    For I = 0 To L�schcount - 1
        Pfad = DirSrc.List(DirSrc.ListIndex) & "\" & LstFilesSrc.List(I)
        If LstFilesSrc.Selected(I) = True Then
            If PathAbs Then
                LstFilesDst.AddItem (Pfad)
            Else
                LstFilesDst.AddItem (Basename(Pfad))
            End If
        End If
    Next I
    
    Pfad = ""
End Sub

Private Sub CmdM3uOpen_Click()
    On Error Resume Next
    LstFilesDst.Clear
    CdgMain.ShowOpen
    Dim Datei As String
    Dim Fnr As Long
    Datei = CdgMain.FileName
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
    On Error Resume Next
    L�schcount = LstFilesDst.ListCount
    For I = 0 To L�schcount - 1
        If LstFilesDst.Selected(L�schcount - I - 1) = True Then LstFilesDst.RemoveItem (L�schcount - I - 1)
    Next I
End Sub

Private Sub CmdM3uSave_Click()
    On Error Resume Next
    DName = "Arschi"
    Liste = "#EXTM3U"
    For I = 0 To LstFilesDst.ListCount - 1
'        If PathAbs Then
        Liste = Liste & vbNewLine & LstFilesDst.List(I)
'        Else
'            Liste = Liste & vbNewLine & Basename(LstFilesDst.List(I))
'        End If
    Next I
    CdgMain.ShowSave
    Dim A
    A = FreeFile
    Open CdgMain.FileName & ".m3u" For Output As A
    Print #A, Liste
    Close A
End Sub

Private Sub CmdSelAll_Click()
    On Error Resume Next
    For I = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(I) = True
    Next I
End Sub

Private Sub CmdSelNone_Click()
    On Error Resume Next
    For I = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(I) = False
    Next I
End Sub

Private Sub CmdSelInv_Click()
    On Error Resume Next
    For I = 0 To LstFilesSrc.ListCount - 1
        LstFilesSrc.Selected(I) = Not LstFilesSrc.Selected(I)
    Next I
End Sub

Private Sub DirSrc_Change()
    On Error Resume Next
    Dim I%, A%
    
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
    
    ' Eintr�ge auf gleicher Ebene wie Pfad
    LstFilesSrc.Clear
    File1.Path = DirSrc.Path
    For I = 0 To File1.ListCount - 1
        If File1.Path Like "?:\" = True Then
            LstFilesSrc.AddItem (File1.Path & File1.List(I))
        Else
            LstFilesSrc.AddItem (File1.Path & "\" & File1.List(I))
        End If
    Next I
    
    ' Eintr�ge auf tieferer Ebene wie Pfad
    For I = 0 To List2.ListCount - 1
        File1.Path = List2.List(I)
        For A = 0 To File1.ListCount - 1
            If File1.Path Like "?:\" = True Then
                LstFilesSrc.AddItem (File1.Path & File1.List(A))
            Else
                LstFilesSrc.AddItem (File1.Path & "\" & File1.List(A))
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
    
    DirSrc.Path = DlbSrc.Drive
    SrcDriveOri = DlbSrc.Drive
    
    Exit Sub

ErrorHandler:
    MsgBox "Couldn't change drive: " & Error, vbExclamation, "Change Drive"
    DlbSrc.Drive = SrcDriveOri
End Sub

Private Sub Form_Load()
    LstFilesDst.Clear
    PathAbs = False
    File1.Pattern = TxtPattern.Text
    File1.Refresh
    
    SrcDriveOri = DlbSrc.Drive
End Sub

Private Sub LstFilesDst_DragDrop(Source As Control, X As Single, Y As Single)
    On Error Resume Next
    'If Source = LstFilesSrc Then
        If LstFilesSrc.ListIndex < 0 Then Exit Sub
            L�schcount = LstFilesSrc.ListCount
            For I = 0 To L�schcount - 1
                If LstFilesSrc.Selected(I) = True Then LstFilesDst.AddItem (LstFilesSrc.List(I))
            Next I
        Pfad = ""
    'End If
End Sub

Private Sub LstFilesSrc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then File1.Drag vbBeginDrag
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
    On Error Resume Next
    File1.Pattern = TxtPattern.Text
    If ChkApplyPattern.Value = 1 Then Call DirSrc_Change
End Sub
