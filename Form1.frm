VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "M3U-Maker"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command8 
      Caption         =   "Markierung umkehren"
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Markierung aufheben"
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Alles markieren"
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Text            =   "*.mp3;*.wav;*.wma;*.wmv"
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pattern anwenden"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   3240
      Value           =   1  'Aktiviert
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Relativ"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   3480
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Absolut"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.FileListBox File1 
      DragIcon        =   "Form1.frx":0442
      Height          =   3210
      Left            =   6360
      MultiSelect     =   2  'Erweitert
      TabIndex        =   8
      Top             =   120
      Width           =   5295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   2280
      TabIndex        =   7
      Top             =   600
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ordner hinzufügen"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Liste speichern"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Markierte Elemente löschen"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "M3U-Liste öffnen"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   240
      MultiSelect     =   2  'Erweitert
      TabIndex        =   1
      Top             =   3960
      Width           =   11415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Einzelne Dateien zu Liste hinzufügen"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3360
      Y1              =   3720
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   2040
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 259

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Dim Liste As String, I As Integer, Pfad As String, Löschcount As Integer, Absolut As Boolean
Dim CommonDialog1 As New CommonDialog

'************************************************************************

Private Sub Check1_Click()
If Check1.Value = 1 Then
  File1.Pattern = Text1.Text
Else
  File1.Pattern = "*.*"
End If
File1.Refresh
End Sub

Private Sub Command1_Click()
On Error Resume Next
If File1.ListIndex < 0 Then Exit Sub
Löschcount = File1.ListCount
For I = 0 To Löschcount - 1
Pfad = Dir1.List(Dir1.ListIndex) & "\" & File1.List(I)
If File1.Selected(I) = True Then List1.AddItem (Pfad)
Next I
Pfad = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
  List1.Clear
  'CommonDialog1.Filter = "M3U-Playlists (*.m3u)|*.m3u"
  If (CommonDialog1.ShowOpen(Form1)) = False Then Exit Sub
  Dim Datei As String
  Dim Fnr As Long
  Dim Zeile As String
  Datei = CommonDialog1.FileTitle
  Fnr = FreeFile
  Open Datei For Input As Fnr
  While Not EOF(Fnr)
    Line Input #Fnr, Zeile
    List1.AddItem Zeile
    DoEvents
  Wend
  Close Fnr
  List1.RemoveItem (0)
End Sub

Private Sub Command3_Click()
On Error Resume Next
  Löschcount = List1.ListCount
  For I = 0 To Löschcount - 1
    If List1.Selected(Löschcount - I - 1) = True Then List1.RemoveItem (Löschcount - I - 1)
  Next I
End Sub

Private Sub Command4_Click()
On Error Resume Next
  Liste = ""
  For I = 0 To List1.ListCount - 1
    If Absolut = True Then
      Liste = Liste & List1.List(I)
      If I < List1.ListCount - 1 Then Liste = Liste & vbNewLine
    Else
      Liste = Liste & GetFile(List1.List(I))
      If I < List1.ListCount - 1 Then Liste = Liste & vbNewLine
    End If
  Next I
  If CommonDialog1.ShowSave(Form1) = True Then
    Dim A
    A = FreeFile
    Open CommonDialog1.FileTitle & ".m3u" For Output As A
    Print #A, Liste
    Close A
  End If
End Sub

Private Sub Command5_Click()
On Error GoTo Ende
  For I = 0 To File1.ListCount - 1
    Pfad = Dir1.List(Dir1.ListIndex) & "\" & File1.List(I)
    List1.AddItem (Pfad)
    Pfad = ""
  Next I
Ende:
End Sub

Private Sub Command6_Click()
On Error Resume Next
  For I = 0 To File1.ListCount - 1
    File1.Selected(I) = True
  Next I
End Sub

Private Sub Command7_Click()
On Error Resume Next
  For I = 0 To File1.ListCount - 1
    File1.Selected(I) = False
  Next I
End Sub

Private Sub Command8_Click()
On Error Resume Next
  For I = 0 To File1.ListCount - 1
    File1.Selected(I) = Not File1.Selected(I)
  Next I
End Sub

Private Sub Dir1_Change()
On Error Resume Next
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then File1.Drag vbBeginDrag
End Sub

Private Sub Form_Load()
On Error GoTo Fehler
Dim Hilf1() As String, Hilf2() As String, Pfad As String, Liste As String
  List1.Clear
  Absolut = False
  File1.Pattern = Text1.Text
  File1.Refresh

  If Command <> "" Then
    Pfad = Mid(Command, 2, Len(Command) - 2)
    GetFiles "*.mp3", Pfad
    For I = 0 To List1.ListCount - 1
      Liste = Liste & List1.List(I)
      If I < List1.ListCount - 1 Then Liste = Liste & vbNewLine
    Next I
    ChDir Pfad
    Dim A
    A = FreeFile

    Hilf1 = Split(Pfad, "\")
    Hilf2 = Split(Hilf1(UBound(Hilf1)), " - ")

    Select Case UBound(Hilf2)
      Case 2:
        Open Pfad & "\" & Hilf2(UBound(Hilf2) - 1) & " - " & Hilf2(UBound(Hilf2)) & ".m3u" For Output As A
      Case Else:
        If Hilf1(UBound(Hilf1)) Like "CD #" Then
          Open Pfad & "\" & Right(Hilf1(UBound(Hilf1) - 1), Len(Hilf1(UBound(Hilf1) - 1)) - InStr(Hilf1(UBound(Hilf1) - 1), " - ") - 2) & " - " & Hilf2(UBound(Hilf2)) & ".m3u" For Output As A
        Else
          Open Pfad & "\" & Hilf2(UBound(Hilf2)) & ".m3u" For Output As A
        End If
    End Select
        
    Print #A, Liste
    Close A
    End
  End If
Exit Sub

Fehler:
  If Err.Number = 76 Then MsgBox "Sie müssen einen gültigen Ordner auf das Programmicon ziehen, um eine M3U-Playlist zu erstellen!", , "Error": End
End Sub

Private Sub List1_DblClick()
  For I = 0 To List1.ListCount - 1
    List1.Selected(I) = True
  Next I
End Sub

Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
  If Source = File1 Then
    If File1.ListIndex < 0 Then Exit Sub
    Löschcount = File1.ListCount
    For I = 0 To Löschcount - 1
    Pfad = Dir1.List(Dir1.ListIndex) & "\" & File1.List(I)
    If File1.Selected(I) = True Then List1.AddItem (Pfad)
    Next I
    Pfad = ""
  End If
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Anz As Integer
  If Button = 2 Then
    Anz = List1.ListCount
    For I = 1 To Anz
      If List1.Selected(Anz - I) = True Then List1.RemoveItem (Anz - I)
    Next I
  End If
End Sub

Private Sub Option1_Click()
  Absolut = True
End Sub

Private Sub Option2_Click()
  Absolut = False
End Sub

Private Function GetFile(ByVal Path As String) As String
Dim I As Integer
  For I = 0 To Len(Path) - 1
    If Mid(Path, Len(Path) - I, 1) = "\" Then
      GetFile = Mid(Path, Len(Path) - I + 1, I)
      Exit Function
    End If
  Next I
End Function

Private Sub Text1_Change()
On Error Resume Next
  File1.Pattern = Text1.Text
  File1.Refresh
End Sub

Private Sub GetFiles(Patt As String, Root As String)
Dim File$, hFile&, FD As WIN32_FIND_DATA
  List1.Clear
  hFile = FindFirstFile(Root & "\" & Patt, FD)
  If hFile = 0 Then Exit Sub
  Do
     File = Left(FD.cFileName, InStr(FD.cFileName, Chr(0)) - 1)
     If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then
       List1.AddItem File
     End If
  Loop While FindNextFile(hFile, FD)
  Call FindClose(hFile)
End Sub
