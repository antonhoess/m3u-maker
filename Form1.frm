VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
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
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command5 
      Caption         =   "OK"
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hinzufügen"
      Height          =   495
      Left            =   11760
      Style           =   1  'Grafisch
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ListBox List3 
      DragIcon        =   "Form1.frx":0442
      Height          =   4155
      Left            =   6360
      MultiSelect     =   2  'Erweitert
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   6975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Unterverzeichnisse einbinden"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Markierung umkehren"
      Height          =   495
      Left            =   9960
      Style           =   1  'Grafisch
      TabIndex        =   13
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Markierung aufheben"
      Height          =   495
      Left            =   8160
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Alles markieren"
      Height          =   495
      Left            =   6360
      Style           =   1  'Grafisch
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "*.mp3;*.wav;*.wma;*.wmv"
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pattern anwenden"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Value           =   1  'Aktiviert
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Relativ"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   4680
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Absolut"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   4680
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.FileListBox File1 
      DragIcon        =   "Form1.frx":0884
      Height          =   285
      Left            =   4080
      MultiSelect     =   2  'Erweitert
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Liste speichern"
      Height          =   615
      Left            =   4440
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Markierte Elemente löschen"
      Height          =   615
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "M3U-Liste öffnen"
      Height          =   615
      Left            =   2280
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   120
      MultiSelect     =   2  'Erweitert
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Liste As String, I As Integer, Pfad As String, Löschcount As Integer, DName As String, Absolut As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
  File1.Pattern = Text1.Text
Else
  File1.Pattern = "*.*"
End If
Call Dir1_Change
End Sub

Private Sub Check2_Click()
Call Dir1_Change
End Sub

Private Sub Command1_Click()
If List3.ListIndex < 0 Then Exit Sub
Löschcount = List3.ListCount
For I = 0 To Löschcount - 1
Pfad = Dir1.List(Dir1.ListIndex) & "\" & List3.List(I)
If List3.Selected(I) = True Then List1.AddItem (Pfad)
Next I
Pfad = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.Clear
CommonDialog1.ShowOpen
Dim Datei As String
Dim Fnr As Long
Datei = CommonDialog1.filename
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
DName = "Arschi"
Liste = "#EXTM3U"
For I = 0 To List1.ListCount - 1
If Absolut = True Then
Liste = Liste & Chr(13) & List1.List(I)
Else
Liste = Liste & Chr(13) & GetFile(List1.List(I))
End If
Next I
CommonDialog1.ShowSave
Dim A
A = FreeFile
Open CommonDialog1.filename & ".m3u" For Output As A
Print #A, Liste
Close A
End Sub

Private Sub Command5_Click()
  If Check1.Value = 1 Then Call Dir1_Change
End Sub

Private Sub Command6_Click()
On Error Resume Next
For I = 0 To List3.ListCount - 1
  List3.Selected(I) = True
Next I
End Sub

Private Sub Command7_Click()
On Error Resume Next
For I = 0 To List3.ListCount - 1
  List3.Selected(I) = False
Next I
End Sub

Private Sub Command8_Click()
On Error Resume Next
For I = 0 To List3.ListCount - 1
  List3.Selected(I) = Not List3.Selected(I)
Next I
End Sub

Private Sub Dir1_Change()
On Error Resume Next
Dim I%, A%

List2.Clear
If Check2.Value = 0 Then
  If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
    Listen Dir1.Path, False
  Else
    Listen Dir1.Path & "\", False
  End If
Else
  If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
    Listen Dir1.Path, True
  Else
    Listen Dir1.Path & "\", True
  End If
End If

List3.Clear
File1.Path = Dir1.Path
For I = 0 To File1.ListCount - 1
  If File1.Path Like "?:\" = True Then
    List3.AddItem (File1.Path & File1.List(I))
  Else
    List3.AddItem (File1.Path & "\" & File1.List(I))
  End If
Next I
' Einträge auf gleicher Ebene wie Pfad

For I = 0 To List2.ListCount - 1
  File1.Path = List2.List(I)
  For A = 0 To File1.ListCount - 1
  If File1.Path Like "?:\" = True Then
    List3.AddItem (File1.Path & File1.List(A))
  Else
    List3.AddItem (File1.Path & "\" & File1.List(A))
  End If
  Next A
Next I
' Einträge auf tieferer Ebene wie Pfad

If Absolut = False Then
  For I = 0 To List3.ListCount - 1
    If Dir1.Path Like "?:\" = True Then
      List3.List(I) = Mid(List3.List(I), Len(Dir1.Path) + 1)
    Else
      List3.List(I) = Mid(List3.List(I), Len(Dir1.Path) + 2)
    End If
  Next I
End If
' Wenn Absolut = True, dann aus relativem einen absoluten Pfad machen
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
List1.Clear
Absolut = True
File1.Pattern = Text1.Text
File1.Refresh
End Sub

Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
'If Source = List3 Then
  If List3.ListIndex < 0 Then Exit Sub
  Löschcount = List3.ListCount
  For I = 0 To Löschcount - 1
  If List3.Selected(I) = True Then List1.AddItem (List3.List(I))
  Next I
  Pfad = ""
'End If
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then File1.Drag vbBeginDrag
End Sub

Private Sub Option1_Click()
Absolut = True
Call Dir1_Change
End Sub

Private Sub Option2_Click()
Absolut = False
Call Dir1_Change
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
End Sub
