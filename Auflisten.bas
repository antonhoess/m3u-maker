Attribute VB_Name = "Auflisten"
Option Explicit

Public Sub Listen(ByVal Pfad As String, Unterverzeichnisse As Boolean)
Dim I%, Anzahl%
I = 0

If Unterverzeichnisse = True Then
  Auflisten Pfad
  While I < Form1.List2.ListCount
    If (GetAttr(Form1.List2.List(I)) And vbDirectory) = vbDirectory Then
      Auflisten Form1.List2.List(I)
    End If
    I = I + 1
  Wend
End If

End Sub

Public Sub Auflisten(ByVal Pfad As String)
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
      Form1.List2.AddItem Pfad1 & Name & "\"
    End If  ' um ein Verzeichnis handelt.
  End If
  Name = Dir ' Nächsten Eintrag abrufen.
Loop

Fehler:
End Sub

