Attribute VB_Name = "FileHelper"
Option Explicit

Public Function IsArrayEmpty(Arr) As Boolean
    Dim Size As Long

    On Error Resume Next
    
    Size = UBound(Arr)
    IsArrayEmpty = (Err.Number <> 0)
End Function

Public Function FileExists(File As String) As Boolean
    On Error GoTo ErrorHandler

    FileLen (File) ' This might cause an error
    FileExists = True
ErrorHandler:
End Function

Public Function FileMatchPattern(File As String, Patterns As String) As Boolean
    Dim Pattern As Variant
    
    ' Split patterns in single pattern and check File against each pattern individually until one matches
    For Each Pattern In Split(Patterns, ";")
        If File Like Pattern Then
            FileMatchPattern = True
            Exit For
        End If
    Next Pattern
End Function

Public Function GetFilesRecursive(BasePath As String, RelPath As String, RecursionDepth As Integer, MaxRecursionDepth As Integer) As String()
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
    If RecursionDepth > MaxRecursionDepth Then GoTo Finish
    
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
    
    ' Fill array with files
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
        ' Fill array with folders
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
                ' Expand array
                ReDim Preserve Files(0 To FileCount + UBound(FilesSubdir))
                
                ' Store new values
                For I = 0 To UBound(FilesSubdir)
                    Files(FileCount + I) = RelPath & TmpFolder & "\" & FilesSubdir(I)
                Next I
                
                FileCount = UBound(Files) + 1
            End If
        Next TmpFolder
    End If
    
Finish:
    GetFilesRecursive = Files
End Function

Public Function GetFiles(BasePath As String, Optional IncludeSubDir As Boolean = False) As String()
    Dim Files() As String

    ' Initiate the recursive call
    Files = GetFilesRecursive(BasePath, "", 0, IIf(IncludeSubDir, 1, 0))
    GetFiles = Files
End Function
