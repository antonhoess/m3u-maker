VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' COMDLG32 '

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" ( _
  pOpenfilename As OPENFILENAME _
) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" ( _
  pOpenfilename As OPENFILENAME _
) As Long
'----------------------------------------------------

Private Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    Flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type
'----------------------------------------------------

Private m_OF            As OPENFILENAME
'----------------------------------------------------

Public Property Get InitialDir() As String
    InitialDir = m_OF.lpstrInitialDir
End Property
Public Property Let InitialDir(RHS As String)
    m_OF.lpstrInitialDir = RHS
End Property

Public Property Get DialogTitle() As String
    DialogTitle = m_OF.lpstrTitle
End Property
Public Property Let DialogTitle(RHS As String)
    m_OF.lpstrTitle = RHS
End Property

Public Property Get Filter() As String
    Filter = m_OF.lpstrFilter
End Property
Public Property Let Filter(RHS As String)
    m_OF.lpstrFilter = RHS
End Property

Public Property Get FileTitle() As String
    FileTitle = _
        Left$( _
            m_OF.lpstrFileTitle, _
            InStr(1, m_OF.lpstrFileTitle, vbNullChar) - 1 _
        )
End Property

Public Property Get FilePath() As String
    FilePath = _
        Left$( _
            m_OF.lpstrFile, _
            InStr(1, m_OF.lpstrFile, vbNullChar) - 1 _
        )
End Property
'----------------------------------------------------

Public Function ShowOpen(Owner As Form) As Boolean
    With m_OF
        .lStructSize = LenB(m_OF)
        .hwndOwner = Owner.hWnd
        .hInstance = App.hInstance
        .lpstrFile = Space$(399)
        .nMaxFile = 400
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
    End With
    
    ShowOpen = CBool(GetOpenFileName(m_OF))
End Function

Public Function ShowSave(Owner As Form) As Boolean
    With m_OF
        .lStructSize = LenB(m_OF)
        .hwndOwner = Owner.hWnd
        .hInstance = App.hInstance
        .lpstrFile = Space$(399)
        .nMaxFile = 400
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
    End With
    
    ShowSave = CBool(GetSaveFileName(m_OF))
End Function






