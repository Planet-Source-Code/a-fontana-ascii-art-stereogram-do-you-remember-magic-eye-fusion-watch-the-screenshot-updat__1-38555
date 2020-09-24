VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
''  IT's NOT MY CODE!
''





Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private FileName As OPENFILENAME

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Public Property Let DefaultExtension(Extention As String)
    ' Default extension for files
    FileName.lpstrDefExt = Extention
End Property

Public Property Get DefaultExtension() As String
    ' Default extension for files
    DefaultExtension = FileName.lpstrDefExt
End Property

Public Property Let ObjectOwner(Objet As Object)
    ' Object that 'owns' the common dialog.
    FileName.hwndOwner = Objet.hwnd
End Property

Public Property Let Filter(CustomFilter As String)
    ' File filter. Format : "Text for filter|filter|Next Text|Next filter". Ex : "All Files (*.*)|*.*|Text files (*.txt)|*.txt"
    Dim intCount As Integer
    FileName.lpstrFilter = ""
    For intCount = 1 To Len(CustomFilter)
        If Mid(CustomFilter, intCount, 1) = "|" Then FileName.lpstrFilter = FileName.lpstrFilter + Chr(0) Else FileName.lpstrFilter = FileName.lpstrFilter + Mid(CustomFilter, intCount, 1)
    Next intCount
    FileName.lpstrFilter = FileName.lpstrFilter + Chr(0)
End Property

Public Property Let WindowTitle(Title As String)
    ' Window title
    FileName.lpstrTitle = Title
End Property

Public Property Get WindowTitle() As String
    ' Window title
    WindowTitle = FileName.lpstrTitle
End Property

Public Property Let InitialDirectory(InitDir As String)
    ' Initial folder
    FileName.lpstrInitialDir = InitDir
End Property

Public Property Get InitialDirectory() As String
    ' Initial folder
    InitialDirectory = FileName.lpstrInitialDir
End Property

Public Function GetFileOpenName(Optional Multiselect As Boolean = False) As String
    ' Show common dialog open file control and returns file name. Use GetMultiFilename function to
    ' extract filenames when the multiselect flag is true. The Count function returns the number
    ' of files selected.
    FileName.hInstance = App.hInstance
    FileName.hwndOwner = FrmMain.hwnd
    FileName.lpstrFile = Chr(0) & Space(259)
    FileName.nMaxFile = 260
    If Multiselect Then FileName.flags = &H80000 Or &H4 Or &H200 Else FileName.flags = &H80000 Or &H4
    FileName.lStructSize = Len(FileName)
    GetOpenFileName FileName
    GetFileOpenName = FileName.lpstrFile
End Function

Public Function GetFileSaveName() As String
    ' Show common dialog save file control and returns file name.
    FileName.hInstance = App.hInstance
    FileName.hwndOwner = FrmMain.hwnd
    FileName.lpstrFile = Chr(0) & Space(259)
    FileName.nMaxFile = 260
    FileName.flags = &H80000 Or &H4
    FileName.lStructSize = Len(FileName)
    GetSaveFileName FileName
    GetFileSaveName = FileName.lpstrFile
End Function

Public Function Count() As Integer
    ' Returns the number of file selected. Use with GetMultiFilename to extract filename when multiselect is true.
    Dim intCount As Integer
    For intCount = 1 To Trim(Len(FileName.lpstrFile))
        If Mid(Trim(FileName.lpstrFile), intCount, 1) = Chr(0) Then Count = Count + 1
    Next intCount
    Count = Count - 2
    If Count < 1 Then Count = Count + 1
End Function

Public Function GetMultiFilename(Filenumber As Integer) As String
    ' Returns the filename of the specified filenumber. Use only with open file dialog when multiselect is true.
    Dim intCount As Integer
    Dim intOne As Integer
    Dim intFile As Integer
    Dim intNext As Integer
    intOne = InStr(1, Trim(FileName.lpstrFile), Chr(0))
    intFile = 1
    For intCount = 1 To Filenumber
        intFile = InStr(intFile + 1, Trim(FileName.lpstrFile), Chr(0))
    Next intCount
    intNext = InStr(intFile + 1, Trim(FileName.lpstrFile), Chr(0))
    GetMultiFilename = IIf(Right(Mid(Trim(FileName.lpstrFile), 1, intOne - 1), 1) = "\", Mid(Trim(FileName.lpstrFile), 1, intOne - 1), Mid(Trim(FileName.lpstrFile), 1, intOne - 1) + "\") + Mid(Trim(FileName.lpstrFile), intFile + 1, intNext - intFile - 1)
    If Right(GetMultiFilename, 1) = "\" Then GetMultiFilename = Left(GetMultiFilename, Len(GetMultiFilename) - 1)
End Function
