Attribute VB_Name = "Module2"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hmem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Type BROWSEINFO
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Enum BIF
    BROWSEFORCOMPUTER = &H1000
    BROWSEFORPRINTER = &H2000
    BROWSEINCLUDEFILES = &H4000
    BROWSEINCLUDEURLS = &H80
    DONTGOBELOWDOMAIN = &H2
    EDITBOX = &H10
    NEWDIALOGSTYLE = &H40
    RETURNFSANCESTORS = &H8
    RETURNONLYFSDIRS = &H1
    SHAREABLE = &H8000
    STATUSTEXT = &H4
    USENEWUI = &H40
    VALIDATE_BIF = &H20
End Enum

Public Const MAX_PATH = 260
Public Function BrowseForFolder(hWndOwner As Long, sTitle As String) As String
    
Dim BInfo As BROWSEINFO
Dim lpIDList As Long

    With BInfo
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF.EDITBOX
    End With

    lpIDList = SHBrowseForFolder(BInfo)
    If lpIDList Then
        BrowseForFolder = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, BrowseForFolder
        CoTaskMemFree lpIDList
        BrowseForFolder = StripNulls(BrowseForFolder)
    End If
End Function
Public Function IsFile(ByVal lpFileName As String) As Boolean
    If PathFileExists(lpFileName) = 1 And PathIsDirectory(lpFileName) = 0 Then
        IsFile = True
    Else
        IsFile = False
    End If
End Function
Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function



