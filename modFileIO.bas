Attribute VB_Name = "modFileIO"
Option Explicit
'Now come some Types.
Public Enum GFNFlags
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_EXPLORER = &H80000
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
End Enum
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
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private strfileName As OPENFILENAME
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Const email = "alienheretic@attbi.com"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public EP As Integer
Public SL As Integer
Public SA As Integer
Public ES As Integer
Public TexDir As String
Private Sub DialogFilter(WantedFilter As String)

  Dim intLoopCount As Integer

    strfileName.lpstrFilter = ""
    For intLoopCount = 1 To Len(WantedFilter)
        If Mid$(WantedFilter, intLoopCount, 1) = "|" Then strfileName.lpstrFilter = _
           strfileName.lpstrFilter + Chr$(0) Else strfileName.lpstrFilter = _
           strfileName.lpstrFilter + Mid$(WantedFilter, intLoopCount, 1)
    Next intLoopCount
    strfileName.lpstrFilter = strfileName.lpstrFilter + Chr$(0)

End Sub

'Wrapper Function.
Public Function DirBox(OwnerhWnd As Long, Msg As String, Directory As String) As String

  'Dimension some variables.
  
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
    
    'Set the message displayed on the dialog.
    szTitle = Msg
    
    'Set up the Type.
    With tBrowseInfo
        .hwndOwner = OwnerhWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    'Show the dialog box.
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    'Process the data returned.
    If (lpIDList) Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        DirBox = sBuffer
    End If

End Function

Public Function fncGetFileNametoSave(strFilter As String, strDefaultExtention As String, strInitalDirectory As String, Optional strDialogTitle As String = "Save") As String

  Dim lngReturnValue As Long
  Dim intRest As Integer

    strfileName.lpstrTitle = strDialogTitle
    strfileName.lpstrDefExt = strDefaultExtention
    strfileName.lpstrInitialDir = strInitalDirectory
    DialogFilter (strFilter)
    strfileName.hInstance = App.hInstance
    strfileName.lpstrFile = Chr$(0) & Space$(259)
    strfileName.nMaxFile = 260
    strfileName.Flags = &H80000 Or &H4
    strfileName.lStructSize = Len(strfileName)
    lngReturnValue = GetSaveFileName(strfileName)
    fncGetFileNametoSave = strfileName.lpstrFile

End Function

'This is the wrapper Function.
Public Function OpenDialog(ByVal OwnerhWnd As Long, _
                           ByVal Filters As String, ByVal FilterIndex As Long, _
                           ByVal FNameLength As Long, _
                           Optional ByVal InitFolder As String = "", _
                           Optional ByVal InitFileName As String = "", _
                           Optional ByVal dlgTitle As String = "", _
                           Optional ByVal Flags As GFNFlags = 0) As String

  'Dimension some variables.
  
  Dim GFN As OPENFILENAME
  Dim i As Long

    'Sort out some of the Type we will pass.
    GFN.lStructSize = Len(GFN)
    GFN.hwndOwner = OwnerhWnd
    GFN.hInstance = App.hInstance
    
    'Sort out the filters.
    For i = 1 To Len(Filters)
        If Mid$(Filters, i, 1) = "|" Then
            Filters = Left$(Filters, i - 1) & Chr$(0) & _
                      Right$(Filters, Len(Filters) - i)
        End If
    Next i
    
    'Finish setting up the Type.
    GFN.lpstrFilter = Filters
    GFN.nFilterIndex = FilterIndex
    GFN.lpstrFile = InitFileName & String$(FNameLength - Len(InitFileName), Chr$(0))
    GFN.nMaxFile = FNameLength
    GFN.lpstrFileTitle = String$(FNameLength, Chr$(0))
    GFN.nMaxFileTitle = FNameLength
    GFN.lpstrInitialDir = InitFolder
    GFN.lpstrTitle = dlgTitle
    GFN.Flags = Flags
    
    'Get and return the filename.
    If GetOpenFileName(GFN) >= 1 Then
        OpenDialog = GFN.lpstrFile
      Else
        OpenDialog = Chr$(0)
    End If

End Function

Public Sub SaveAllSettings()

    SaveSetting App.EXEName, "Info", "Version", App.Major & "." & App.Minor & "." & App.Revision

End Sub

'This is the wrapper Function.
Public Function SaveDialog(ByVal OwnerhWnd As Long, ByVal Filters As String, strDefaultExtension As String, ByVal FilterIndex As Long, ByVal FNameLength As Long, Optional ByVal InitFolder As String = "", Optional ByVal InitFileName As String = "", Optional ByVal dlgTitle As String = "", Optional ByVal Flags As GFNFlags = 0) As String

  'Dimension some variables.
  
  Dim GFN As OPENFILENAME
  Dim i As Long

    'Sort out some of the Type we will pass.
    GFN.lStructSize = Len(GFN)
    GFN.hwndOwner = OwnerhWnd
    GFN.hInstance = App.hInstance
    
    'Sort out the filters.
    For i = 1 To Len(Filters)
        If Mid$(Filters, i, 1) = "|" Then
            Filters = Left$(Filters, i - 1) & Chr$(0) & _
                      Right$(Filters, Len(Filters) - i)
        End If
    Next i
    
    'Finish setting up the Type.
    GFN.lpstrFilter = Filters
    GFN.nFilterIndex = FilterIndex
    GFN.lpstrFile = InitFileName & String$(FNameLength - Len(InitFileName), Chr$(0))
    GFN.nMaxFile = FNameLength
    GFN.lpstrFileTitle = String$(FNameLength, Chr$(0))
    GFN.nMaxFileTitle = FNameLength
    GFN.lpstrInitialDir = InitFolder
    GFN.nFileExtension = strDefaultExtension
    MsgBox GFN.nFileExtension
    GFN.lpstrTitle = dlgTitle
    GFN.Flags = Flags
    
    'Get and return the filename.
    If GetSaveFileName(GFN) >= 1 Then
        SaveDialog = GFN.lpstrFile
      Else
        SaveDialog = Chr$(0)
    End If

End Function

Public Sub sendemail()

  Dim Success As Long

    Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

':) Ulli's VB Code Formatter V2.5.12 (1/24/2002 12:41:23 AM) 69 + 169 = 238 Lines
