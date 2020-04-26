VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EOF data in EXE by The trick"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdActivate 
      Caption         =   "Activate..."
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1980
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove file..."
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Add file..."
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   900
      Width           =   1815
   End
   Begin VB.CommandButton cmdSaveEOF 
      Caption         =   "Save EOF data..."
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox lstFiles 
      Height          =   2100
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblList 
      Caption         =   "Files list:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   2835
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' //
' // This form demonstrates how to work with CEmbeddedFiles class
' // You can add files to any module using this utility
' // By The trick, 2020
' //

Option Explicit

Private Const SCS_32BIT_BINARY        As Long = 0
Private Const FILE_ATTRIBUTE_NORMAL   As Long = &H80
Private Const INVALID_HANDLE_VALUE    As Long = -1
Private Const GENERIC_READ            As Long = &H80000000
Private Const GENERIC_WRITE           As Long = &H40000000
Private Const OPEN_EXISTING           As Long = 3
Private Const FILE_SHARE_READ         As Long = &H1
Private Const OFN_ALLOWMULTISELECT    As Long = &H200
Private Const OFN_EXPLORER            As Long = &H80000
Private Const SND_ASYNC               As Long = &H1
Private Const SND_MEMORY              As Long = &H4
Private Const MAX_PATH                As Long = 260
Private Const FILE_BEGIN              As Long = 0
Private Const CREATE_ALWAYS           As Long = 2

Private Type LARGE_INTEGER
    lowPart             As Long
    highPart            As Long
End Type

Private Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As Long
    lpstrCustomFilter   As Long
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As Long
    nMaxFile            As Long
    lpstrFileTitle      As Long
    nMaxFileTitle       As Long
    lpstrInitialDir     As Long
    lpstrTitle          As Long
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As Long
End Type

Private Type STARTUPINFO
    cb                  As Long
    lpReserved          As Long
    lpDesktop           As Long
    lpTitle             As Long
    dwX                 As Long
    dwY                 As Long
    dwXSize             As Long
    dwYSize             As Long
    dwXCountChars       As Long
    dwYCountChars       As Long
    dwFillAttribute     As Long
    dwFlags             As Long
    wShowWindow         As Integer
    cbReserved2         As Integer
    lpReserved2         As Long
    hStdInput           As Long
    hStdOutput          As Long
    hStdError           As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess            As Long
    hThread             As Long
    dwProcessId         As Long
    dwThreadId          As Long
End Type

Private Declare Function GetBinaryType Lib "kernel32" _
                         Alias "GetBinaryTypeW" ( _
                         ByVal lpApplicationName As Long, _
                         ByRef lpBinaryType As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
                         Alias "GetSaveFileNameW" ( _
                         pOpenfilename As OPENFILENAME) As Long
Private Declare Function CreateFile Lib "kernel32" _
                         Alias "CreateFileW" ( _
                         ByVal lpFileName As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwShareMode As Long, _
                         ByRef lpSecurityAttributes As Any, _
                         ByVal dwCreationDisposition As Long, _
                         ByVal dwFlagsAndAttributes As Long, _
                         ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByRef lpBuffer As Any, _
                         ByVal nNumberOfBytesToRead As Long, _
                         ByRef lpNumberOfBytesRead As Long, _
                         ByRef lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByRef lpBuffer As Any, _
                         ByVal nNumberOfBytesToWrite As Long, _
                         ByRef lpNumberOfBytesWritten As Long, _
                         ByRef lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByVal lDistanceToMove As Long, _
                         ByRef lpDistanceToMoveHigh As Any, _
                         ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileSize As LARGE_INTEGER) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
                         Alias "GetOpenFileNameW" ( _
                         ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare Function PathFindFileName Lib "shlwapi" _
                         Alias "PathFindFileNameW" ( _
                         ByVal pszPath As Long) As Long
Private Declare Function PathRemoveFileSpec Lib "shlwapi" _
                         Alias "PathRemoveFileSpecW" ( _
                         ByVal pszPath As Long) As Long
Private Declare Function PathYetAnotherMakeUniqueName Lib "shell32" ( _
                         ByVal pszUniqueName As Long, _
                         ByVal pszPath As Long, _
                         ByVal pszShort As Long, _
                         ByVal pszFileSpec As Long) As Long
Private Declare Function PathFindExtension Lib "shlwapi" _
                         Alias "PathFindExtensionW" ( _
                         ByVal pszPath As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" _
                         Alias "sndPlaySoundW" ( _
                         ByRef lpszSoundName As Any, _
                         ByVal uFlags As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" _
                         Alias "GetModuleFileNameW" ( _
                         ByVal hModule As Long, _
                         ByVal lpFileName As Long, _
                         ByVal nSize As Long) As Long
Private Declare Function lstrcmpi Lib "kernel32" _
                         Alias "lstrcmpiW" ( _
                         ByVal lpString1 As Long, _
                         ByVal lpString2 As Long) As Long
Private Declare Function MoveFile Lib "kernel32" _
                         Alias "MoveFileW" ( _
                         ByVal lpExistingFileName As Long, _
                         ByVal lpNewFileName As Long) As Long
Private Declare Function CopyFile Lib "kernel32" _
                         Alias "CopyFileW" ( _
                         ByVal lpExistingFileName As Long, _
                         ByVal lpNewFileName As Long, _
                         ByVal bFailIfExists As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" ( _
                         ByVal hFile As Long) As Long
Private Declare Function CommandLineToArgvW Lib "shell32" ( _
                         ByVal lpCmdLine As Long, _
                         ByRef pNumArgs As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32" _
                         Alias "lstrcpynW" ( _
                         ByVal lpString1 As Long, _
                         ByVal lpString2 As Long, _
                         ByVal iMaxLength As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" ( _
                         ByVal hMem As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetCommandLine Lib "kernel32" _
                         Alias "GetCommandLineW" () As Long
Private Declare Function lstrlen Lib "kernel32" _
                         Alias "lstrlenW" ( _
                         ByVal lpString As Long) As Long
Private Declare Function PathUnquoteSpaces Lib "shlwapi" _
                         Alias "PathUnquoteSpacesW" ( _
                         ByVal p As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" _
                         Alias "DeleteFileW" ( _
                         ByVal lpFileName As Long) As Long
Private Declare Function Sleep Lib "kernel32" ( _
                         ByVal dwMilliseconds As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" _
                         Alias "PathFileExistsW" ( _
                         ByVal pszFile As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" _
                         Alias "CreateProcessW" ( _
                         ByVal lpApplicationName As Long, _
                         ByVal lpCommandLine As Long, _
                         ByRef lpProcessAttributes As Any, _
                         ByRef lpThreadAttributes As Any, _
                         ByVal bInheritHandles As Long, _
                         ByVal dwCreationFlags As Long, _
                         ByRef lpEnvironment As Any, _
                         ByVal lpCurrentDirectory As Long, _
                         ByRef lpStartupInfo As STARTUPINFO, _
                         ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Sub GetStartupInfo Lib "kernel32" _
                    Alias "GetStartupInfoW" ( _
                    ByRef lpStartupInfo As STARTUPINFO)
                         
Private m_cEmbData      As CEmbeddedFiles   ' // Current module.
                                            ' // To use it in your own project just include it and create an instance.
                                            ' // Then you can use embedded files by calling Initialize(App.hInstance).
                                            ' // Alternatively you can specify other modules (dll/ocx).
Private m_bPreviewMode  As Boolean          ' // Preview mode for form

' // If true. Form just loads for preview content (just to reduce forms in project).
Public Property Let PreviewMode( _
                    ByVal bValue As Boolean)
    m_bPreviewMode = bValue
End Property

' // Preview a picture
Public Sub PreviewPicture( _
           ByVal cPic As StdPicture)
    
    Me.Caption = "Picture preview"
    Me.Width = Me.ScaleX(cPic.Width, vbHimetric, vbTwips) + (Me.Width - Me.ScaleWidth)
    Me.Height = Me.ScaleY(cPic.Height, vbHimetric, vbTwips) + (Me.Height - Me.ScaleHeight)
    
    Set Me.Picture = cPic
            
    Me.Show vbModal
            
End Sub

' // Preview a sound. The sound should be in WAVE format
Public Sub PreviewSound( _
           ByRef bData() As Byte)
    
    Me.Caption = "Sound preview"
    Me.Width = 3000 + (Me.Width - Me.ScaleWidth)
    Me.Height = 200 + (Me.Height - Me.ScaleHeight)
    
    If sndPlaySound(bData(0), SND_ASYNC Or SND_MEMORY) Then
        Me.Show vbModal
        sndPlaySound ByVal 0&, 0
    End If
    
End Sub

' // Preview data as HEX
Public Sub PreviewRaw( _
           ByRef bData() As Byte)
    Dim lIndex  As Long
    Dim lRows   As Long
    Dim lHeight As Long
    Dim sHex    As String
    Dim sText   As String
    
    Me.Caption = "Raw data"
    Me.FontName = "Courier New"
    
    If (Not Not bData) = 0 Then
    
        Me.Width = Me.TextWidth("0") * 78 + (Me.Width - Me.ScaleWidth)
        Me.Height = Me.TextHeight("0") * 1 + (Me.Height - Me.ScaleHeight)
        Me.Show vbModal
        Exit Sub
        
    End If
    
    lRows = -Int(-(UBound(bData) + 1) / 16) + 1
    lHeight = Me.TextHeight("0") * lRows + (Me.Height - Me.ScaleHeight)
    
    If lHeight > Screen.Height * 0.75 Then
        lHeight = Screen.Height * 0.75
    End If
    
    Me.Width = Me.TextWidth("0") * 78 + (Me.Width - Me.ScaleWidth)
    Me.Height = lHeight
    Me.AutoRedraw = True
    
    For lIndex = 0 To UBound(bData)
        
        If (lIndex Mod 16) = 0 Then
            
            sHex = Hex$(lIndex)

            Do While Len(sHex) < 8
                sHex = "0" & sHex
            Loop
            
            If Len(sText) Then
                Me.Print ": "; sText
            End If
            
            sText = vbNullString
            
            If Me.CurrentY > Me.ScaleHeight Then Exit For
            
            Me.Print sHex; ": ";
            
        End If
        
        sHex = Hex$(bData(lIndex))
        
        If Len(sHex) = 1 Then
            sHex = "0" & sHex
        End If
        
        Me.Print sHex; " ";
        
        If bData(lIndex) >= 32 Then
            sText = sText & Chr$(bData(lIndex))
        Else
            sText = sText & "."
        End If
        
    Next
    
    If Len(sText) Then
        Me.Print Spc((16 - (lIndex Mod 16)) * 3); ": "; sText
    End If
    
    Me.Show vbModal
    
End Sub

' // Extract command line
Private Function ParseCommandLine( _
                 ByRef sArgs() As String) As Long
    Dim pCmdLine    As Long
    Dim lCount      As Long
    Dim lIndex      As Long
    Dim lStrLength  As Long
    Dim pArg        As Long
    Dim bIsInIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    If bIsInIDE Then
        pCmdLine = CommandLineToArgvW(StrPtr(Chr$(34) & App.Path & "\" & App.EXEName & _
                                      Chr$(34) & " " & Command$()), lCount)
    Else
        pCmdLine = CommandLineToArgvW(GetCommandLine(), lCount)
    End If
    
    If lCount < 1 Then Exit Function
    
    ReDim sArgs(lCount - 1)
    
    For lIndex = 0 To lCount - 1
    
        GetMem4 ByVal pCmdLine + lIndex * 4, pArg
        
        lStrLength = lstrlen(ByVal pArg)
        sArgs(lIndex) = Space$(lStrLength)
        lstrcpyn ByVal StrPtr(sArgs(lIndex)), ByVal pArg, lStrLength + 1
        
    Next
    
    GlobalFree pCmdLine
    
    ParseCommandLine = lCount
                    
End Function

Private Function MakeTrue( _
                 ByRef bvar As Boolean) As Boolean
    bvar = True: MakeTrue = True
End Function

' // Get saved file name
Private Function GetSaveFile( _
                 ByVal hWnd As Long, _
                 ByRef sTitle As String, _
                 ByRef sFilter As String, _
                 ByRef sDefExtension As String) As String
    Dim tOfn            As OPENFILENAME
    Dim strOutputFile   As String
    
    With tOfn
    
        .nMaxFile = 260
        strOutputFile = String$(.nMaxFile, vbNullChar)
        
        .hwndOwner = hWnd
        .lpstrTitle = StrPtr(sTitle)
        .lpstrFile = StrPtr(strOutputFile)
        .lStructSize = Len(tOfn)
        .lpstrFilter = StrPtr(sFilter)
        .lpstrDefExt = StrPtr(sDefExtension)
        
        If GetSaveFileName(tOfn) = 0 Then Exit Function
        
        GetSaveFile = Left$(strOutputFile, InStr(1, strOutputFile, vbNullChar) - 1)
        
    End With

End Function

' // Get list of opened files. The first entry contains the directory.
Private Function GetOpenFiles( _
                 ByVal hWnd As Long, _
                 ByRef sTitle As String, _
                 ByRef sFilter As String) As String()
    Dim tOfn    As OPENFILENAME:   Dim sOut    As String
    Dim lIc     As Long:           Dim lIo     As Long
    Dim lPos    As Long:           Dim sRet()  As String
    Dim lIndex  As Long
    
    tOfn.flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
    tOfn.nMaxFile = 32767
    
    sOut = String(32767, vbNullChar)

    tOfn.hwndOwner = hWnd
    tOfn.lpstrTitle = StrPtr(sTitle)
    tOfn.lpstrFile = StrPtr(sOut)
    tOfn.lStructSize = Len(tOfn)
    tOfn.lpstrFilter = StrPtr(sFilter)
    
    If GetOpenFileName(tOfn) Then
    
        ReDim sRet(9)
        
        sRet(0) = Left$(sOut, tOfn.nFileOffset - 1)
        lIo = tOfn.nFileOffset + 1: lPos = lPos + 1
        lIc = InStr(lIo, sOut, vbNullChar)
        lIndex = 1
        
        Do Until lIc = lIo
            
            If lIndex > UBound(sRet) Then
                ReDim Preserve sRet(lIndex + 10)
            End If
            
            sRet(lIndex) = Mid$(sOut, lIo, lIc - lIo)
            lIo = lIc + 1: lPos = lPos + 1: lIndex = lIndex + 1
            lIc = InStr(lIo, sOut, vbNullChar)
            
        Loop
            
        ReDim Preserve sRet(lIndex - 1)
        
        GetOpenFiles = sRet
        
    End If
    
End Function

' // Get file extension
Private Function GetFileExtension( _
                 ByRef sPath As String) As String
    Dim pExt    As Long
    
    pExt = PathFindExtension(StrPtr(sPath))
    GetFileExtension = Mid$(sPath, (pExt - StrPtr(sPath)) \ 2 + 1)
    
End Function

' // Load data from file to byte array
Private Function LoadDataFromFile( _
                 ByRef sFileName As String, _
                 ByRef bData() As Byte) As Boolean
    Dim hFile   As Long
    Dim tSize   As LARGE_INTEGER
    
    hFile = CreateFile(StrPtr(sFileName), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    
    If GetFileSizeEx(hFile, tSize) = 0 Then
        GoTo CleanUp
    End If
    
    If tSize.highPart <> 0 Or tSize.lowPart < 0 Or tSize.lowPart > 100000000 Then
        GoTo CleanUp
    End If
    
    If tSize.lowPart > 0 Then
        
        ReDim bData(tSize.lowPart - 1)
        
        If ReadFile(hFile, bData(0), tSize.lowPart, tSize.highPart, ByVal 0&) = 0 Then
            GoTo CleanUp
        End If
        
        ReDim Preserve bData(tSize.highPart - 1)
        
    End If
    
    LoadDataFromFile = True
    
CleanUp:
    
    If hFile Then
        CloseHandle hFile
    End If
    
    If Not LoadDataFromFile Then
        Erase bData
    End If
    
End Function

' // Get free file name in list based on template.
' // For example if there is a file with the name "test" and you pass "test.bin"
' // it returns "test(1).bin". If the index is already presented it just modifies index in braces.
Private Function GetFreeFileNameFromTemplate( _
                 ByRef sFileName As String) As String
    Dim pExt        As Long:    Dim pPath       As Long
    Dim sExt        As String:  Dim sTitle      As String
    Dim sNumber     As String:  Dim lBr(1)      As Long
    Dim lIndex      As Long:    Dim sNewName    As String
    
    If m_cEmbData.FileExists(sFileName) Then
        
        pPath = StrPtr(sFileName)
        pExt = PathFindExtension(pPath)
        
        sExt = Mid$(sFileName, (pExt - pPath) \ 2 + 1)
        sTitle = Mid$(sFileName, 1, (pExt - pPath) \ 2)
                         
        lBr(1) = InStrRev(sTitle, ")")
        If lBr(1) Then
            lBr(0) = InStrRev(sTitle, "(", lBr(1) - 1)
            If lBr(0) Then
                sNumber = Mid$(sTitle, lBr(0) + 1, lBr(1) - lBr(0) - 1)
                If IsNumeric(sNumber) Then
                    lIndex = CLng(sNumber)
                    sTitle = Left$(sTitle, lBr(0) - 1)
                End If
            End If
        End If
        
        Do
        
            lIndex = lIndex + 1
            sNewName = sTitle & "(" & CStr(lIndex) & ")" & sExt
            
        Loop While m_cEmbData.FileExists(sNewName)
        
        GetFreeFileNameFromTemplate = sNewName
        
    Else
        GetFreeFileNameFromTemplate = sFileName
    End If
    
End Function

' // Preview file
Private Sub cmdActivate_Click()
    Dim lIndex      As Long
    Dim vData       As Variant
    Dim frmPreview  As frmMain
    Dim cCtl        As Control
    Dim cPic        As StdPicture
    Dim bData()     As Byte
    
    lIndex = lstFiles.ListIndex
    If lIndex = -1 Then Exit Sub
    
    Set frmPreview = New frmMain
    
    frmPreview.PreviewMode = True
    
    If IsObject(m_cEmbData.FileData(lstFiles.List(lIndex))) Then
        
        Set vData = m_cEmbData.FileData(lstFiles.List(lIndex))
    
        If TypeOf vData Is StdPicture Then
            frmPreview.PreviewPicture vData
        Else
            ' // You can use your own types to preview
            MsgBox "Unimplemented object. Can't preview", vbCritical
        End If
    Else
    
        ' // Binary
        bData = m_cEmbData.FileData(lstFiles.List(lIndex))
        
        If Not Not bData Then
            If UBound(bData) >= 44 Then
                ' // RIFF WAVE
                If bData(0) = &H52 And bData(1) = &H49 And bData(2) = &H46 And bData(3) = &H46 And _
                   bData(8) = &H57 And bData(9) = &H41 And bData(10) = &H56 And bData(11) = &H45 Then
                    frmPreview.PreviewSound bData()
                    Exit Sub
                End If
            End If
        End If
        
        frmPreview.PreviewRaw bData()
        
    End If

End Sub

Private Sub cmdAddFile_Click()
    Dim sFiles()    As String
    Dim lIndex      As Long
    Dim sFullPath   As String
    Dim cPic        As StdPicture
    Dim bData()     As Byte
    Dim sEmbName    As String
    
    sFiles = GetOpenFiles(Me.hWnd, "Open files", "All files" & vbNullChar & "*.*" & vbNullChar)
    If (Not Not sFiles) = 0 Then Exit Sub
    
    For lIndex = 1 To UBound(sFiles)
    
        sFullPath = sFiles(0) & "\" & sFiles(lIndex)
        sEmbName = GetFreeFileNameFromTemplate(sFiles(lIndex))
        
        ' // Check file extension
        Select Case LCase$(GetFileExtension(sFiles(lIndex)))
        Case ".jpg", ".jpeg", ".bmp", ".gif"
        
            On Error Resume Next

            Err.Clear
            
            Set cPic = LoadPicture(sFullPath)
            
            If Err.Number Then
                MsgBox "Unable to load data from file """ & sFiles(lIndex) & """", vbCritical
                GoTo continue
            End If
            
            On Error GoTo 0
            
            m_cEmbData.Add sEmbName, cPic
            
        Case Else
        
            If Not LoadDataFromFile(sFullPath, bData()) Then
                MsgBox "Unable to load data from file """ & sFiles(lIndex) & """", vbCritical
                GoTo continue
            End If
            
            m_cEmbData.Add sEmbName, bData
            
        End Select
        
        lstFiles.AddItem sEmbName
        
continue:

    Next
    
End Sub

Private Sub cmdRemove_Click()
    Dim lIndex  As Long
    
    lIndex = lstFiles.ListIndex
    If lIndex = -1 Then Exit Sub
    
    If MsgBox("Are you sure?", vbQuestion Or vbYesNo) <> vbYes Then Exit Sub
    
    m_cEmbData.Remove lstFiles.List(lIndex)
    lstFiles.RemoveItem lIndex
    
End Sub

' //
' // Get free file name based on specified one
' //
Private Function GetFreeFileName( _
                 ByRef sFileName As String) As String
    Dim sOut    As String

    sOut = Space$(MAX_PATH)

    If PathYetAnotherMakeUniqueName(StrPtr(sOut), StrPtr(sFileName), 0, StrPtr("temp")) = 0 Then
        Exit Function
    End If
    
    GetFreeFileName = Left$(sOut, InStr(1, sOut, vbNullChar) - 1)
    
End Function

' // Replace EOF data inside itself
Private Sub ReplaceItself( _
            ByRef bData() As Byte, _
            ByVal lOffset As Long)
    Dim sNewFileName    As String
    Dim sCurFileName    As String
    Dim hFile           As Long
    Dim tPI             As PROCESS_INFORMATION
    Dim tSI             As STARTUPINFO
    Dim bIsInIDE        As Boolean
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    If bIsInIDE Then
        MsgBox "You can't write to itself in IDE", vbInformation
        Exit Sub
    End If
    
    ' // Get current file name
    sCurFileName = Space$(MAX_PATH)
    If GetModuleFileName(App.hInstance, StrPtr(sCurFileName), MAX_PATH + 1) = 0 Then
        MsgBox "An error occured", vbCritical
        Exit Sub
    End If
    
    sCurFileName = Left$(sCurFileName, InStr(1, sCurFileName, vbNullChar) - 1)
    
    ' // Get temporary file name
    sNewFileName = GetFreeFileName(sCurFileName)
    
    If Len(sNewFileName) = 0 Then
        MsgBox "Unable to get free file name", vbCritical
        Exit Sub
    End If
    
    ' // Rename me
    If MoveFile(StrPtr(sCurFileName), StrPtr(sNewFileName)) = 0 Then
        MsgBox "Unable to rename executable", vbCritical
        Exit Sub
    End If
    
    ' // Copy the new one with current EXE name
    If CopyFile(StrPtr(sNewFileName), StrPtr(sCurFileName), 1) = 0 Then
        MsgBox "Unable to copy executable", vbCritical
        Exit Sub
    End If
    
    ' // Write EOF data to itself
    hFile = CreateFile(StrPtr(sCurFileName), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, _
                        ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to open executable", vbCritical
        Exit Sub
    End If
    
    SetFilePointer hFile, lOffset, ByVal 0&, FILE_BEGIN
    
    If WriteFile(hFile, bData(0), UBound(bData) + 1, 0, ByVal 0&) = 0 Then
        MsgBox "Unable to write to executable", vbCritical
        CloseHandle hFile
        Exit Sub
    End If
    
    SetEndOfFile hFile
    
    CloseHandle hFile
    
    ' // Run new copy. Close current
    GetStartupInfo tSI
    
    If CreateProcess(0, StrPtr("""" & sCurFileName & """ -del:""" & sNewFileName & _
                     """ -winpos:" & CStr(Int(Me.Left)) & "," & CStr(Int(Me.Top))), ByVal 0&, ByVal 0&, _
                     0, 0, ByVal 0&, ByVal 0&, tSI, tPI) = 0 Then
        MsgBox "Unable to launch executable", vbCritical
        Exit Sub
    End If
    
    CloseHandle tPI.hProcess
    CloseHandle tPI.hThread

    Unload Me
    
End Sub

Private Sub cmdSaveEOF_Click()
    Dim sOutFile        As String
    Dim lOffset         As Long
    Dim bHasEOF         As Boolean
    Dim sMePath         As String
    Dim hFile           As Long
    Dim bSerialized()   As Byte
    Dim dwBinaryType    As Long
    
    bSerialized = m_cEmbData.Serialized()
    
    sOutFile = GetSaveFile(Me.hWnd, "Save EOF data", "PE files" & vbNullChar & "*.exe;*.dll;*.sys;*ocx" & vbNullChar, vbNullString)
    If Len(sOutFile) = 0 Then Exit Sub
    
    If GetBinaryType(StrPtr(sOutFile), dwBinaryType) = 0 Then
        dwBinaryType = -1
    End If
    
    If dwBinaryType = SCS_32BIT_BINARY Then
    
        lOffset = m_cEmbData.GetPEEOFPositionOfFile(sOutFile, bHasEOF)
        If bHasEOF Then
            If MsgBox("The file already has EOF data." & vbNewLine & "Do you want to replace it?", _
                        vbQuestion Or vbYesNo) <> vbYes Then
                Exit Sub
            End If
        End If
        
        sMePath = Space$(MAX_PATH)
        If GetModuleFileName(App.hInstance, StrPtr(sMePath), MAX_PATH + 1) = 0 Then
            MsgBox "An error occured", vbCritical
            Exit Sub
        End If
        
        If lstrcmpi(StrPtr(sOutFile), StrPtr(sMePath)) = 0 Then
            ' // Replace in itself
            ReplaceItself bSerialized, lOffset
        Else
            
            hFile = CreateFile(StrPtr(sOutFile), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, _
                                ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
            If hFile = INVALID_HANDLE_VALUE Then
                MsgBox "Unable to open file", vbCritical
                Exit Sub
            End If
            
            SetFilePointer hFile, lOffset, ByVal 0&, FILE_BEGIN
            
            If WriteFile(hFile, bSerialized(0), UBound(bSerialized) + 1, 0, ByVal 0&) = 0 Then
                MsgBox "Unable to write data", vbCritical
                CloseHandle hFile
                Exit Sub
            End If
            
            CloseHandle hFile
            
        End If
    Else
        
        ' // Make biary file
        hFile = CreateFile(StrPtr(sOutFile), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, _
                            ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        If hFile = INVALID_HANDLE_VALUE Then
            MsgBox "Unable to create file", vbCritical
            Exit Sub
        End If
        
        If WriteFile(hFile, bSerialized(0), UBound(bSerialized) + 1, 0, ByVal 0&) = 0 Then
            MsgBox "Unable to write data", vbCritical
        End If
        
        CloseHandle hFile
        
    End If
    
End Sub

Private Sub ProcessCommandLine()
    Dim sArgs()     As String
    Dim cArgs       As Long
    Dim sFileToDel  As String
    Dim lIndex      As Long
    Dim sCoords(1)  As String
    
    cArgs = ParseCommandLine(sArgs())
    
    For lIndex = 1 To cArgs - 1

        Select Case Left$(sArgs(lIndex), 1)
        
        ' // The command
        Case "-", "/"
        
            ' // -del command: delete a temporary file
            If Mid$(sArgs(lIndex), 2, 4) = "del:" Then
                
                sFileToDel = Mid$(sArgs(lIndex), 6)
                ' // Remove quotes
                PathUnquoteSpaces StrPtr(sFileToDel)
                
                ' // Try to delete file
                Do Until DeleteFile(StrPtr(sFileToDel))
                    Sleep 100
                Loop
            
            ' // Set window pos
            ElseIf Mid$(sArgs(lIndex), 2, 7) = "winpos:" Then
                
                sCoords(0) = Left$(Mid$(sArgs(lIndex), 9), InStr(1, sArgs(lIndex), ",") - 9)
                sCoords(1) = Mid$(sArgs(lIndex), 9 + Len(sCoords(0)) + 1)
                
                If IsNumeric(sCoords(0)) And IsNumeric(sCoords(1)) Then
                    Me.Move Int(sCoords(0)), Int(sCoords(1))
                End If
                
            End If
            
        End Select

    Next
    
End Sub

Private Sub Form_Load()
    Dim lIndex  As Long
    Dim cCtl    As Control
    
    If m_bPreviewMode Then
        
        ' // For preview mode hide all the controls
        For Each cCtl In Me.Controls
            cCtl.Visible = False
        Next
    
        Exit Sub
        
    End If
    
    ProcessCommandLine
    
    Set m_cEmbData = New CEmbeddedFiles
    m_cEmbData.Initialize App.hInstance
    
    For lIndex = 0 To m_cEmbData.FilesCount - 1
        lstFiles.AddItem m_cEmbData.FileName(lIndex)
    Next
    
End Sub

