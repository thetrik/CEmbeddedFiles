VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmbeddedFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CEmbeddedFiles.cls
' // Version 1.0
' // This class allows to manage files inside a PE module
' // By The trick, 2020
' //

Option Explicit

Private Const IMAGE_FILE_MACHINE_I386         As Long = &H14C
Private Const IMAGE_DOS_SIGNATURE             As Long = &H5A4D
Private Const IMAGE_NT_SIGNATURE              As Long = &H4550&
Private Const IMAGE_NT_OPTIONAL_HDR32_MAGIC   As Long = &H10B&
Private Const IMAGE_DIRECTORY_ENTRY_SECURITY  As Long = 4
Private Const MAX_PATH                        As Long = 260
Private Const FILE_ATTRIBUTE_NORMAL           As Long = &H80
Private Const INVALID_HANDLE_VALUE            As Long = -1
Private Const GENERIC_READ                    As Long = &H80000000
Private Const OPEN_EXISTING                   As Long = 3
Private Const FILE_SHARE_READ                 As Long = &H1
Private Const E_INVALIDARG                    As Long = &H80070057
Private Const E_UNEXPECTED                    As Long = &H8000FFFF
Private Const FILE_BEGIN                      As Long = 0
Private Const FILE_CURRENT                    As Long = 1
Private Const FILE_END                        As Long = 2
Private Const EOF_SIGNATURE                   As Long = &H9DE463AC
Private Const VT_BYREF                        As Long = &H4000
Private Const LOAD_LIBRARY_AS_DATAFILE        As Long = 2
Private Const FILE_MAP_WRITE                  As Long = 2
Private Const FILE_MAP_READ                   As Long = 4
Private Const PAGE_READWRITE                  As Long = 4&
Private Const PAGE_READONLY                   As Long = 2&
Private Const FILE_SHARE_DELETE               As Long = 4

Private Type UUID
    Data1                       As Long
    Data2                       As Integer
    Data3                       As Integer
    Data4(0 To 7)               As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements                   As Long
    lLBound                     As Long
End Type

Private Type LARGE_INTEGER
    lowPart                     As Long
    highPart                    As Long
End Type

Private Type IMAGE_DOS_HEADER
    e_magic                     As Integer
    e_cblp                      As Integer
    e_cp                        As Integer
    e_crlc                      As Integer
    e_cparhdr                   As Integer
    e_minalloc                  As Integer
    e_maxalloc                  As Integer
    e_ss                        As Integer
    e_sp                        As Integer
    e_csum                      As Integer
    e_ip                        As Integer
    e_cs                        As Integer
    e_lfarlc                    As Integer
    e_ovno                      As Integer
    e_res(0 To 3)               As Integer
    e_oemid                     As Integer
    e_oeminfo                   As Integer
    e_res2(0 To 9)              As Integer
    e_lfanew                    As Long
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress              As Long
    Size                        As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic                       As Integer
    MajorLinkerVersion          As Byte
    MinorLinkerVersion          As Byte
    SizeOfCode                  As Long
    SizeOfInitializedData       As Long
    SizeOfUnitializedData       As Long
    AddressOfEntryPoint         As Long
    BaseOfCode                  As Long
    BaseOfData                  As Long
    ImageBase                   As Long
    SectionAlignment            As Long
    FileAlignment               As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion           As Integer
    MinorImageVersion           As Integer
    MajorSubsystemVersion       As Integer
    MinorSubsystemVersion       As Integer
    W32VersionValue             As Long
    SizeOfImage                 As Long
    SizeOfHeaders               As Long
    CheckSum                    As Long
    SubSystem                   As Integer
    DllCharacteristics          As Integer
    SizeOfStackReserve          As Long
    SizeOfStackCommit           As Long
    SizeOfHeapReserve           As Long
    SizeOfHeapCommit            As Long
    LoaderFlags                 As Long
    NumberOfRvaAndSizes         As Long
    DataDirectory(15)           As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_FILE_HEADER
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type

Private Type IMAGE_NT_HEADERS
    Signature                   As Long
    FileHeader                  As IMAGE_FILE_HEADER
    OptionalHeader              As IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_SECTION_HEADER
    SectionName(1)              As Long
    VirtualSize                 As Long
    VirtualAddress              As Long
    SizeOfRawData               As Long
    PointerToRawData            As Long
    PointerToRelocations        As Long
    PointerToLinenumbers        As Long
    NumberOfRelocations         As Integer
    NumberOfLinenumbers         As Integer
    Characteristics             As Long
End Type

Private Declare Function GetFileSizeEx Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileSize As LARGE_INTEGER) As Long
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
Private Declare Function SetFilePointerEx Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByVal liDistanceToMovel As Long, _
                         ByVal liDistanceToMoveh As Long, _
                         ByRef lpNewFilePointer As Any, _
                         ByVal dwMoveMethod As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByVal lDistanceToMove As Long, _
                         ByRef lpDistanceToMoveHigh As Any, _
                         ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" _
                         Alias "GetModuleFileNameW" ( _
                         ByVal hModule As Long, _
                         ByVal lpFileName As Long, _
                         ByVal nSize As Long) As Long
Private Declare Function SafeArrayCreate Lib "oleaut32" ( _
                         ByVal vt As Long, _
                         ByVal cDims As Long, _
                         ByRef rgsaBound As SAFEARRAYBOUND) As Long
Private Declare Function SafeArrayGetElemsize Lib "oleaut32" ( _
                         ByVal pArr As Long) As Long
Private Declare Function SafeArrayAccessData Lib "oleaut32" ( _
                         ByVal pArr As Long, _
                         ByRef pData As Any) As Long
Private Declare Function SafeArrayUnaccessData Lib "oleaut32" ( _
                         ByVal pArr As Long) As Long
Private Declare Function SafeArrayDestroy Lib "oleaut32" ( _
                         ByVal pArr As Long) As Long
Private Declare Function OleLoadFromStream Lib "ole32" ( _
                         ByVal cStm As IUnknown, _
                         ByRef iidInterface As UUID, _
                         ByRef ppvObj As Any) As Long
Private Declare Function OleSaveToStream Lib "ole32" ( _
                         ByVal ppvObj As Long, _
                         ByVal cStm As IUnknown) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function SHCreateMemStream Lib "shlwapi" _
                         Alias "#12" ( _
                         ByRef pInit As Any, _
                         ByVal cbInit As Long) As IUnknown
Private Declare Function VariantCopy Lib "oleaut32" ( _
                         ByRef pvargDest As Variant, _
                         ByRef pvargSrc As Variant) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" ( _
                         ByVal hGlobal As Long, _
                         ByVal fDeleteOnRelease As Long, _
                         ByRef ppstm As IUnknown) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" ( _
                         ByVal pstm As IUnknown, _
                         ByRef phglobal As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" ( _
                         ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" ( _
                         ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" ( _
                         ByVal hMem As Long) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32" ( _
                         ByRef psa As Any) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" _
                         Alias "LoadLibraryExW" ( _
                         ByVal lpLibFileName As Long, _
                         ByVal hFile As Long, _
                         ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
                         ByVal hLib As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" _
                        Alias "CreateFileMappingW" ( _
                        ByVal hFile As Long, _
                        ByRef lpFileMappingAttributes As Any, _
                        ByVal flProtect As Long, _
                        ByVal dwMaximumSizeHigh As Long, _
                        ByVal dwMaximumSizeLow As Long, _
                        ByVal lpName As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" ( _
                        ByVal hFileMappingObject As Long, _
                        ByVal dwDesiredAccess As Long, _
                        ByVal dwFileOffsetHigh As Long, _
                        ByVal dwFileOffsetLow As Long, _
                        ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" ( _
                        ByVal lpBaseAddress As Long) As Long
Private Declare Function vbaCastObj Lib "msvbvm60" _
                        Alias "__vbaCastObj" ( _
                        ByVal cObj As Any, _
                        ByRef iid As UUID) As Long
Private Declare Function vbaObjSet Lib "msvbvm60" _
                        Alias "__vbaObjSet" ( _
                        ByRef pDst As Any, _
                        ByVal pSrc As Long) As Long
Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
Private Declare Sub memset Lib "kernel32" _
                    Alias "RtlFillMemory" ( _
                    ByRef Destination As Any, _
                    ByVal Length As Long, _
                    ByVal Fill As Byte)

' //
' // The raw structures in file.
' // These structures isn't proper VB structures because they have the variable size.

' // 1. Executable followed by this header
Private Type tEOFDataHeader
    lSignature      As Long         ' // Signature (should be EOF_SIGNATURE)
    lFilesCount     As Long         ' // Number of files (tEOFFileDesc structures)
End Type

' // 2. File descriptor
' // Array of these headers follows tEOFDataHeader
Private Type tEOFFileDesc
    lFileNameSize   As Long         ' // Length of sFileName in characters
    sFileName       As String       ' // File name in UNICODE (without null-terminal)
    vt              As VbVarType    ' // Data type
    lFileSize       As Long         ' // Data size in bytes
    ' // Here are the raw file data
    ' // The following fields aren't presented in the file
    lDataOffset     As Long         ' // The raw data offset in file
    bData()         As Byte         ' // The raw data
End Type

' // End of the structures in file

' // File descriptor for class usage
Private Type tFile
    sName           As String       ' // Name
    bLoaded         As Boolean      ' // Determines if a file already loaded. If so vData contains the valid data
    lFileOffset     As Long         ' // Offset in file where data are located. It's used if bLoaded = false
    lSize           As Long         ' // Size of file
    vt              As VbVarType    ' // Data type
    vData           As Variant      ' // Parsed data
End Type

Private m_tFiles()      As tFile
Private m_lFilesCount   As Long
Private m_hFile         As Long

' //
' // Initialize list based on module base address
' //
Public Sub Initialize( _
           ByVal hInstance As Long)
    Dim lIndex      As Long
    Dim sModName    As String
    Dim hFile       As Long
    Dim tEOFHdr     As tEOFDataHeader
    Dim tEOFFiles() As tEOFFileDesc
    Dim liSize      As LARGE_INTEGER
    Dim lOffset     As Long
    
    Clear
    
    If m_hFile Then
        CloseHandle m_hFile
        m_hFile = 0
    End If
    
    ' // Check EOF position
    lOffset = GetPEEOFPosition(hInstance)
    
    sModName = Space$(MAX_PATH)
    
    lIndex = GetModuleFileName(hInstance, StrPtr(sModName), Len(sModName) + 1)
    If lIndex = 0 Then
        Err.Raise 7
    End If
    
    hFile = CreateFile(StrPtr(sModName), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_DELETE, _
                        ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Err.Raise 5
    End If
    
    If SetFilePointerEx(hFile, 0, 0, liSize, FILE_END) = 0 Then
        CloseHandle hFile
        Err.Raise 5
    End If
    
    ' // Check if the file has the valid size and has the EOF data
    If liSize.highPart <> 0 Or liSize.lowPart < 0 Or liSize.lowPart = lOffset Then
        GoTo CleanUp
    End If
    
    ' // Move to EOF data
    SetFilePointer hFile, lOffset, ByVal 0&, FILE_BEGIN
    
    ' // Read header
    If ReadFile(hFile, tEOFHdr, Len(tEOFHdr), 0, ByVal 0&) = 0 Then
        GoTo CleanUp
    End If
    
    If tEOFHdr.lSignature <> EOF_SIGNATURE Then
        GoTo CleanUp
    End If
    
    If tEOFHdr.lFilesCount <= 0 Then
        GoTo CleanUp
    End If
    
    ReDim tEOFFiles(tEOFHdr.lFilesCount - 1)
    
    For lIndex = 0 To tEOFHdr.lFilesCount - 1
        
        With tEOFFiles(lIndex)
            
            ' // Read file name size
            If ReadFile(hFile, .lFileNameSize, Len(.lFileNameSize), 0, ByVal 0&) = 0 Then
                GoTo CleanUp
            End If
        
            .sFileName = Space$(.lFileNameSize)
            
            ' // Read file name
            If ReadFile(hFile, ByVal StrPtr(.sFileName), .lFileNameSize * 2, 0, ByVal 0&) = 0 Then
                GoTo CleanUp
            End If
            
            ' // Read .vt and .lFileSize
            If ReadFile(hFile, .vt, 8, 0, ByVal 0&) = 0 Then
                GoTo CleanUp
            End If
            
            ' // Save data offset
            .lDataOffset = SetFilePointer(hFile, 0, ByVal 0&, FILE_CURRENT)
            
            ' // Next file
            SetFilePointer hFile, .lFileSize, ByVal 0&, FILE_CURRENT
            
        End With
        
    Next
    
    ' // Update class data
    m_hFile = hFile:    hFile = 0
    m_lFilesCount = tEOFHdr.lFilesCount
    
    ReDim m_tFiles(tEOFHdr.lFilesCount - 1)
    
    For lIndex = 0 To tEOFHdr.lFilesCount - 1
    
        m_tFiles(lIndex).sName = tEOFFiles(lIndex).sFileName
        m_tFiles(lIndex).lSize = tEOFFiles(lIndex).lFileSize
        m_tFiles(lIndex).vt = tEOFFiles(lIndex).vt
        m_tFiles(lIndex).lFileOffset = tEOFFiles(lIndex).lDataOffset
        
    Next

CleanUp:
    
    If hFile Then
        CloseHandle hFile
    End If
    
End Sub

' // Add file to list
Public Sub Add( _
           ByRef sName As String, _
           ByRef vData As Variant)
     
    If FileExists(sName) Then
        Err.Raise 58
    End If
    
    If m_lFilesCount Then
        If m_lFilesCount > UBound(m_tFiles) Then
            ReDim Preserve m_tFiles(m_lFilesCount + 10)
        End If
    Else
        ReDim m_tFiles(9)
    End If
    
    m_tFiles(m_lFilesCount).sName = sName
    m_tFiles(m_lFilesCount).bLoaded = True
    
    If IsObject(vData) Then
        
        Set m_tFiles(m_lFilesCount).vData = vData
        
        ' // We can't use VarType if an object instance has Default property
        ' // Read it from Variant variable directly
        GetMem2 vData, m_tFiles(m_lFilesCount).vt
            
        ' // Clear VT_BYREF flag if exists
        m_tFiles(m_lFilesCount).vt = m_tFiles(m_lFilesCount).vt And (Not VT_BYREF)
                
    Else
        m_tFiles(m_lFilesCount).vData = vData
        m_tFiles(m_lFilesCount).vt = VarType(vData)
    End If
    
    m_lFilesCount = m_lFilesCount + 1
    
End Sub

' // Remove file by name
Public Sub Remove( _
           ByRef sName As String)
    Dim lIndex  As Long
    
    lIndex = GetFileIndex(sName)
    
    If lIndex = -1 Then
        Err.Raise 53
    End If
    
    m_tFiles(lIndex).sName = vbNullString
    m_tFiles(lIndex).vData = Null
    
    If lIndex < m_lFilesCount - 1 Then
        
        ' // Shift array
        memcpy ByVal VarPtr(m_tFiles(lIndex)), ByVal VarPtr(m_tFiles(lIndex + 1)), (m_lFilesCount - lIndex - 1) * LenB(m_tFiles(lIndex))
        
        ' // Zero space
        memset ByVal VarPtr(m_tFiles(m_lFilesCount - 1)), LenB(m_tFiles(lIndex)), 0
        
    End If
    
    m_lFilesCount = m_lFilesCount - 1
               
End Sub

' // Rename file in list
Public Sub RenameFile( _
           ByRef sOldName As String, _
           ByRef sNewName As String)
    Dim lIndex  As Long
    
    lIndex = GetFileIndex(sOldName)
    
    If lIndex = -1 Then
        Err.Raise 76
    End If
    
    If GetFileIndex(sNewName) <> -1 Then
        Err.Raise 58
    End If
    
    m_tFiles(lIndex).sName = sNewName

End Sub

' // Remove all the files
Public Sub Clear()
    Erase m_tFiles
    m_lFilesCount = 0
End Sub

' // Get number of files in list
Public Property Get FilesCount() As Long
    FilesCount = m_lFilesCount
End Property

' // Get file name
Public Property Get FileName( _
                    ByVal lIndex As Long) As String
                        
    If lIndex < 0 Or lIndex >= m_lFilesCount Then
        Err.Raise 9
    End If
    
    FileName = m_tFiles(lIndex).sName
                        
End Property

' // Get file data
Public Property Get FileData( _
                    ByRef sName As String) As Variant
    Dim lIndex  As Long
    
    lIndex = GetFileIndex(sName)
    If lIndex = -1 Then
        Err.Raise 53
    End If
    
    ' // Check if a file is already loaded
    If Not m_tFiles(lIndex).bLoaded Then
    
        ' // To bypass Set issues use API
        VariantCopy m_tFiles(lIndex).vData, LoadData(m_tFiles(lIndex).vt, m_tFiles(lIndex).lFileOffset, m_tFiles(lIndex).lSize)
        m_tFiles(lIndex).bLoaded = True
        
    End If
    
    If IsObject(m_tFiles(lIndex).vData) Then
        Set FileData = m_tFiles(lIndex).vData
    Else
        FileData = m_tFiles(lIndex).vData
    End If

End Property

' // Check if file exists in list
Public Function FileExists( _
                ByRef sName As String) As Boolean
    FileExists = GetFileIndex(sName) <> -1
End Function

' // Get serialized data
' // If you have big data it's reasonable to retrieve the data using a stream instead byte array
Public Property Get Serialized() As Byte()
    Dim tEOFHdr     As tEOFDataHeader
    Dim tEOFFiles() As tEOFFileDesc
    Dim lIndex      As Long
    Dim bFileData() As Byte
    Dim bRet()      As Byte
    Dim lTotalSize  As Long
    Dim lOffset     As Long
    
    tEOFHdr.lSignature = EOF_SIGNATURE
    tEOFHdr.lFilesCount = m_lFilesCount
    
    lTotalSize = Len(tEOFHdr)
    
    If m_lFilesCount Then
        
        ReDim tEOFFiles(m_lFilesCount - 1)
        
        lTotalSize = lTotalSize + m_lFilesCount * 12
        
        For lIndex = 0 To m_lFilesCount - 1
            
            With tEOFFiles(lIndex)
                
                .sFileName = m_tFiles(lIndex).sName
                .lFileNameSize = Len(.sFileName)
                .vt = m_tFiles(lIndex).vt
                
                lTotalSize = lTotalSize + .lFileNameSize * 2
                
                ' // If file already loaded save it
                If m_tFiles(lIndex).bLoaded Then
                
                    .bData = SaveData(m_tFiles(lIndex).vData)
                    
                    If Not Not .bData Then
                        m_tFiles(lIndex).lSize = UBound(.bData) + 1
                    Else
                        m_tFiles(lIndex).lSize = 0
                    End If
                    
                Else
                    
                    ' // Just use the old raw data
                    If m_tFiles(lIndex).lSize Then
                    
                        ReDim .bData(m_tFiles(lIndex).lSize - 1)
                    
                        SetFilePointer m_hFile, m_tFiles(lIndex).lFileOffset, ByVal 0&, FILE_BEGIN
                    
                        If ReadFile(m_hFile, .bData(0), m_tFiles(lIndex).lSize, 0, ByVal 0&) = 0 Then
                            Err.Raise 75
                        End If
                    
                    End If
                    
                End If
                                
                .lFileSize = m_tFiles(lIndex).lSize
                
                lTotalSize = lTotalSize + .lFileSize
                
            End With
            
        Next
        
    End If
    
    ReDim bRet(lTotalSize - 1)
    
    memcpy bRet(0), tEOFHdr, Len(tEOFHdr): lOffset = Len(tEOFHdr)
    
    For lIndex = 0 To m_lFilesCount - 1
        
        With tEOFFiles(lIndex)
            
            memcpy bRet(lOffset), .lFileNameSize, 4:    lOffset = lOffset + 4
            memcpy bRet(lOffset), ByVal StrPtr(.sFileName), .lFileNameSize * 2: lOffset = lOffset + .lFileNameSize * 2
            memcpy bRet(lOffset), .vt, 8:    lOffset = lOffset + 8
            
            If .lFileSize > 0 Then
                memcpy bRet(lOffset), .bData(0), .lFileSize:    lOffset = lOffset + .lFileSize
            End If
            
        End With
        
    Next
    
    Serialized = bRet
    
End Property

' // Get EOF position in specified PE file
' // bHasEOFData receives True if the file already have unknown EOF data
Public Function GetPEEOFPositionOfFile( _
                ByRef sFileName As String, _
                ByRef bHasEOFData As Boolean) As Long
    Dim hFile   As Long
    Dim hMap    As Long
    Dim pData   As Long
    Dim lErr    As Long
    Dim lSize   As Long
    Dim lOffset As Long
    Dim lSign   As Long
    
    On Error GoTo CleanUp
    
    ' // Map file as data to memory
    hFile = CreateFile(StrPtr(sFileName), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Err.Raise 5
    End If
    
    hMap = CreateFileMapping(hFile, ByVal 0&, PAGE_READONLY, 0, 0, 0)
    If hMap = 0 Then
        Err.Raise 7
    End If

    pData = MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, 0)
    If pData = 0 Then
        Err.Raise 7
    End If
    
    ' // Get offset like in loaded module
    lOffset = GetPEEOFPosition(pData)
    
    lSize = SetFilePointer(hFile, 0, ByVal 0&, FILE_END)
    
    If lSize > lOffset Then
        If lSize > lOffset + 4 Then
            memcpy lSign, ByVal pData + lOffset, 4
            If lSign <> EOF_SIGNATURE Then
                bHasEOFData = True
            End If
        End If
    End If
    
    GetPEEOFPositionOfFile = lOffset
    
CleanUp:
    
    lErr = Err.Number
    
    If pData Then
        UnmapViewOfFile pData
    End If
    
    If hMap Then
        CloseHandle hMap
    End If
    
    If hFile Then
        CloseHandle hFile
    End If
    
    If lErr Then
        Err.Raise lErr
    End If
    
End Function

' // Get EOF position based on file in-memory
Private Function GetPEEOFPosition( _
                 ByVal hModule As Long) As Long
    Dim tDosHdr     As IMAGE_DOS_HEADER
    Dim tNtHdr      As IMAGE_NT_HEADERS
    Dim lOffset     As Long
    Dim tSections() As IMAGE_SECTION_HEADER
    Dim lIndex      As Long
    
    memcpy tDosHdr, ByVal hModule, Len(tDosHdr)

    If tDosHdr.e_magic <> IMAGE_DOS_SIGNATURE Then
        Err.Raise 5
    End If
    
    memcpy tNtHdr, ByVal hModule + tDosHdr.e_lfanew, Len(tNtHdr)
    
    If tNtHdr.Signature <> IMAGE_NT_SIGNATURE Or _
        tNtHdr.OptionalHeader.Magic <> IMAGE_NT_OPTIONAL_HDR32_MAGIC Or _
        tNtHdr.FileHeader.Machine <> IMAGE_FILE_MACHINE_I386 Or _
        tNtHdr.FileHeader.NumberOfSections <= 0 Or _
        tNtHdr.FileHeader.SizeOfOptionalHeader <> Len(tNtHdr.OptionalHeader) Then
        Err.Raise 5
    End If
    
    ' // Security directory isn't mapped to memory so we should check if it's placed behind the image
    If tNtHdr.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_SECURITY).Size Then
        lOffset = tNtHdr.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_SECURITY).VirtualAddress + _
                  tNtHdr.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_SECURITY).Size
    End If
    
    ReDim tSections(tNtHdr.FileHeader.NumberOfSections - 1)
    
    memcpy tSections(0), ByVal hModule + tDosHdr.e_lfanew + Len(tNtHdr), Len(tSections(0)) * tNtHdr.FileHeader.NumberOfSections
    
    ' // Scan all the section and check maximum offset
    For lIndex = 0 To tNtHdr.FileHeader.NumberOfSections - 1
        If tSections(lIndex).SizeOfRawData > 0 Then
            If tSections(lIndex).PointerToRawData > lOffset Then
                lOffset = tSections(lIndex).PointerToRawData + tSections(lIndex).SizeOfRawData
            End If
        End If
    Next
    
    GetPEEOFPosition = lOffset
    
End Function

' // Serialize a variant to byte array
Private Function SaveData( _
                 ByRef vData As Variant) As Byte()
    Dim bRet()  As Byte
    Dim vt      As Long
    
    ' // Check if array
    If VarType(vData) And vbArray Then
        SaveData = SaveArray(vData)
        Exit Function
    Else
        
        ' // Get var type
        GetMem2 vData, vt:  vt = vt And (Not VT_BYREF)
    
        Select Case vt
        Case vbNull
        Case vbEmpty
        Case vbInteger
            ReDim bRet(1)
            GetMem2 CInt(vData), bRet(0)
        Case vbLong
            ReDim bRet(3)
            GetMem4 CLng(vData), bRet(0)
        Case vbSingle
            ReDim bRet(3)
            GetMem4 CSng(vData), bRet(0)
        Case vbDouble
            ReDim bRet(7)
            GetMem8 CDbl(vData), bRet(0)
        Case vbCurrency
            ReDim bRet(7)
            GetMem8 CCur(vData), bRet(0)
        Case vbDate
            ReDim bRet(7)
            GetMem8 CDate(vData), bRet(0)
        Case vbString
            ' // The string data is placed as length-prefixes one
            ReDim bRet(4 + Len(vData) * 2 - 1)
            GetMem4 Len(vData), bRet(0)
            memcpy bRet(4), ByVal StrPtr(vData), Len(vData) * 2
        Case vbObject, vbDataObject
            bRet = SaveObject(vData)
        Case vbError
            ReDim bRet(3)
            GetMem4 ByVal VarPtr(vData) + 8, bRet(0)
        Case vbBoolean
            ReDim bRet(1)
            GetMem2 CBool(vData), bRet(0)
        Case vbVariant
            Err.Raise E_UNEXPECTED
        Case vbDecimal
            ' // Decimal contains all 16 bytes
            ReDim bRet(15)
            memcpy bRet(0), vData, 16
        Case vbByte
            ReDim bRet(0)
            bRet(0) = vData
        End Select
    End If
    
    SaveData = bRet
    
End Function

' // Sarialize an array to byte array
Private Function SaveArray( _
                 ByRef vData As Variant) As Byte()
    Dim iDimensions As Integer
    Dim vt          As Integer
    Dim pArray      As Long
    Dim bRet()      As Byte
    Dim lIndex      As Long
    Dim lElemCount  As Long
    Dim lElemSize   As Long
    Dim tBounds()   As SAFEARRAYBOUND
    Dim lTotalSize  As Long
    Dim pData       As Long
    
    ' // Layout
    ' // cDim : [tBounds ....] : [raw bytes]
    
    GetMem2 vData, vt
    
    If vt And VT_BYREF Then
        GetMem4 ByVal VarPtr(vData) + 8, pArray
        GetMem4 ByVal pArray, pArray
    Else
        GetMem4 ByVal VarPtr(vData) + 8, pArray
    End If
    
    lTotalSize = 2
    
    iDimensions = SafeArrayGetDim(ByVal pArray)
    
    If iDimensions > 0 Then
        
        ReDim tBounds(iDimensions - 1)
        
        memcpy tBounds(0), ByVal pArray + &H10, Len(tBounds(0)) * iDimensions
        lElemCount = 1
        
        For lIndex = 0 To iDimensions - 1
            lElemCount = lElemCount * tBounds(lIndex).cElements
        Next
        
        lElemSize = SafeArrayGetElemsize(pArray)
        
        lTotalSize = lTotalSize + Len(tBounds(0)) * iDimensions + lElemSize * lElemCount
        
        ReDim bRet(lTotalSize - 1)
        
        SafeArrayAccessData pArray, pData
        
        memcpy bRet(2), ByVal pArray + &H10, Len(tBounds(0)) * iDimensions
        memcpy bRet(2 + Len(tBounds(0)) * iDimensions), ByVal pData, lElemCount * lElemSize
        
        SafeArrayUnaccessData pArray
        
    Else
        ReDim bRet(lTotalSize - 1)
    End If
    
    GetMem2 iDimensions, bRet(0)
    
    SaveArray = bRet
    
End Function

' // Serialize an object to byte array
Private Function SaveObject( _
                 ByRef vData As Variant) As Byte()
    Dim hr      As Long
    Dim pPerst  As Long
    Dim cObj    As IUnknown
    Dim cStm    As IUnknown
    Dim hGlobal As Long
    Dim lSize   As Long
    Dim pMemory As Long
    Dim bRet()  As Byte
    
    hr = CreateStreamOnHGlobal(0, 1, cStm)
    If hr < 0 Then
        Err.Raise hr
    End If

    Set cObj = vData
    pPerst = vbaCastObj(cObj, IID_IPersistStream)

    hr = OleSaveToStream(pPerst, cStm)
    
    vbaObjSet pPerst, 0
    
    If hr < 0 Then
        Err.Raise hr
    End If
    
    hr = GetHGlobalFromStream(cStm, hGlobal)
    If hr < 0 Then
        Err.Raise hr
    End If
    
    lSize = GlobalSize(hGlobal)
    If lSize = 0 Then
        Err.Raise 7
    End If
    
    ReDim bRet(lSize - 1)
    
    pMemory = GlobalLock(hGlobal)
    If pMemory = 0 Then
        Err.Raise 7
    End If
    
    memcpy bRet(0), ByVal pMemory, lSize
    
    GlobalUnlock hGlobal
    
    SaveObject = bRet
    
End Function

' // Load data from file
Private Function LoadData( _
                 ByVal vt As Long, _
                 ByVal lDataOffset As Long, _
                 ByVal lSize As Long) As Variant
    Dim vValue      As Variant:     Dim iValue      As Integer
    Dim bValue      As Byte:        Dim lValue      As Long
    Dim fValue      As Single:      Dim dValue      As Single
    Dim curValue    As Currency:    Dim datValue    As Date
    Dim sValue      As String:      Dim cStm        As IUnknown
    Dim bBuffer()   As Byte
    
    If m_hFile = 0 Then
        Err.Raise E_UNEXPECTED
    End If
    
    If SetFilePointerEx(m_hFile, lDataOffset, 0, ByVal 0&, FILE_BEGIN) = 0 Then
        Err.Raise 7
    End If
    
    If vt And vbArray Then
        LoadData = LoadArray(vt And (Not vbArray), lDataOffset, lSize)
    Else
        Select Case vt
        Case vbNull
            LoadData = Null
        Case vbEmpty
            LoadData = Empty
        Case vbInteger
            If ReadFile(m_hFile, iValue, 2, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = iValue
        Case vbLong
            If ReadFile(m_hFile, lValue, 4, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = lValue
        Case vbSingle
            If ReadFile(m_hFile, fValue, 4, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = fValue
        Case vbDouble
            If ReadFile(m_hFile, dValue, 8, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = dValue
        Case vbCurrency
            If ReadFile(m_hFile, curValue, 8, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = curValue
        Case vbDate
            If ReadFile(m_hFile, datValue, 8, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = datValue
        Case vbString
        
            If ReadFile(m_hFile, lValue, 4, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            
            If lValue > 0 Then
                sValue = Space$(lValue)
                
                If ReadFile(m_hFile, ByVal StrPtr(sValue), lValue * 2, 0, ByVal 0&) = 0 Then
                    Err.Raise 7
                End If
            
                LoadData = sValue
            Else
                LoadData = vbNullString
            End If
        Case vbObject, vbDataObject
        
            ReDim bBuffer(lSize - 1)
            
            If ReadFile(m_hFile, bBuffer(0), lSize, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            
            Set cStm = SHCreateMemStream(bBuffer(0), lSize)
            
            Set LoadData = LoadObject(vt, cStm)
            
        Case vbError
            If ReadFile(m_hFile, lValue, 4, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = CVErr(lValue)
        Case vbBoolean
            If ReadFile(m_hFile, iValue, 2, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = CBool(iValue)
        Case vbVariant
            Err.Raise E_UNEXPECTED
        Case vbDecimal
            LoadData = Null
            If ReadFile(m_hFile, LoadData, 16, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
        Case vbByte
            If ReadFile(m_hFile, bValue, 1, 0, ByVal 0&) = 0 Then
                Err.Raise 7
            End If
            LoadData = bValue
        End Select
        
    End If

End Function

' // Load an object from memory
Private Function LoadObject( _
                 ByVal vt As Long, _
                 ByVal cStm As IUnknown) As Variant
    Dim hr      As Long
    Dim cObj    As IUnknown
    Dim cDisp   As Object
    
    ' // ReadClassStm / CoCreateInstance / IPersistStream::Load
    hr = OleLoadFromStream(cStm, IID_IUnknown, cObj)
    If hr < 0 Then
        Err.Raise hr
    End If
    
    If vt = vbObject Then
        Set cDisp = cObj
        Set LoadObject = cDisp
    Else
        Set LoadObject = cObj
    End If
    
End Function

Private Function LoadArray( _
                 ByVal vt As Long, _
                 ByVal lDataOffset As Long, _
                 ByVal lSize As Long) As Variant
    Dim iDimensions As Integer
    Dim tBounds()   As SAFEARRAYBOUND
    Dim lElemCount  As Long
    Dim lIndex      As Long
    Dim pArray      As Long
    Dim lTotalSize  As Long
    Dim pArrData    As Long
    
    If SetFilePointerEx(m_hFile, lDataOffset, 0, ByVal 0&, FILE_BEGIN) = 0 Then
        Err.Raise 7
    End If
    
    If ReadFile(m_hFile, iDimensions, 2, 0, ByVal 0&) = 0 Then
        Err.Raise 7
    End If

    If iDimensions > 0 Then
    
        ReDim tBounds(iDimensions - 1)
        
        If ReadFile(m_hFile, tBounds(0), iDimensions * Len(tBounds(0)), 0, ByVal 0&) = 0 Then
            Err.Raise 7
        End If

        lElemCount = 1
        
        For lIndex = 0 To iDimensions - 1
            lElemCount = lElemCount * tBounds(lIndex).cElements
        Next
        
        pArray = SafeArrayCreate(vt, iDimensions, tBounds(0))
        lTotalSize = SafeArrayGetElemsize(pArray) * lElemCount
        
        SafeArrayAccessData pArray, pArrData
        
        If ReadFile(m_hFile, ByVal pArrData, lTotalSize, 0, ByVal 0&) = 0 Then
            SafeArrayUnaccessData pArray
            SafeArrayDestroy pArray
            Err.Raise 7
        End If

        SafeArrayUnaccessData pArray
    
    Else
        pArray = 0
    End If

    GetMem2 vt Or vbArray, LoadArray
    GetMem4 pArray, ByVal VarPtr(LoadArray) + 8
                
End Function

Private Function IID_IUnknown() As UUID
    IID_IUnknown.Data4(0) = &HC0
    IID_IUnknown.Data4(7) = &H46
End Function

Private Function IID_IPersistStream() As UUID
    IID_IPersistStream.Data1 = &H109
    IID_IPersistStream.Data4(0) = &HC0
    IID_IPersistStream.Data4(7) = &H46
End Function

Private Function GetFileIndex( _
                 ByRef sName As String) As Long
    Dim lIndex  As Long
    
    GetFileIndex = -1
    
    For lIndex = 0 To FilesCount - 1
        If StrComp(m_tFiles(lIndex).sName, sName, vbTextCompare) = 0 Then
            GetFileIndex = lIndex
            Exit For
        End If
    Next
    
End Function

Private Sub Class_Terminate()
    If m_hFile Then
        CloseHandle m_hFile
    End If
End Sub
