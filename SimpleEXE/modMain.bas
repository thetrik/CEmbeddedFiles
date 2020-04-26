Attribute VB_Name = "modMain"
' //
' // Test CEmbeddedFiles class usage
' //

Option Explicit

Sub Main()
    Dim cEmb    As CEmbeddedFiles
    Dim lIndex  As Long
    Dim vData   As Variant
    Dim bData() As Byte
    Dim lHex    As Long
    Dim sHex    As String
    
    Set cEmb = New CEmbeddedFiles
    
    cEmb.Initialize App.hInstance
    
    Open "dump.txt" For Output As #1
    
    For lIndex = 0 To cEmb.FilesCount - 1
    
        Print #1, "File: "; lIndex
        Print #1, "Name: "; cEmb.FileName(lIndex)
        
        If IsObject(cEmb.FileData(cEmb.FileName(lIndex))) Then
            
            Print #1, "Type: Object"
            
        Else
        
            bData = cEmb.FileData(cEmb.FileName(lIndex))
            Print #1, "Type: Binary"
            
            If Not Not bData Then
            
                Print #1, "Size: "; UBound(bData) + 1
                
                Print #1, "First 16 bytes:"
                
                For lHex = 0 To IIf(UBound(bData) > 15, 15, UBound(bData))
                    
                    sHex = Hex$(bData(lHex))
                    
                    If Len(sHex) = 1 Then sHex = "0" & sHex
                    
                    Print #1, sHex; " ";
                    
                Next
                
                Print #1, ""
                
            Else
                Print #1, "Size: 0"
            End If
            
        End If

    Next
    
End Sub
