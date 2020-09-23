Attribute VB_Name = "modLoadDLL"
'--Loads a DLL file into memory, Note lpLibFileName as to be null-terminated to work
 Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
 
'--Returns the Address of the function in the DLL. hModule is the Return Value of LoadLibrary and
'--lpProcName is the Function you are after i.e. GetTickCount again null-terminated
 Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
 
'--Last but not Least FreeUp memory used
 Declare Function FreeLibrary Lib "kernel32" (ByVal hModule As Long) As Long

 Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

 Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Enum e_BinaryData
        DefineByte = 1                          '  8 Bits Data
        DefineWord = 2                          ' 16 Bits Data
        DefineDoubleWord = 4                    ' 32 Bits Data
        DefineQuadWord = 8                      ' 64 Bits Data
End Enum

    ' =============================================================================
    ' Allows Direct Reading from Memory Pointed by MemPointer
    ' with definition of bytes used as in Asm (DB, DW, DD, DX)
    ' =============================================================================
    Function ReadMem(ByVal MemPointer As Long, _
                     SizeInBytes As e_BinaryData)
        Select Case SizeInBytes
            Case DefineByte
                Dim DB As Byte
                CopyMemory DB, ByVal MemPointer, 1
                ReadMem = DB
            Case DefineWord
                Dim DW As Integer
                CopyMemory DW, ByVal MemPointer, 2
                ReadMem = DW
            Case DefineDoubleWord
                Dim DD As Long
                CopyMemory DD, ByVal MemPointer, 4
                ReadMem = DD
            Case DefineQuadWord
                Dim DX As Double
                CopyMemory DX, ByVal MemPointer, 8
                ReadMem = DX
        End Select
    End Function




