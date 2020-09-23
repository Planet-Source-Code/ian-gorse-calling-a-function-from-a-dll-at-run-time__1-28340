VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dynamically Load DLL"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFunction 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "GetTickCount"
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtDLLName 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "kernel32.dll"
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Function"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "DLL Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--Loads a DLL file into memory, Note lpLibFileName as to be null-terminated to work
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
 
'--Returns the Address of the function in the DLL. hModule is the Return Value of LoadLibrary and
'--lpProcName is the Function you are after i.e. GetTickCount again null-terminated
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
 
'--Last but not Least FreeUp memory used
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hModule As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

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




Private Sub cmdLoad_Click()
    Dim lngRetLib As Long
    Dim lngRetFree As Long
    Dim lngRetProc As Long
    
    '--Loads the Dll
    lngRetLib = LoadLibrary(txtDLLName.Text & Chr$(0))
    
    '--Get Memory address of function
    lngRetProc = GetProcAddress(lngRetLib, txtFunction.Text & Chr$(0))
    
   
    '--We Will call the GetTickCount API call based from its memory location and compare it with declared api call( should be same )
    '--By the Way this is guessing you are checking the GetTickCount ( default textbox )
    MsgBox "GetTickCount from vb Declared :" & GetTickCount() & vbCrLf & _
           "GetTickCount from API address :" & CallWindowProc(lngRetProc, 0, 0, 0, 0)
    
'####Just some debug stuff###'
    '--Function from Chris Vega - Accessing Memory by 32-bit Addresing in Windows using Visual Basic
    'MsgBox Hex(ReadMem(lngRetProc, DefineDoubleWord))
'####End Debug Stuff###'
    
    '--Free Up memory
    lngRetFree = FreeLibrary(lngRetLib)
    
End Sub

