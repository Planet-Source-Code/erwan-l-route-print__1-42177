Attribute VB_Name = "mdlapi"
Option Explicit

   Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&


Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type





Public Declare Function FormatMessage Lib "kernel32" _
Alias "FormatMessageA" _
(ByVal dwFlags As Long, _
lpSource As Any, _
ByVal dwMessageId As Long, _
ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, _
ByVal nSize As Long, _
Arguments As Long) _
As Long

   Public Declare Function GetLastError Lib "kernel32" () As Long



Public Const FORMAT_MESSAGE_FROM_SYSTEM = 4096


Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal DllName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hDll As Long, ByVal FuncName As String) As Long

Declare Function LoadLibraryEx _
    Lib "kernel32" _
    Alias "LoadLibraryExA" ( _
    ByVal lpLibFileName As String, _
    ByVal hFile As Long, _
    ByVal dwFlags As Long) As Long

Declare Function FreeLibrary Lib "kernel32" _
       (ByVal hLibModule As Long) As Long
     
    
    Declare Function CopyPointer2String _
    Lib "kernel32" _
    Alias "lstrcpyA" ( _
    ByVal NewString As String, _
    ByVal OldString As Long) As Long
    
    Declare Function PtrToStr Lib "kernel32" _
        Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
    
       
    Declare Function StrToPtr Lib "kernel32" _
        Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long
    
    Declare Function PtrToInt Lib "kernel32" Alias _
        "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, _
        ByVal nCharCount As Long) As Long
    
    Declare Function StrLen Lib "kernel32" Alias _
        "lstrlenW" (ByVal Ptr As Long) As Long
        
    Declare Sub CopyMemory Lib "kernel32" _
        Alias "RtlMoveMemory" (ByRef hpvDest As Any, _
        ByVal hpvSource As Long, ByVal cbCopy As Long)
        
       Public Declare Sub CopyMemory_any Lib "kernel32" _
        Alias "RtlMoveMemory" (ByRef hpvDest As Any, _
          hpvSource As Any, ByVal cbCopy As Long)
          
     Declare Function MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef dest As Any, ByRef src As Any, ByVal Size As Long) As Long
        
    Declare Function lstrlen Lib "kernel32" _
    Alias "lstrlenA" (ByVal lpString As Any) As Long
    
       

    
    
    Declare Function GetVersionEx Lib "kernel32" _
Alias "GetVersionExA" _
(lpVersionInformation As OSVERSIONINFO) As Long




Public Const GMEM_FIXED As Long = 0
Public Const OBJ_PEN As Long = 1
Public Const ERROR_INSUFFICIENT_BUFFER = &H7A
