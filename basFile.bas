Attribute VB_Name = "basFile"
Option Explicit

Private Const MAX_FILENAME_LEN = 256

' File and Disk functions.
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_UNKNOWN = 0    'Unknown, or unable to be determined.


Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)

Private Declare Function GetWindowsDirectoryA Lib "kernel32" _
   (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   
Private Declare Function GetTempPathA Lib "kernel32" _
   (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const UNIQUE_NAME = &H0

Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal _
   lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique _
   As Long, ByVal lpTempFileName As String) As Long
   
Private Declare Function GetSystemDirectoryA Lib "kernel32" _
   (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Declare Function ShellExecute Lib _
   "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpdirectory As String, _
   ByVal nShowCmd As Long) As Long
   
Private Const SW_HIDE = 0             ' = vbHide
Private Const SW_SHOWNORMAL = 1       ' = vbNormal
Private Const SW_SHOWMINIMIZED = 2    ' = vbMinimizeFocus
Private Const SW_SHOWMAXIMIZED = 3    ' = vbMaximizedFocus
Private Const SW_SHOWNOACTIVATE = 4   ' = vbNormalNoFocus
Private Const SW_MINIMIZE = 6         ' = vbMinimizedNofocus

Private Declare Function GetShortPathNameA Lib "kernel32" _
   (ByVal lpszLongPath As String, ByVal lpszShortPath _
   As String, ByVal cchBuffer As Long) As Long
   
Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAborted As Boolean
        hNameMaps As Long
        sProgress As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_SILENT = &H4
Private Const FOF_NOCONFIRMATION = &H10

Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
   "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

Private Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CreateProcessA Lib "kernel32" _
    (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function FindExecutableA Lib "shell32.dll" _
   (ByVal lpFile As String, ByVal lpdirectory As _
   String, ByVal lpResult As String) As Long

Private Declare Function SetVolumeLabelA Lib "kernel32" _
   (ByVal lpRootPathName As String, _
   ByVal lpVolumeName As String) As Long

'
'  Finds the executable associated with a file
'
'  Returns "" if no file is found.
'
Public Function AddBackslash(s As String) As String
   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If
End Function

Public Function GetSystemDirectory() As String
   Dim s As String
   Dim i As Integer
   i = GetSystemDirectoryA("", 0)
   s = Space(i)
   Call GetSystemDirectoryA(s, i)
   GetSystemDirectory = AddBackslash(Left$(s, i - 1))
End Function

