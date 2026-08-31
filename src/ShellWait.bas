Attribute VB_Name = "ShellWait"
Option Explicit

#If Mac Then
Public Function ExecuteRedirected(ByVal CommandLine As String, _
                                  ByVal StartupDir As String, _
                                  ByVal StdInputPath As String, _
                                  ByVal StdOutputPath As String, _
                                  ByVal StdErrorPath As String, _
                                  Optional ByVal debugMode As Boolean = False, _
                                  Optional ByVal WaitTime As Long = -1) As Long
    Dim RedirectedCommand As String
    RedirectedCommand = CommandLine & " < " & ShellEscape(StdInputPath) & _
        " > " & ShellEscape(StdOutputPath) & " 2> " & ShellEscape(StdErrorPath)
    ExecuteRedirected = Execute(RedirectedCommand, StartupDir, debugMode, WaitTime)
End Function

Public Function Execute(ByVal CommandLine As String, StartupDir As String, Optional debugMode As Boolean = False, Optional WaitTime As Long = -1) As Long
    If debugMode Then
        ShowError vbNullString, CommandLine, "Debug mode", "Next command:", "Continue"
    End If
    Execute = CLng(AppleScriptTask("IguanaTex.scpt", "MacExecute", _
        "export PATH=" & ShellEscape("/usr/local/bin:/opt/homebrew/bin:/Applications/Docker.app/Contents/Resources/bin") & _
        ":""$PATH"" && " & _
        "cd " & ShellEscape(StartupDir) & " && " & _
        CommandLine))
End Function

#Else
' Portions of code below taken from:
' http://www.mvps.org/access/api/api0004.htm
' Courtesy of Terry Kreft

Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const STARTF_USESTDHANDLES As Long = &H100
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const WAIT_TIMEOUT As Long = 258
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const OPEN_EXISTING As Long = 3
Private Const CREATE_ALWAYS As Long = 2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1

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
    dwProcessID As Long
    dwThreadId As Long
End Type

#If VBA7 Then
Private Type SECURITY_ATTRIBUTES_REDIRECT
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Private Type STARTUPINFO_REDIRECT
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
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
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION_REDIRECT
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessID As Long
    dwThreadId As Long
End Type

Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
    
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long
    
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
    
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Public Declare PtrSafe Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, _
  ByVal FileName As String, Optional ByVal Parameters As String, _
  Optional ByVal directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Private Declare PtrSafe Function CreateFileRedirect Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES_REDIRECT, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

Private Declare PtrSafe Function CreateProcessRedirect Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As LongPtr, ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO_REDIRECT, _
    lpProcessInformation As PROCESS_INFORMATION_REDIRECT) As Long

Private Declare PtrSafe Function WaitForRedirectProcess Lib "kernel32" Alias "WaitForSingleObject" ( _
    ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long

Private Declare PtrSafe Function GetRedirectExitCode Lib "kernel32" Alias "GetExitCodeProcess" ( _
    ByVal hProcess As LongPtr, lpExitCode As Long) As Long

Private Declare PtrSafe Function TerminateRedirectProcess Lib "kernel32" Alias "TerminateProcess" ( _
    ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long

Private Declare PtrSafe Function CloseRedirectHandle Lib "kernel32" Alias "CloseHandle" ( _
    ByVal hObject As LongPtr) As Long

#Else
Private Type SECURITY_ATTRIBUTES_REDIRECT
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO_REDIRECT
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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

Private Type PROCESS_INFORMATION_REDIRECT
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadId As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long
    
Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
    
Private Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, _
  ByVal Filename As String, Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Private Declare Function CreateFileRedirect Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES_REDIRECT, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function CreateProcessRedirect Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO_REDIRECT, _
    lpProcessInformation As PROCESS_INFORMATION_REDIRECT) As Long

Private Declare Function WaitForRedirectProcess Lib "kernel32" Alias "WaitForSingleObject" ( _
    ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function GetRedirectExitCode Lib "kernel32" Alias "GetExitCodeProcess" ( _
    ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function TerminateRedirectProcess Lib "kernel32" Alias "TerminateProcess" ( _
    ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CloseRedirectHandle Lib "kernel32" Alias "CloseHandle" ( _
    ByVal hObject As Long) As Long
#End If

Public Function ExecuteRedirected(ByVal CommandLine As String, _
                                  ByVal StartupDir As String, _
                                  ByVal StdInputPath As String, _
                                  ByVal StdOutputPath As String, _
                                  ByVal StdErrorPath As String, _
                                  Optional ByVal debugMode As Boolean = False, _
                                  Optional ByVal WaitTime As Long = -1) As Long
    Dim SecurityAttributes As SECURITY_ATTRIBUTES_REDIRECT
    Dim StartInfo As STARTUPINFO_REDIRECT
    Dim ProcessInfo As PROCESS_INFORMATION_REDIRECT
    #If VBA7 Then
    Dim InputHandle As LongPtr
    Dim OutputHandle As LongPtr
    Dim ErrorHandle As LongPtr
    #Else
    Dim InputHandle As Long
    Dim OutputHandle As Long
    Dim ErrorHandle As Long
    #End If
    Dim Created As Long
    Dim WaitResult As Long
    Dim ExitCode As Long

    SecurityAttributes.nLength = LenB(SecurityAttributes)
    SecurityAttributes.bInheritHandle = 1

    InputHandle = CreateFileRedirect(StdInputPath, GENERIC_READ, FILE_SHARE_READ, _
        SecurityAttributes, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If InputHandle = INVALID_HANDLE_VALUE Then GoTo RedirectError

    OutputHandle = CreateFileRedirect(StdOutputPath, GENERIC_WRITE, FILE_SHARE_READ, _
        SecurityAttributes, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If OutputHandle = INVALID_HANDLE_VALUE Then GoTo RedirectError

    ErrorHandle = CreateFileRedirect(StdErrorPath, GENERIC_WRITE, FILE_SHARE_READ, _
        SecurityAttributes, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If ErrorHandle = INVALID_HANDLE_VALUE Then GoTo RedirectError

    With StartInfo
        .cb = LenB(StartInfo)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = 0
        .hStdInput = InputHandle
        .hStdOutput = OutputHandle
        .hStdError = ErrorHandle
    End With

    If debugMode Then
        ShowError vbNullString, CommandLine, "Debug mode", "Next command:", "Continue"
    End If

    Created = CreateProcessRedirect(0, CommandLine, 0, 0, 1, _
        NORMAL_PRIORITY_CLASS, 0, StartupDir, StartInfo, ProcessInfo)
    If Created = 0 Then GoTo RedirectError

    If WaitTime > 0 Then
        WaitResult = WaitForRedirectProcess(ProcessInfo.hProcess, WaitTime)
    Else
        WaitResult = WaitForRedirectProcess(ProcessInfo.hProcess, INFINITE)
    End If

    If WaitResult = WAIT_TIMEOUT Then
        Call TerminateRedirectProcess(ProcessInfo.hProcess, WAIT_TIMEOUT)
        Call WaitForRedirectProcess(ProcessInfo.hProcess, 5000)
        ExitCode = WAIT_TIMEOUT
    ElseIf GetRedirectExitCode(ProcessInfo.hProcess, ExitCode) = 0 Then
        ExitCode = 1
    End If

    ExecuteRedirected = ExitCode
    GoTo RedirectCleanup

RedirectError:
    ExecuteRedirected = 1

RedirectCleanup:
    If ProcessInfo.hThread <> 0 Then Call CloseRedirectHandle(ProcessInfo.hThread)
    If ProcessInfo.hProcess <> 0 Then Call CloseRedirectHandle(ProcessInfo.hProcess)
    If ErrorHandle <> 0 And ErrorHandle <> INVALID_HANDLE_VALUE Then Call CloseRedirectHandle(ErrorHandle)
    If OutputHandle <> 0 And OutputHandle <> INVALID_HANDLE_VALUE Then Call CloseRedirectHandle(OutputHandle)
    If InputHandle <> 0 And InputHandle <> INVALID_HANDLE_VALUE Then Call CloseRedirectHandle(InputHandle)
End Function

    
Public Function ShellWait(pathname As String, Optional StartupDir As String, Optional WindowStyle As Long, Optional WaitTime As Long = -1) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long
    Dim exitcode As Long
    Dim lastError As Long
    Dim retWait As Long
    
    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        End If
    End With
    Dim sdir As String
    If IsMissing(StartupDir) Then
        sdir = vbNullString
    Else
        sdir = StartupDir
    End If

    ' Start the shelled application:
    ret& = CreateProcessA(0&, pathname, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, sdir, start, proc)
    lastError& = GetLastError()
    If (ret& = 0) Then
        MsgBox "Could not start process: '" & pathname & "'. GetLastError returned " & Str$(lastError&)
        ShellWait = 1
        Exit Function
    End If
        
    ' Wait for the shelled application to finish:
    If WaitTime > 0 Then
        retWait& = WaitForSingleObject(proc.hProcess, WaitTime)
    Else
        retWait& = WaitForSingleObject(proc.hProcess, INFINITE)
    End If
    ' Get return value
    exitcode& = 1234
    ret& = GetExitCodeProcess(proc.hProcess, exitcode&)
    If (ret& = 0) Then
        lastError& = GetLastError()
        MsgBox "GetExitCodeProcess returned " + Str$(ret&) + ", GetLastError returned " + Str$(lastError&)
    End If
    ' Tidy up if time out
    If (retWait& = 258) Then
        ret& = TerminateProcess(proc.hProcess, 0)
    End If
    ' Close handle
    ret& = CloseHandle(proc.hProcess)
    ShellWait = exitcode&
End Function

Public Function Execute(CommandLine As String, StartupDir As String, Optional debugMode As Boolean = False, Optional WaitTime As Long = -1) As Long
    Dim RetVal As Long
    If debugMode Then
        ' Clipboard CommandLine
        ' MsgBox CommandLine, , StartupDir
        ShowError vbNullString, CommandLine, "Debug mode", "Next command:", "Continue"
        RetVal = ShellWait(CommandLine, StartupDir, 1&, WaitTime)
    Else
        RetVal = ShellWait(CommandLine, StartupDir, , WaitTime)
    End If
    Execute = RetVal
End Function
#End If ' Mac


