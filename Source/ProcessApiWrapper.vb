Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class ProcessApiWrapper

    ' --- WIN32 API STRUCTURES ---

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure STARTUPINFO
        Public cb As Integer
        Public lpReserved As String
        Public lpDesktop As String
        Public lpTitle As String
        Public dwX As Integer
        Public dwY As Integer
        Public dwXSize As Integer
        Public dwYSize As Integer
        Public dwXCountChars As Integer
        Public dwYCountChars As Integer
        Public dwFillAttribute As Integer
        Public dwFlags As Integer
        Public wShowWindow As Short
        Public cbReserved2 As Short
        Public lpReserved2 As IntPtr
        Public hStdInput As IntPtr
        Public hStdOutput As IntPtr
        Public hStdError As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PROCESS_INFORMATION
        Public hProcess As IntPtr
        Public hThread As IntPtr
        Public dwProcessId As Integer
        Public dwThreadId As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure SECURITY_ATTRIBUTES
        Public nLength As Integer
        Public lpSecurityDescriptor As IntPtr
        Public bInheritHandle As Boolean
    End Structure

    Public Const STARTF_USESHOWWINDOW As Integer = &H1

    ' --- WIN32 API IMPORTS ---

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function CreateProcess(
        lpApplicationName As String,
        lpCommandLine As String,
        ByRef lpProcessAttributes As SECURITY_ATTRIBUTES,
        ByRef lpThreadAttributes As SECURITY_ATTRIBUTES,
        bInheritHandles As Boolean,
        dwCreationFlags As Integer,
        lpEnvironment As IntPtr,
        lpCurrentDirectory As String,
        ByRef lpStartupInfo As STARTUPINFO,
        ByRef lpProcessInformation As PROCESS_INFORMATION
    ) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function CloseHandle(hObject As IntPtr) As Boolean
    End Function

    Public Shared Function CreateProcessExAPI(
        prName As String,
        prArguments As String,
        Optional prWorkDir As String = Nothing,
        Optional prWindowStyle As Integer = 1,
        Optional inheritHandles As Boolean = False,
        Optional creationFlags As Integer = 0
    ) As Process

        Dim procAttr As New SECURITY_ATTRIBUTES()
        procAttr.nLength = Marshal.SizeOf(procAttr)

        Dim threadAttr As New SECURITY_ATTRIBUTES()
        threadAttr.nLength = Marshal.SizeOf(threadAttr)

        Dim startInfo As New STARTUPINFO()
        startInfo.cb = Marshal.SizeOf(startInfo)
        startInfo.wShowWindow = CType(prWindowStyle, Short)
        startInfo.dwFlags = STARTF_USESHOWWINDOW

        Dim procInfo As New PROCESS_INFORMATION()

        Dim success As Boolean = CreateProcess(
            prName,
            prArguments,
            procAttr,
            threadAttr,
            inheritHandles,
            creationFlags,
            IntPtr.Zero, ' lpEnvironment
            prWorkDir,
            startInfo,
            procInfo
        )

        If Not success Then
            Dim errCode As Integer = Marshal.GetLastWin32Error()
            Throw New Win32Exception(errCode, "Failed to call CreateProcessEx.")
        End If

        CloseHandle(procInfo.hProcess)
        CloseHandle(procInfo.hThread)

        Try
            Dim managedProcess As Process = Process.GetProcessById(procInfo.dwProcessId)
            Return managedProcess
        Catch ex As ArgumentException
            Throw New Exception("Process has been runned but ended after (PID: " & procInfo.dwProcessId & ").")
        End Try

    End Function
End Class