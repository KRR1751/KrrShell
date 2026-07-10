Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text

<ComVisible(False)>
Public Class ShellApi

    Public Const TPM_RIGHTBUTTON As Integer = &H2
    Public Const TPM_LEFTBUTTON As Integer = &H0

    Public Const CMIC_MASK_UNICODE As Integer = &H4000
    Public Const CMIC_MASK_ASYNCOK As Integer = &H100000

    Public Const TPM_NONOTIFY As Integer = &H80
    Public Const TPM_RETURNCMD As Integer = &H100

    Public Const GCS_VERBA As UInteger = &H0 ' Pro ANSI názvy (open, edit...)
    Public Const GCS_VERBW As UInteger = &H4 ' Pro Unicode názvy

    ' Funkce
    <DllImport("user32.dll")>
    Public Shared Function GetDesktopWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function ExtractIconEx(
    ByVal lpszFile As String,
    ByVal nIconIndex As Integer,
    ByVal phiconLarge() As IntPtr,
    ByVal phiconSmall() As IntPtr,
    ByVal nIcons As UInteger
) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function

    Public Shared Function GetIcon(ByVal filePath As String, ByVal iconIndex As Integer, ByVal isSmallIcon As Boolean) As Icon
        Dim hIconLarge(0) As IntPtr
        Dim hIconSmall(0) As IntPtr
        Dim iconToReturn As Icon = Nothing

        Try
            Dim result As UInteger = ExtractIconEx(filePath, iconIndex, hIconLarge, hIconSmall, 1)

            If result > 0 Then
                Dim hIcon As IntPtr = If(isSmallIcon, hIconSmall(0), hIconLarge(0))

                If hIcon <> IntPtr.Zero Then
                    iconToReturn = Icon.FromHandle(hIcon).Clone()

                    DestroyIcon(hIcon)
                End If
            End If

        Catch ex As Exception
            Return Nothing
        End Try

        Return iconToReturn
    End Function

#Region "Structures"

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure CMINVOKECOMMANDINFO
        Public cbSize As Integer
        Public fMask As CMIC
        Public hwnd As IntPtr
        Public lpVerb As IntPtr ' For string or MenuItem id (MAKEINTRESOURCE)
        Public lpParameters As String
        Public lpDirectory As String
        Public nShow As SHOW_WINDOW ' <-- ZMĚNA ZDE
        Public dwHotKey As Integer
        Public hIcon As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure CMINVOKECOMMANDINFOEX
        Public cbSize As Integer
        Public fMask As CMIC
        Public hwnd As IntPtr
        Public lpVerb As IntPtr
        Public lpParameters As String
        Public lpDirectory As String
        Public nShow As SHOW_WINDOW ' <-- ZMĚNA ZDE
        Public dwHotKey As Integer
        Public hIcon As IntPtr
        Public dwHotKey2 As Integer
        Public lpVerbW As IntPtr ' Unicode verb for LPVerb
        Public lpParametersW As String
        Public lpDirectoryW As String
        Public lpTitle As String
        Public lpReserved As IntPtr
        Public dwReserved As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure POINT
        Public x As Integer
        Public y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure LVITEM
        Public mask As UInteger
        Public iItem As Integer
        Public iSubItem As Integer
        Public state As UInteger
        Public stateMask As UInteger
        <MarshalAs(UnmanagedType.LPTStr)> Public pszText As String
        Public cchTextMax As Integer
        Public iImage As Integer
        Public lParam As IntPtr
        Public iIndent As Integer
        Public iGroupId As Integer
        Public cColumns As UInteger
        Public puColumns As IntPtr
    End Structure

#End Region

#Region "Enums and Constants"

    Public Enum CMIC As UInteger
        CMIC_HOTKEY = &H20
        CMIC_NOASYNC = &H100
        CMIC_NOERRORUI = &H400
        CMIC_NO_CONSOLE_REDIRECTION = &H8000
        CMIC_PTINVOKE = &H20000000
        CMIC_TB_ITEM = &H80000
        CMIC_UNICODE = &H4000
        CMIC_WAITFORINPUTIDLE = &H1000000
        CMIC_NOUI = &H4
        CMIC_INPUTRELEASED = &H10000000
        CMIC_ASYNC = &H100
    End Enum

    Public Enum SHOW_WINDOW As Integer
        SW_HIDE = 0
        SW_SHOWNORMAL = 1
        SW_NORMAL = 1
        SW_SHOWMINIMIZED = 2
        SW_SHOWMAXIMIZED = 3
        SW_MAXIMIZE = 3
        SW_SHOWNOACTIVATE = 4
        SW_SHOW = 5
        SW_MINIMIZE = 6
        SW_SHOWMINNOACTIVE = 7
        SW_SHOWNA = 8
        SW_RESTORE = 9
        SW_SHOWDEFAULT = 10
        SW_FORCEMINIMIZE = 11
    End Enum

    ' Defined by shell32.dll
    Public Const CMF_NORMAL As UInteger = &H0
    Public Const CMF_DEFAULTONLY As UInteger = &H1
    Public Const CMF_VERBSONLY As UInteger = &H2
    Public Const CMF_EXPLORE As UInteger = &H4
    Public Const CMF_NOVERBS As UInteger = &H8
    Public Const CMF_CANRENAME As UInteger = &H10
    Public Const CMF_NODEFAULT As UInteger = &H20
    Public Const CMF_INCLUDESTATIC As UInteger = &H40
    Public Const CMF_EXTENDEDVERBS As UInteger = &H100 ' New in XP
    Public Const CMF_DISABLEDVERBS As UInteger = &H200 ' New in Vista
    Public Const CMF_ASYNCVERBSTATE As UInteger = &H400 ' New in Vista
    Public Const CMF_OPTIMIZEFORINVOKE As UInteger = &H800 ' New in Vista
    Public Const CMF_SYNCCASCADEMENU As UInteger = &H1000 ' New in Vista
    Public Const CMF_DONTADDTODROPDOWN As UInteger = &H2000 ' New in Vista
    Public Const CMF_ITEMMENU As ULong = &H80000000UI ' <-- ZMĚNA ZDE: ULong a UInteger suffix


    ' Define the HRESULT codes
    Public Const S_OK As Integer = &H0
    Public Const S_FALSE As Integer = &H1
    Public Const E_FAIL As Integer = &H80004005
    Public Const E_INVALIDARG As Integer = &H80070057
    Public Const E_OUTOFMEMORY As Integer = &H8007000E

#End Region

#Region "COM Interfaces"

    ' IUnknown Interface (Base for all COM interfaces)
    <ComImport(), Guid("00000000-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IUnknown
        <PreserveSig()> Function QueryInterface(ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer
        <PreserveSig()> Function AddRef() As UInteger
        <PreserveSig()> Function Release() As UInteger
    End Interface

    ' IEnumIDList Interface
    <ComImport(), Guid("000214F2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IEnumIDList
        <PreserveSig()> Function [Next](ByVal celt As UInteger, ByRef rgelt As IntPtr, ByRef pceltFetched As UInteger) As Integer
        <PreserveSig()> Function Skip(ByVal celt As UInteger) As Integer
        <PreserveSig()> Function Reset() As Integer
        <PreserveSig()> Function Clone(ByRef ppenum As IEnumIDList) As Integer
    End Interface

    ' IShellFolder Interface
    ' Represents a shell folder object
    <ComImport(), Guid("000214E6-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IShellFolder
        <PreserveSig()> Function ParseDisplayName(ByVal hwnd As IntPtr, ByVal pbc As IntPtr, ByVal pszDisplayName As String, ByRef pchEaten As UInteger, ByRef ppidl As IntPtr, ByRef pdwAttributes As UInteger) As Integer
        <PreserveSig()> Function EnumObjects(ByVal hwnd As IntPtr, ByVal grfFlags As UInteger, ByRef ppenumIDList As IEnumIDList) As Integer
        <PreserveSig()> Function BindToObject(ByVal pidl As IntPtr, ByVal pbc As IntPtr, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
        <PreserveSig()> Function BindToStorage(ByVal pidl As IntPtr, ByVal pbc As IntPtr, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
        <PreserveSig()> Function CompareIDs(ByVal lParam As IntPtr, ByVal pidl1 As IntPtr, ByVal pidl2 As IntPtr) As Integer
        <PreserveSig()> Function CreateViewObject(ByVal hwndOwner As IntPtr, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
        <PreserveSig()> Function GetAttributesOf(ByVal cidl As UInteger, ByVal apidl As IntPtr(), ByRef rgfInOut As UInteger) As Integer
        <PreserveSig>
        Function GetUIObjectOf(
        ByVal hwndOwner As IntPtr,
        ByVal cidl As UInteger,
        <MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.SysInt), [In]> ByVal apidl As IntPtr(),
        <[In]> ByRef riid As Guid,
        ByRef rgfReserved As UInteger,
        <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppv As IContextMenu) As Integer
        <PreserveSig()> Function GetDisplayNameOf(ByVal pidl As IntPtr, ByVal uFlags As UInteger, ByRef lpName As STRRET) As Integer
        <PreserveSig()> Function SetNameOf(ByVal hwndOwner As IntPtr, ByVal pidl As IntPtr, ByVal pszName As String, ByVal uFlags As UInteger, ByRef ppidlOut As IntPtr) As Integer
    End Interface

    <ComImport(), Guid("000214e4-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IContextMenu
        <PreserveSig> Function QueryContextMenu(ByVal hmenu As IntPtr, ByVal indexMenu As UInteger, ByVal idCmdFirst As UInteger, ByVal idCmdLast As UInteger, ByVal uFlags As UInteger) As Integer

        ' Tady musí být "ByRef" a typ "CMINVOKECOMMANDINFOEX"
        <PreserveSig> Function InvokeCommand(ByRef pici As CMINVOKECOMMANDINFOEX) As Integer

        <PreserveSig>
        Function GetCommandString(ByVal idCmd As IntPtr, ByVal uType As UInteger, ByVal pwReserved As IntPtr, ByVal pszName As StringBuilder, ByVal cchMax As Integer) As Integer
    End Interface

    <ComImport(), Guid("000214f4-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IContextMenu2
        Inherits IContextMenu
        <PreserveSig> Function HandleMenuMsg(ByVal uMsg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Interface

    <ComImport(), Guid("bcfce0a0-1817-11d0-9d1d-00a0c9034933"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IContextMenu3
        Inherits IContextMenu2

        <PreserveSig>
        Overloads Function HandleMenuMsg(ByVal uMsg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByRef plResult As IntPtr) As Integer
    End Interface

    ' Other necessary COM interfaces/structs
    ' STRRET (for display names)
    <StructLayout(LayoutKind.Explicit, CharSet:=CharSet.Auto)>
    Public Structure STRRET
        <FieldOffset(0)> Public uType As UInteger
        <FieldOffset(4)> Public [pOleStr] As IntPtr ' Ptr to Ansi or Unicode string
        <FieldOffset(4)> Public uOffset As UInteger ' Offset into SHITEMID
        <FieldOffset(4)> Public [CStr] As IntPtr ' Ptr to Ansi or Unicode string
    End Structure

    ' PIDL (Pointer to Item ID List) - Not directly used, but conceptually important
    ' It's typically returned as IntPtr and managed by Shell functions

#End Region


    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetDesktopFolder(ByRef ppv As IShellFolder) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetSpecialFolderLocation(ByVal hwndOwner As IntPtr, ByVal nFolder As Integer, ByRef ppidl As IntPtr) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetPathFromIDList(ByVal pidl As IntPtr, ByVal pszPath As StringBuilder) As Boolean
    End Function

    <DllImport("shell32.dll", EntryPoint:="ShellExecuteExW", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function ShellExecuteEx(ByRef pExecInfo As SHELLEXECUTEINFO) As Boolean
    End Function

    <DllImport("shlwapi.dll", CharSet:=CharSet.Auto)>
    Public Shared Function PathCombine(ByVal lpszDest As StringBuilder, ByVal lpszDir As String, ByVal lpszFile As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function DestroyMenu(ByVal hMenu As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function TrackPopupMenuEx(ByVal hMenu As IntPtr, ByVal uFlags As UInteger, ByVal x As Integer, ByVal y As Integer, ByVal hWnd As IntPtr, ByVal lptpm As IntPtr) As Integer
    End Function

    <DllImport("user32.dll")>
    Public Shared Function CreatePopupMenu() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetMenuItemID(ByVal hMenu As IntPtr, ByVal nPos As Integer) As UInteger
    End Function

    ' For ListView hit testing (to get item at mouse click)
    <DllImport("user32.dll")>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Public Const LVM_GETITEM As UInteger = &H1000 + 75
    Public Const LVIF_TEXT As UInteger = &H1

    ' ShellExecuteInfo Structure
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure SHELLEXECUTEINFO
        Public cbSize As Integer
        Public fMask As UInteger
        Public hwnd As IntPtr
        <MarshalAs(UnmanagedType.LPTStr)> Public lpVerb As String
        <MarshalAs(UnmanagedType.LPTStr)> Public lpFile As String
        <MarshalAs(UnmanagedType.LPTStr)> Public lpParameters As String
        <MarshalAs(UnmanagedType.LPTStr)> Public lpDirectory As String
        Public nShow As Integer
        Public hInstApp As IntPtr
        Public lpIDList As IntPtr
        <MarshalAs(UnmanagedType.LPTStr)> Public lpClass As String
        Public hkeyClass As IntPtr
        Public dwHotKey As UInteger
        Public hIcon As IntPtr
        Public hProcess As IntPtr
    End Structure

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetFileInfo(ByVal pszPath As String,
                                             ByVal dwFileAttributes As UInteger,
                                             ByRef psfi As SHFILEINFO,
                                             ByVal cbFileInfo As UInteger,
                                             ByVal uFlags As UInteger) As IntPtr
    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> Public szTypeName As String
    End Structure

    ' SHELLEXECUTEINFO fMask flags
    Public Const SEE_MASK_IDLIST As UInteger = &H4
    Public Const SEE_MASK_CLASSNAME As UInteger = &H1000
    Public Const SEE_MASK_CLASSKEY As UInteger = &H3000
    Public Const SEE_MASK_HOTKEY As UInteger = &H20
    Public Const SEE_MASK_NOASYNC As UInteger = &H100
    Public Const SEE_MASK_FLAG_NO_UI As UInteger = &H400
    Public Const SEE_MASK_CONNECTNETDRV As UInteger = &H80
    Public Const SEE_MASK_DOENVSUBST As UInteger = &H200
    Public Const SEE_MASK_FLAG_DDEWAIT As UInteger = &H100
    Public Const SEE_MASK_NOCLOSEPROCESS As UInteger = &H40
    Public Const SEE_MASK_NO_CONSOLE As UInteger = &H8000
    Public Const SEE_MASK_ASYNCOK As UInteger = &H100000

End Class