Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

End Class

<System.Runtime.InteropServices.StructLayoutAttribute(System.Runtime.InteropServices.LayoutKind.Sequential)>
Public Structure HWND__

    '''int
    Public unused As Integer
End Structure

Partial Public Class NativeMethods

    '''Return Type: BOOL->int
    '''hWndNewOwner: HWND->HWND__*
    <System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint:="OpenClipboard")>
    Public Shared Function OpenClipboard(<System.Runtime.InteropServices.InAttribute()> ByVal hWndNewOwner As System.IntPtr) As <System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)> Boolean
    End Function
    '''Return Type: BOOL->int
    <System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint:="CloseClipboard")>
    Public Shared Function CloseClipboard() As <System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)> Boolean
    End Function
    '''Return Type: HANDLE->void*
    '''uFormat: UINT->unsigned int
    '''hMem: HANDLE->void*
    <System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint:="SetClipboardData")>
    Public Shared Function SetClipboardData(ByVal uFormat As UInteger, <System.Runtime.InteropServices.InAttribute()> ByVal hMem As System.IntPtr) As System.IntPtr
    End Function
    '''Return Type: BOOL->int
    <System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint:="EmptyClipboard")>
    Public Shared Function EmptyClipboard() As <System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)> Boolean
    End Function
    '''Return Type: LPVOID->void*
    '''hMem: HGLOBAL->HANDLE->void*
    <System.Runtime.InteropServices.DllImportAttribute("kernel32.dll", EntryPoint:="GlobalLock")>
    Public Shared Function GlobalLock(<System.Runtime.InteropServices.InAttribute()> ByVal hMem As System.IntPtr) As System.IntPtr
    End Function
    '''Return Type: BOOL->int
    '''hMem: HGLOBAL->HANDLE->void*
    <System.Runtime.InteropServices.DllImportAttribute("kernel32.dll", EntryPoint:="GlobalUnlock")>
    Public Shared Function GlobalUnlock(<System.Runtime.InteropServices.InAttribute()> ByVal hMem As System.IntPtr) As <System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)> Boolean
    End Function
    '''Return Type: HGLOBAL->HANDLE->void*
    '''uFlags: UINT->unsigned int
    '''dwBytes: SIZE_T->ULONG_PTR->unsigned int
    <System.Runtime.InteropServices.DllImportAttribute("kernel32.dll", EntryPoint:="GlobalAlloc")>
    Public Shared Function GlobalAlloc(ByVal uFlags As UInteger, ByVal dwBytes As UInteger) As System.IntPtr
    End Function
    '''Return Type: void*
    '''_Dst: void*
    '''_Src: void*
    '''_Size: size_t->unsigned int
    <System.Runtime.InteropServices.DllImportAttribute("ntdll.dll", EntryPoint:="memcpy", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Public Shared Function memcpy(ByVal _Dst As System.IntPtr, <System.Runtime.InteropServices.InAttribute()> ByVal _Src As System.IntPtr, <System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.SysUInt)> ByVal _Size As UInteger) As System.IntPtr
    End Function
End Class

Public Class Glossary
    Public Phrase As String
    Public Translation As String
End Class
