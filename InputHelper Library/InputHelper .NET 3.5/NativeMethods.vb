'+---------------------------------------------------------------------------------+'
'|                               === InputHelper ===                               |'
'|                                                                                 |'
'|                  Created by Vincent "Visual Vincent" Bengtsson                  |'
'|                    Website: https://www.mydoomsite.com/                         |'
'|                    GitHub:  https://github.com/Visual-Vincent                   |'
'|                                                                                 |'
'|                                                                                 |'
'|                            === COPYRIGHT LICENSE ===                            |'
'|                                                                                 |'
'| Copyright (c) 2016-2019, Vincent Bengtsson                                      |'
'| All rights reserved.                                                            |'
'|                                                                                 |'
'| Redistribution and use in source and binary forms, with or without              |'
'| modification, are permitted provided that the following conditions are met:     |'
'|                                                                                 |'
'| 1. Redistributions of source code must retain the above copyright notice, this  |'
'|    list of conditions and the following disclaimer.                             |'
'|                                                                                 |'
'| 2. Redistributions in binary form must reproduce the above copyright notice,    |'
'|    this list of conditions and the following disclaimer in the documentation    |'
'|    and/or other materials provided with the distribution.                       |'
'|                                                                                 |'
'| 3. Neither the name of the copyright holder nor the names of its                |'
'|    contributors may be used to endorse or promote products derived from         |'
'|    this software without specific prior written permission.                     |'
'|                                                                                 |'
'| THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"     |'
'| AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE       |'
'| IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE  |'
'| DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE    |'
'| FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL      |'
'| DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR      |'
'| SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER      |'
'| CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,   |'
'| OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE   |'
'| OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.            |'
'+---------------------------------------------------------------------------------+'

Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Text

''' <summary>
''' A class containing all the native WinAPI methods, structures, declarations, etc. used by InputHelper.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class NativeMethods
    Private Sub New()
    End Sub

#Region "Delegates"
    Public Delegate Function KeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Public Delegate Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Public Delegate Function LowLevelMouseProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
#End Region

#Region "Methods"

#Region "Hook methods"
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetWindowsHookEx(ByVal idHook As HookType, ByVal lpfn As KeyboardProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetWindowsHookEx(ByVal idHook As HookType, ByVal lpfn As LowLevelKeyboardProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetWindowsHookEx(ByVal idHook As HookType, ByVal lpfn As LowLevelMouseProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
#End Region

#Region "Input methods"
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SendInput(ByVal nInputs As UInteger, <MarshalAs(UnmanagedType.LPArray)> ByVal pInputs() As INPUT, ByVal cbSize As Integer) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function MapVirtualKeyEx(uCode As UInteger, uMapType As UInteger, dwhkl As IntPtr) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetKeyboardLayout(idThread As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetKeyboardState(<MarshalAs(UnmanagedType.LPArray)> ByVal lpKeyState As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetAsyncKeyState(ByVal vKey As Keys) As Short
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function ToUnicodeEx(ByVal wVirtKey As UInteger, ByVal wScanCode As UInteger, <MarshalAs(UnmanagedType.LPArray)> ByVal lpKeyState As Byte(), <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszBuff As StringBuilder, ByVal cchBuff As Integer, ByVal wFlags As UInteger, ByVal dwhkl As IntPtr) As Integer
    End Function
#End Region

#Region "Window and thread methods"
    'CharSet.Unicode is required for CP437 ALT codes to work!
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetFocus() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function AttachThreadInput(ByVal idAttach As UInteger, ByVal idAttachTo As UInteger, <MarshalAs(UnmanagedType.Bool)> fAttach As Boolean) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function GetCurrentThreadId() As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, <Out(), [Optional]()> lpdwProcessId As UIntPtr) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function WindowFromPoint(ByVal p As NATIVEPOINT) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function ScreenToClient(ByVal hWnd As IntPtr, ByRef lpPoint As NATIVEPOINT) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function
#End Region

#Region "System information methods"
    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetDoubleClickTime() As UInteger
    End Function
#End Region

#End Region

#Region "Enumerations"
    Public Enum HookType As Integer
        WH_CALLWNDPROC = 4
        WH_CALLWNDPROCRET = 12
        WH_CBT = 5
        WH_DEBUG = 9
        WH_FOREGROUNDIDLE = 11
        WH_GETMESSAGE = 3
        WH_JOURNALPLAYBACK = 1
        WH_JOURNALRECORD = 0
        WH_KEYBOARD = 2
        WH_KEYBOARD_LL = 13
        WH_MOUSE = 7
        WH_MOUSE_LL = 14
        WH_MSGFILTER = -1
        WH_SHELL = 10
        WH_SYSMSGFILTER = 6
    End Enum

    Public Enum HookCode As Integer
        HC_ACTION = 0
        HC_NOREMOVE = 3
    End Enum

    Public Enum KeyboardFlags As UInteger
        KF_EXTENDED = &H100
        KF_DLGMODE = &H800
        KF_MENUMODE = &H1000
        KF_ALTDOWN = &H2000
        KF_REPEAT = &H4000
        KF_UP = &H8000
    End Enum

    Public Enum LowLevelKeyboardHookFlags As UInteger
        LLKHF_EXTENDED = &H1
        LLKHF_LOWER_IL_INJECTED = &H2
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum

    Public Enum LowLevelMouseHookFlags As UInteger
        LLMHF_INJECTED = &H1
        LLMHF_LOWER_IL_INJECTED = &H2
    End Enum

    Public Enum INPUTTYPE As UInteger
        MOUSE = 0
        KEYBOARD = 1
        HARDWARE = 2
    End Enum

    <Flags()> _
    Public Enum KEYEVENTF As UInteger
        ''' <summary>
        ''' If specified, the scan code was preceded by a prefix byte that has the value 0xE0 (224).
        ''' </summary>
        ''' <remarks></remarks>
        EXTENDEDKEY = &H1

        ''' <summary>
        ''' If specified, the key is being released. If not specified, the key is being pressed.
        ''' </summary>
        ''' <remarks></remarks>
        KEYUP = &H2

        ''' <summary>
        ''' If specified, wScan identifies the key and wVk is ignored.
        ''' </summary>
        ''' <remarks></remarks>
        SCANCODE = &H8

        ''' <summary>
        ''' If specified, the system synthesizes a VK_PACKET keystroke. The wVk parameter must be zero. This flag can only be combined with the KEYEVENTF_KEYUP flag.
        ''' </summary>
        ''' <remarks></remarks>
        UNICODE = &H4
    End Enum

    <Flags()> _
    Public Enum MOUSEEVENTF As UInteger
        ''' <summary>
        ''' The dx and dy members contain normalized absolute (screen) coordinates.
        ''' </summary>
        ''' <remarks></remarks>
        ABSOLUTE = &H8000

        ''' <summary>
        ''' The wheel was moved horizontally, if the mouse has a wheel. The amount of movement is specified in mouseData.
        ''' </summary>
        ''' <remarks></remarks>
        HWHEEL = &H1000

        ''' <summary>
        ''' Movement occurred.
        ''' </summary>
        ''' <remarks></remarks>
        MOVE = &H1

        ''' <summary>
        ''' The WM_MOUSEMOVE messages will not be coalesced. The default behavior is to coalesce WM_MOUSEMOVE messages.
        ''' </summary>
        ''' <remarks></remarks>
        MOVE_NOCOALESCE = &H2000

        ''' <summary>
        ''' The left button was pressed.
        ''' </summary>
        ''' <remarks></remarks>
        LEFTDOWN = &H2

        ''' <summary>
        ''' The left button was released.
        ''' </summary>
        ''' <remarks></remarks>
        LEFTUP = &H4

        ''' <summary>
        ''' The right button was pressed.
        ''' </summary>
        ''' <remarks></remarks>
        RIGHTDOWN = &H8

        ''' <summary>
        ''' The right button was released.
        ''' </summary>
        ''' <remarks></remarks>
        RIGHTUP = &H10

        ''' <summary>
        ''' The middle button was pressed.
        ''' </summary>
        ''' <remarks></remarks>
        MIDDLEDOWN = &H20

        ''' <summary>
        ''' The middle button was released.
        ''' </summary>
        ''' <remarks></remarks>
        MIDDLEUP = &H40

        ''' <summary>
        ''' Maps coordinates to the entire desktop. Must be used with MOUSEEVENTF_ABSOLUTE.
        ''' </summary>
        ''' <remarks></remarks>
        VIRTUALDESK = &H4000

        ''' <summary>
        ''' The wheel was moved, if the mouse has a wheel. The amount of movement is specified in mouseData.
        ''' </summary>
        ''' <remarks></remarks>
        WHEEL = &H800

        ''' <summary>
        ''' An X button was pressed.
        ''' </summary>
        ''' <remarks></remarks>
        XDOWN = &H80

        ''' <summary>
        ''' An X button was released.
        ''' </summary>
        ''' <remarks></remarks>
        XUP = &H100
    End Enum

    Public Enum MouseXButton As UInteger
        XBUTTON1 = &H1
        XBUTTON2 = &H2
    End Enum

    Public Enum KeyMessage As UInteger
        WM_KEYDOWN = &H100
        WM_KEYUP = &H101
        WM_CHAR = &H102
        WM_DEADCHAR = &H103
        WM_SYSKEYDOWN = &H104
        WM_SYSKEYUP = &H105
        WM_SYSCHAR = &H106
        WM_SYSDEADCHAR = &H107
        WM_UNICHAR = &H109
    End Enum

    Public Enum MouseMessage As Integer
        WM_MOUSEMOVE = &H200
        WM_LBUTTONDOWN = &H201
        WM_LBUTTONUP = &H202
        WM_LBUTTONDBLCLK = &H203
        WM_MBUTTONDOWN = &H207
        WM_MBUTTONUP = &H208
        WM_MBUTTONDBLCLK = &H209
        WM_RBUTTONDOWN = &H204
        WM_RBUTTONUP = &H205
        WM_RBUTTONDBLCLK = &H206
        WM_MOUSEWHEEL = &H20A
        WM_MOUSEHWHEEL = &H20E
        WM_XBUTTONDOWN = &H20B
        WM_XBUTTONUP = &H20C
        WM_XBUTTONDBLCLK = &H20D
    End Enum
#End Region

#Region "Structures"
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure KBDLLHOOKSTRUCT
        Public vkCode As UInteger
        Public scanCode As UInteger
        Public flags As LowLevelKeyboardHookFlags
        Public time As UInteger
        Public dwExtraInfo As UIntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MSLLHOOKSTRUCT
        Public pt As NATIVEPOINT
        Public mouseData As UInteger
        Public flags As UInteger
        Public time As UInteger
        Public dwExtraInfo As UIntPtr
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Public Structure DWORD
        <FieldOffset(0)> Public Value As UInteger
        <FieldOffset(0)> Public Low As UShort
        <FieldOffset(2)> Public High As UShort

        <FieldOffset(0)> Public SignedValue As Integer
        <FieldOffset(0)> Public SignedLow As Short
        <FieldOffset(2)> Public SignedHigh As Short

        Public Shared Widening Operator CType(ByVal DWORD As DWORD) As IntPtr
            Return New IntPtr(DWORD.SignedValue)
        End Operator

        Public Shared Widening Operator CType(ByVal DWORD As DWORD) As UInteger
            Return DWORD.Value
        End Operator

        Public Shared Widening Operator CType(ByVal DWORD As DWORD) As Integer
            Return DWORD.SignedValue
        End Operator

        Public Sub New(ByVal Value As UInteger)
            Me.Value = Value
        End Sub

        Public Sub New(ByVal Value As Integer)
            Me.SignedValue = Value
        End Sub

        Public Sub New(ByVal Low As UShort, ByVal High As UShort)
            Me.Low = Low
            Me.High = High
        End Sub

        Public Sub New(ByVal Low As Short, ByVal High As Short)
            Me.SignedLow = Low
            Me.SignedHigh = High
        End Sub
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Public Structure QWORD
        <FieldOffset(0)> Public Value As ULong
        <FieldOffset(0)> Public Low As UInteger
        <FieldOffset(4)> Public High As UInteger

        <FieldOffset(0)> Public SignedValue As Long
        <FieldOffset(0)> Public SignedLow As Integer
        <FieldOffset(4)> Public SignedHigh As Integer

        <FieldOffset(0)> Public LowWord As DWORD
        <FieldOffset(4)> Public HighWord As DWORD

        Public Shared Widening Operator CType(ByVal QWORD As QWORD) As ULong
            Return QWORD.Value
        End Operator

        Public Shared Widening Operator CType(ByVal QWORD As QWORD) As Long
            Return QWORD.SignedValue
        End Operator

        Public Sub New(ByVal Value As ULong)
            Me.Value = Value
        End Sub

        Public Sub New(ByVal Value As Long)
            Me.SignedValue = Value
        End Sub

        Public Sub New(ByVal Low As UInteger, ByVal High As UInteger)
            Me.Low = Low
            Me.High = High
        End Sub

        Public Sub New(ByVal Low As Integer, ByVal High As Integer)
            Me.SignedLow = Low
            Me.SignedHigh = High
        End Sub
    End Structure

    <StructLayout(LayoutKind.Explicit)> _
    Public Structure INPUTUNION
        <FieldOffset(0)> Public mi As MOUSEINPUT
        <FieldOffset(0)> Public ki As KEYBDINPUT
        <FieldOffset(0)> Public hi As HARDWAREINPUT
    End Structure

    Public Structure INPUT
        Public type As INPUTTYPE
        Public U As INPUTUNION
    End Structure

    Public Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As UInteger
        Public dwFlags As UInteger
        Public time As UInteger
        Public dwExtraInfo As UIntPtr
    End Structure

    Public Structure KEYBDINPUT
        Public wVk As UShort
        Public wScan As UShort
        Public dwFlags As UInteger
        Public time As UInteger
        Public dwExtraInfo As UIntPtr
    End Structure

    Public Structure HARDWAREINPUT
        Public uMsg As UInteger
        Public wParamL As UShort
        Public wParamH As UShort
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure NATIVEPOINT
        Public x As Integer
        Public y As Integer

        Public Sub New(ByVal X As Integer, ByVal Y As Integer)
            Me.x = X
            Me.y = Y
        End Sub
    End Structure
#End Region

End Class
