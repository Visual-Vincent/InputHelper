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
'| Copyright (c) 2016-2022, Vincent Bengtsson                                      |'
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
Imports System.ComponentModel
Imports InputHelper.EventArgs

Namespace Hooks
    ''' <summary>
    ''' A local keyboard hook that raises events when a key is pressed or released in a specific thread.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class LocalKeyboardHook
        Implements IDisposable

        ''' <summary>
        ''' Occurs when a key is pressed or held down.
        ''' </summary>
        ''' <remarks></remarks>
        Public Event KeyDown As EventHandler(Of KeyboardHookEventArgs)

        ''' <summary>
        ''' Occurs when a key is released.
        ''' </summary>
        ''' <remarks></remarks>
        Public Event KeyUp As EventHandler(Of KeyboardHookEventArgs)

        Private hHook As IntPtr = IntPtr.Zero
        Private HookProcedureDelegate As New NativeMethods.KeyboardProc(AddressOf HookCallback)

        Private Modifiers As ModifierKeys = ModifierKeys.None

        Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
            Dim Block As Boolean = False

            If nCode >= NativeMethods.HookCode.HC_ACTION Then
                Dim KeyCode As Keys = CType(wParam.ToInt32(), Keys)
                Dim KeyFlags As New NativeMethods.QWORD(lParam.ToInt64())
                Dim ScanCode As Byte = BitConverter.GetBytes(KeyFlags)(2) 'The scan code is the third byte in the integer (bits 16-23).
                Dim Extended As Boolean = (KeyFlags.LowWord.High And NativeMethods.KeyboardFlags.KF_EXTENDED) = NativeMethods.KeyboardFlags.KF_EXTENDED
                Dim AltDown As Boolean = (KeyFlags.LowWord.High And NativeMethods.KeyboardFlags.KF_ALTDOWN) = NativeMethods.KeyboardFlags.KF_ALTDOWN
                Dim KeyUp As Boolean = (KeyFlags.LowWord.High And NativeMethods.KeyboardFlags.KF_UP) = NativeMethods.KeyboardFlags.KF_UP

                'Set the ALT modifier if the KF_ALTDOWN flag is set.
                If AltDown = True _
                    AndAlso Internal.IsModifier(KeyCode, ModifierKeys.Alt) = False _
                     AndAlso (Me.Modifiers And ModifierKeys.Alt) <> ModifierKeys.Alt Then

                    Me.Modifiers = Me.Modifiers Or ModifierKeys.Alt
                End If

                'Raise KeyDown/KeyUp event.
                If KeyUp = False Then
                    Dim HookEventArgs As New KeyboardHookEventArgs(KeyCode, ScanCode, Extended, KeyState.Down, Me.Modifiers)

                    If Internal.IsModifier(KeyCode, ModifierKeys.Control) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Control
                    If Internal.IsModifier(KeyCode, ModifierKeys.Shift) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Shift
                    If Internal.IsModifier(KeyCode, ModifierKeys.Alt) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Alt
                    If Internal.IsModifier(KeyCode, ModifierKeys.Windows) = True Then Me.Modifiers = Me.Modifiers Or ModifierKeys.Windows

                    RaiseEvent KeyDown(Me, HookEventArgs)
                    Block = HookEventArgs.Block
                Else
                    'Must be done before creating the HookEventArgs during KeyUp.
                    If Internal.IsModifier(KeyCode, ModifierKeys.Control) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Control
                    If Internal.IsModifier(KeyCode, ModifierKeys.Shift) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Shift
                    If Internal.IsModifier(KeyCode, ModifierKeys.Alt) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Alt
                    If Internal.IsModifier(KeyCode, ModifierKeys.Windows) = True Then Me.Modifiers = Me.Modifiers And Not ModifierKeys.Windows

                    Dim HookEventArgs As New KeyboardHookEventArgs(KeyCode, ScanCode, Extended, KeyState.Up, Me.Modifiers)

                    RaiseEvent KeyUp(Me, HookEventArgs)
                    Block = HookEventArgs.Block
                End If
            End If

            Return If(Block, New IntPtr(1), NativeMethods.CallNextHookEx(hHook, nCode, wParam, lParam))
        End Function

        ''' <summary>
        ''' Initializes a new instance of the LocalKeyboardHook class attached to the current thread.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(NativeMethods.GetCurrentThreadId())
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the LocalKeyboardHook class attached to the specified thread.
        ''' </summary>
        ''' <param name="ThreadID">The thread to attach the hook to.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ThreadID As UInteger)
            hHook = NativeMethods.SetWindowsHookEx(NativeMethods.HookType.WH_KEYBOARD, HookProcedureDelegate, IntPtr.Zero, ThreadID)
            If hHook = IntPtr.Zero Then
                Dim Win32Error As Integer = Marshal.GetLastWin32Error()
                Throw New Win32Exception(Win32Error, "Failed to create local keyboard hook! (" & Win32Error & ")")
            End If
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
                If hHook <> IntPtr.Zero Then NativeMethods.UnhookWindowsHookEx(hHook)
            End If
            Me.disposedValue = True
        End Sub

        Protected Overrides Sub Finalize()
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(False)
            MyBase.Finalize()
        End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class
End Namespace