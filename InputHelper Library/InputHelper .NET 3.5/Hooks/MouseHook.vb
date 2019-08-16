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

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports InputHelper.EventArgs

Namespace Hooks
    ''' <summary>
    ''' A global low-level mouse hook that raises events when a mouse event occurs.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class MouseHook
        Implements IDisposable

        ''' <summary>
        ''' Occurs when a mouse button is pressed or held down.
        ''' </summary>
        ''' <remarks></remarks>
        Public Event MouseDown As EventHandler(Of MouseHookEventArgs)

        ''' <summary>
        ''' Occurs when a mouse button is released.
        ''' </summary>
        ''' <remarks></remarks>
        Public Event MouseUp As EventHandler(Of MouseHookEventArgs)

        ''' <summary>
        ''' Occurs when the mouse moves.
        ''' </summary>
        ''' <remarks></remarks>
        Public Event MouseMove As EventHandler(Of MouseHookEventArgs)

        ''' <summary>
        ''' Occurs when the mouse wheel is scrolled.
        ''' </summary>
        ''' <remarks></remarks>
        Public Event MouseWheel As EventHandler(Of MouseHookEventArgs)

        Private hHook As IntPtr = IntPtr.Zero
        Private HookProcedureDelegate As New NativeMethods.LowLevelMouseProc(AddressOf HookCallback)

        Private LeftClickTimeStamp As Integer = 0
        Private MiddleClickTimeStamp As Integer = 0
        Private RightClickTimeStamp As Integer = 0
        Private X1ClickTimeStamp As Integer = 0
        Private X2ClickTimeStamp As Integer = 0

        Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
            Dim Block As Boolean = False

            If nCode >= NativeMethods.HookCode.HC_ACTION AndAlso _
                (wParam = NativeMethods.MouseMessage.WM_LBUTTONDOWN OrElse _
                 wParam = NativeMethods.MouseMessage.WM_LBUTTONUP OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MBUTTONDOWN OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MBUTTONUP OrElse _
                 wParam = NativeMethods.MouseMessage.WM_RBUTTONDOWN OrElse _
                 wParam = NativeMethods.MouseMessage.WM_RBUTTONUP OrElse _
                 wParam = NativeMethods.MouseMessage.WM_XBUTTONDOWN OrElse _
                 wParam = NativeMethods.MouseMessage.WM_XBUTTONUP OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MOUSEWHEEL OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MOUSEHWHEEL OrElse _
                 wParam = NativeMethods.MouseMessage.WM_MOUSEMOVE) Then

                Dim MouseEventInfo As NativeMethods.MSLLHOOKSTRUCT = _
                    CType(Marshal.PtrToStructure(lParam, GetType(NativeMethods.MSLLHOOKSTRUCT)), NativeMethods.MSLLHOOKSTRUCT)

                Select Case wParam
                    Case NativeMethods.MouseMessage.WM_LBUTTONDOWN
                        Dim DoubleClick As Boolean = (Environment.TickCount - LeftClickTimeStamp) <= NativeMethods.GetDoubleClickTime()
                        LeftClickTimeStamp = Environment.TickCount

                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.Left, KeyState.Down, DoubleClick, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseDown(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_LBUTTONUP
                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.Left, KeyState.Up, False, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MBUTTONDOWN
                        Dim DoubleClick As Boolean = (Environment.TickCount - MiddleClickTimeStamp) <= NativeMethods.GetDoubleClickTime()
                        MiddleClickTimeStamp = Environment.TickCount

                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.Middle, KeyState.Down, DoubleClick, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseDown(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MBUTTONUP
                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.Middle, KeyState.Up, False, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_RBUTTONDOWN
                        Dim DoubleClick As Boolean = (Environment.TickCount - RightClickTimeStamp) <= NativeMethods.GetDoubleClickTime()
                        RightClickTimeStamp = Environment.TickCount

                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.Right, KeyState.Down, DoubleClick, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseDown(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_RBUTTONUP
                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.Right, KeyState.Up, False, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_XBUTTONDOWN
                        Dim IsXButton2 As Boolean = (New NativeMethods.DWORD(MouseEventInfo.mouseData).High = 2)
                        Dim DoubleClick As Boolean = (Environment.TickCount - If(IsXButton2, X2ClickTimeStamp, X1ClickTimeStamp)) <= NativeMethods.GetDoubleClickTime()

                        If IsXButton2 = True Then X2ClickTimeStamp = Environment.TickCount _
                                             Else X1ClickTimeStamp = Environment.TickCount

                        Dim HookEventArgs As New MouseHookEventArgs(If(IsXButton2, MouseButtons.XButton2, MouseButtons.XButton1), KeyState.Down, DoubleClick, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_XBUTTONUP
                        Dim IsXButton2 As Boolean = (New NativeMethods.DWORD(MouseEventInfo.mouseData).High = 2)
                        Dim HookEventArgs As New MouseHookEventArgs(If(IsXButton2, MouseButtons.XButton2, MouseButtons.XButton1), KeyState.Up, False, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseUp(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MOUSEWHEEL
                        Dim Delta As Integer = New NativeMethods.DWORD(MouseEventInfo.mouseData).SignedHigh
                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.None, KeyState.Up, False, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.Vertical, Delta)
                        RaiseEvent MouseWheel(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MOUSEHWHEEL
                        Dim Delta As Integer = New NativeMethods.DWORD(MouseEventInfo.mouseData).SignedHigh
                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.None, KeyState.Up, False, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.Horizontal, Delta)
                        RaiseEvent MouseWheel(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                    Case NativeMethods.MouseMessage.WM_MOUSEMOVE
                        Dim HookEventArgs As New MouseHookEventArgs(MouseButtons.None, KeyState.Up, False, _
                                                                    New Point(MouseEventInfo.pt.x, MouseEventInfo.pt.y), ScrollDirection.None, 0)
                        RaiseEvent MouseMove(Me, HookEventArgs)
                        Block = HookEventArgs.Block


                End Select
            End If

            Return If(Block, New IntPtr(1), NativeMethods.CallNextHookEx(hHook, nCode, wParam, lParam))
        End Function

        ''' <summary>
        ''' Initializes a new instance of the MouseHook clas.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            hHook = NativeMethods.SetWindowsHookEx(NativeMethods.HookType.WH_MOUSE_LL, HookProcedureDelegate, NativeMethods.GetModuleHandle(Nothing), 0)
            If hHook = IntPtr.Zero Then
                Dim Win32Error As Integer = Marshal.GetLastWin32Error()
                Throw New Win32Exception(Win32Error, "Failed to create mouse hook! (" & Win32Error & ")")
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