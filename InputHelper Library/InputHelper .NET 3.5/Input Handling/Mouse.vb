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

''' <summary>
''' A static class for handling and simulating physical mouse input.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Mouse
    Private Sub New()
    End Sub

#Region "Public methods"

#Region "IsButtonDown()"
    ''' <summary>
    ''' Checks whether the specified mouse button is currently held down.
    ''' </summary>
    ''' <param name="Button">The mouse button to check.</param>
    ''' <remarks></remarks>
    Public Shared Function IsButtonDown(ByVal Button As MouseButtons) As Boolean
        Dim Key As Keys = Keys.None
        Select Case Button
            Case MouseButtons.Left : Key = Keys.LButton
            Case MouseButtons.Middle : Key = Keys.MButton
            Case MouseButtons.Right : Key = Keys.RButton
            Case MouseButtons.XButton1 : Key = Keys.XButton1
            Case MouseButtons.XButton2 : Key = Keys.XButton2
            Case Else
                Throw New ArgumentException("Invalid mouse button " & Button.ToString() & "!", "Button")
        End Select

        Return Keyboard.IsKeyDown(Key)
    End Function
#End Region

#Region "IsButtonUp()"
    ''' <summary>
    ''' Checks whether the specified mouse button is currently NOT held down.
    ''' </summary>
    ''' <param name="Button">The mouse button to check.</param>
    ''' <remarks></remarks>
    Public Shared Function IsButtonUp(ByVal Button As MouseButtons) As Boolean
        Return Mouse.IsButtonDown(Button) = False
    End Function
#End Region

#Region "PressButton()"
    ''' <summary>
    ''' Simulates a mouse button click.
    ''' </summary>
    ''' <param name="Button">The button to press.</param>
    ''' <remarks></remarks>
    Public Shared Sub PressButton(ByVal Button As MouseButtons)
        Mouse.SetButtonState(Button, True)
        Mouse.SetButtonState(Button, False)
    End Sub
#End Region

#Region "SetButtonState()"
    ''' <summary>
    ''' Simulates a mouse button being pushed down or released.
    ''' </summary>
    ''' <param name="Button">The button which to simulate.</param>
    ''' <param name="MouseDown">Whether to push down or release the mouse button.</param>
    ''' <remarks></remarks>
    Public Shared Sub SetButtonState(ByVal Button As MouseButtons, ByVal MouseDown As Boolean)
        Dim InputList As NativeMethods.INPUT() = _
            New NativeMethods.INPUT(1 - 1) {Mouse.GetMouseClickInputStructure(Button, MouseDown)}
        NativeMethods.SendInput(CType(InputList.Length, UInteger), InputList, Marshal.SizeOf(GetType(NativeMethods.INPUT)))
    End Sub
#End Region

#End Region

#Region "Internal methods"

#Region "GetMouseClickInputStructure()"
    ''' <summary>
    ''' Constructs a native mouse INPUT structure for click events that can be passed to SendInput().
    ''' </summary>
    ''' <param name="Button">The button of the event.</param>
    ''' <param name="MouseDown">Whether to push down or release the mouse button.</param>
    ''' <remarks></remarks>
    Private Shared Function GetMouseClickInputStructure(ByVal Button As MouseButtons, ByVal MouseDown As Boolean) As NativeMethods.INPUT
        Dim Position As Point = Cursor.Position
        Dim MouseFlags As NativeMethods.MOUSEEVENTF
        Dim MouseData As UInteger = 0

        Select Case Button
            Case MouseButtons.Left : MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.LEFTDOWN, NativeMethods.MOUSEEVENTF.LEFTUP)
            Case MouseButtons.Middle : MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.MIDDLEDOWN, NativeMethods.MOUSEEVENTF.MIDDLEUP)
            Case MouseButtons.Right : MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.RIGHTDOWN, NativeMethods.MOUSEEVENTF.RIGHTUP)
            Case MouseButtons.XButton1
                MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.XDOWN, NativeMethods.MOUSEEVENTF.XUP)
                MouseData = NativeMethods.MouseXButton.XBUTTON1
            Case MouseButtons.XButton2
                MouseFlags = If(MouseDown, NativeMethods.MOUSEEVENTF.XDOWN, NativeMethods.MOUSEEVENTF.XUP)
                MouseData = NativeMethods.MouseXButton.XBUTTON2
            Case Else
                Throw New ArgumentException("Invalid mouse button " & Button.ToString() & "!", "Button")
        End Select

        Dim MouseInput As New NativeMethods.MOUSEINPUT With {
            .dx = Position.X,
            .dy = Position.Y,
            .mouseData = MouseData,
            .dwFlags = MouseFlags,
            .time = 0,
            .dwExtraInfo = UIntPtr.Zero
        }

        Dim Union As New NativeMethods.INPUTUNION With {.mi = MouseInput}
        Dim Input As New NativeMethods.INPUT With {
            .type = NativeMethods.INPUTTYPE.MOUSE,
            .U = Union
        }

        Return Input
    End Function
#End Region

#End Region

End Class