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

''' <summary>
''' A static class for handling and simulating physical keyboard input.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Keyboard
    Private Sub New()
    End Sub

#Region "Public methods"

#Region "IsKeyDown()"
    ''' <summary>
    ''' Checks whether the specified key is currently held down.
    ''' </summary>
    ''' <param name="Key">The key to check.</param>
    ''' <remarks></remarks>
    Public Shared Function IsKeyDown(ByVal Key As Keys) As Boolean
        Dim Modifiers As Keys() = Internal.ExtractModifiers(Key)
        For Each Modifier As Keys In Modifiers
            If (NativeMethods.GetAsyncKeyState(Modifier) And Constants.KeyDownBit) <> Constants.KeyDownBit Then
                Return False
            End If
        Next

        If Key = Keys.None Then Return True 'All modifiers are held down, no more keys left to check.

        Return (NativeMethods.GetAsyncKeyState(Key) And Constants.KeyDownBit) = Constants.KeyDownBit
    End Function
#End Region

#Region "IsKeyUp()"
    ''' <summary>
    ''' Checks whether the specified key is currently NOT held down.
    ''' </summary>
    ''' <param name="Key">The key to check.</param>
    ''' <remarks></remarks>
    Public Shared Function IsKeyUp(ByVal Key As Keys) As Boolean
        Return Keyboard.IsKeyDown(Key) = False
    End Function
#End Region

#Region "PressKey()"
    ''' <summary>
    ''' Simulates a keystroke.
    ''' </summary>
    ''' <param name="Key">The key to press.</param>
    ''' <param name="HardwareKey">Whether to simulate the keystroke using its virtual key code or its hardware scan code.</param>
    ''' <remarks></remarks>
    Public Shared Sub PressKey(ByVal Key As Keys, Optional ByVal HardwareKey As Boolean = False)
        Keyboard.SetKeyState(Key, True, HardwareKey)
        Keyboard.SetKeyState(Key, False, HardwareKey)
    End Sub
#End Region

#Region "SetKeyState()"
    ''' <summary>
    ''' Simulates a key being pushed down or released.
    ''' </summary>
    ''' <param name="Key">The key which to simulate.</param>
    ''' <param name="KeyDown">Whether to push down or release the key.</param>
    ''' <param name="HardwareKey">Whether to simulate the event using the key's virtual key code or its hardware scan code.</param>
    ''' <remarks></remarks>
    Public Shared Sub SetKeyState(ByVal Key As Keys, ByVal KeyDown As Boolean, Optional ByVal HardwareKey As Boolean = False)
        Dim InputList As New List(Of NativeMethods.INPUT)
        Dim Modifiers As Keys() = Internal.ExtractModifiers(Key)

        For Each Modifier As Keys In Modifiers
            InputList.Add(Keyboard.GetKeyboardInputStructure(Modifier, KeyDown, HardwareKey))
        Next
        InputList.Add(Keyboard.GetKeyboardInputStructure(Key, KeyDown, HardwareKey))

        NativeMethods.SendInput(CType(InputList.Count, UInteger), InputList.ToArray(), Marshal.SizeOf(GetType(NativeMethods.INPUT)))
    End Sub
#End Region

#End Region

#Region "Internal methods"

#Region "GetKeyboardInputStructure()"
    ''' <summary>
    ''' Constructs a native keyboard INPUT structure that can be passed to SendInput().
    ''' </summary>
    ''' <param name="Key">The key to send.</param>
    ''' <param name="KeyDown">Whether to send a KeyDown/KeyUp stroke.</param>
    ''' <param name="HardwareKey">Whether to send the key's hardware scan code instead of its virtual key code.</param>
    ''' <remarks></remarks>
    Private Shared Function GetKeyboardInputStructure(ByVal Key As Keys, ByVal KeyDown As Boolean, ByVal HardwareKey As Boolean) As NativeMethods.INPUT
        Dim KeyboardInput As New NativeMethods.KEYBDINPUT With {
            .wVk = If(HardwareKey = False, (Key And UShort.MaxValue), 0),
            .wScan = If(HardwareKey,
                        NativeMethods.MapVirtualKeyEx(CType(Key, UInteger), 0, NativeMethods.GetKeyboardLayout(0)),
                        0) And UShort.MaxValue,
            .time = 0,
            .dwFlags = If(HardwareKey, NativeMethods.KEYEVENTF.SCANCODE, 0) Or
                        If(KeyDown = False, NativeMethods.KEYEVENTF.KEYUP, 0),
            .dwExtraInfo = UIntPtr.Zero
        }

        Dim Union As New NativeMethods.INPUTUNION With {.ki = KeyboardInput}
        Dim Input As New NativeMethods.INPUT With {
            .type = NativeMethods.INPUTTYPE.KEYBOARD,
            .U = Union
        }

        Return Input
    End Function
#End Region

#End Region

End Class