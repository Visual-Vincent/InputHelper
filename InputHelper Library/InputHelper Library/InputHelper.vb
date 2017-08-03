'+---------------------------------------------------------------------------------+'
'|                               === InputHelper ===                               |'
'|                                                                                 |'
'|                  Created by Vincent "Visual Vincent" Bengtsson                  |'
'|                      Website: https://www.mydoomsite.com/                       |'
'|                                                                                 |'
'|                                                                                 |'
'|                            === COPYRIGHT LICENSE ===                            |'
'|                                                                                 |'
'| Copyright (c) 2016-2017, Vincent Bengtsson                                      |'
'| All rights reserved.                                                            |'
'|                                                                                 |'
'| Redistribution and use in source and binary forms, with or without              |'
'| modification, are permitted provided that the following conditions are met:     |'
'|                                                                                 |'
'| 1. Redistributions of source code must retain the above copyright notice, this  |'
'|    list of conditions and the following disclaimer.                             |'
'| 2. Redistributions in binary form must reproduce the above copyright notice,    |'
'|    this list of conditions and the following disclaimer in the documentation    |'
'|    and/or other materials provided with the distribution.                       |'
'|                                                                                 |'
'| THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND |'
'| ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED   |'
'| WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE          |'
'| DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR |'
'| ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES  |'
'| (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;    |'
'| LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND     |'
'| ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT      |'
'| (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS   |'
'| SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.                    |'
'+---------------------------------------------------------------------------------+'

Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text

''' <summary>
''' A static class for handling and simulating mouse and keyboard input.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class InputHelper
    Private Sub New()
    End Sub

#Region "Fields"
    ''' <summary>
    ''' The Alt Gr key.
    ''' </summary>
    ''' <remarks></remarks>
    Public Const AltGr As Keys = Keys.Control Or Keys.Alt

    ''' <summary>
    ''' The bit value to check if a key is held down when calling GetAsyncKeyState().
    ''' </summary>
    ''' <remarks></remarks>
    Private Const KeyDownBit As Integer = &H8000

    ''' <summary>
    ''' An array holding all keys' current states (down/up).
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared KeyboardState As Byte() = New Byte(256 - 1) {}

    ''' <summary>
    ''' Indicates whether we are currently in an ALT+numpad combination (ALT code).
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared IsAltCodeCombination As Boolean = False

    ''' <summary>
    ''' A list holding the keys pressed during an ALT+numpad combination.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared AltCode As New List(Of Keys)

    ''' <summary>
    ''' Windows-1252 encoding for the ALT+numpad combinations.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared ReadOnly Windows1252 As Encoding = Encoding.GetEncoding("windows-1252")

    ''' <summary>
    ''' Whether the current alt code is a Codepage 437 code.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared IsAltCodeCP437 As Boolean = False

    ''' <summary>
    ''' Lookup table for Codepage 437-to-Unicode character codes.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared ReadOnly CP437LookupTable As Integer() = _
        New Integer(256 - 1) { _
            0, 9786, 9787, 9829, 9830, 9827, 9824, _
            8226, 9688, 9675, 9689, 9794, 9792, 9834, 9835, _
            9788, 9658, 9668, 8597, 8252, 182, 167, 9644, _
            8616, 8593, 8595, 8594, 8592, 8735, 8596, 9650, _
            9660, 32, 33, 34, 35, 36, 37, 38, _
            39, 40, 41, 42, 43, 44, 45, 46, _
            47, 48, 49, 50, 51, 52, 53, 54, _
            55, 56, 57, 58, 59, 60, 61, 62, _
            63, 64, 65, 66, 67, 68, 69, 70, _
            71, 72, 73, 74, 75, 76, 77, 78, _
            79, 80, 81, 82, 83, 84, 85, 86, _
            87, 88, 89, 90, 91, 92, 93, 94, _
            95, 96, 97, 98, 99, 100, 101, 102, _
            103, 104, 105, 106, 107, 108, 109, 110, _
            111, 112, 113, 114, 115, 116, 117, 118, _
            119, 120, 121, 122, 123, 124, 125, 126, _
            8962, 199, 252, 233, 226, 228, 224, 229, _
            231, 234, 235, 232, 239, 238, 236, 196, _
            197, 201, 230, 198, 244, 246, 242, 251, _
            249, 255, 214, 220, 162, 163, 165, 8359, _
            402, 225, 237, 243, 250, 241, 209, 170, _
            186, 191, 8976, 172, 189, 188, 161, 171, _
            187, 9617, 9618, 9619, 9474, 9508, 9569, 9570, _
            9558, 9557, 9571, 9553, 9559, 9565, 9564, 9563, _
            9488, 9492, 9524, 9516, 9500, 9472, 9532, 9566, _
            9567, 9562, 9556, 9577, 9574, 9568, 9552, 9580, _
            9575, 9576, 9572, 9573, 9561, 9560, 9554, 9555, _
            9579, 9578, 9496, 9484, 9608, 9604, 9612, 9616, _
            9600, 945, 223, 915, 960, 931, 963, 181, _
            964, 934, 920, 937, 948, 8734, 966, 949, _
            8745, 8801, 177, 8805, 8804, 8992, 8993, 247, _
            8776, 176, 8729, 183, 8730, 8319, 178, 9632, _
            160 _
        }
#End Region

#Region "Enumerations"
    Public Enum ModifierKey As Integer
        Control = 0
        Shift
        Alt
    End Enum
#End Region

#Region "Subclasses"

#Region "Keyboard"
    ''' <summary>
    ''' A static class for handling and simulating physical keyboard input.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Keyboard
        Private Sub New()
        End Sub

#Region "Public methods"

#Region "PressKey()"
        ''' <summary>
        ''' Virtually presses a key.
        ''' </summary>
        ''' <param name="Key">The key to press.</param>
        ''' <param name="HardwareKey">Whether or not to press the key using its hardware scan code.</param>
        ''' <remarks></remarks>
        Public Shared Sub PressKey(ByVal Key As Keys, Optional ByVal HardwareKey As Boolean = False)
            Keyboard.SetKeyState(Key, True, HardwareKey)
            Keyboard.SetKeyState(Key, False, HardwareKey)
        End Sub
#End Region

#Region "SetKeyState()"
        ''' <summary>
        ''' Virtually sends a key event.
        ''' </summary>
        ''' <param name="Key">The key of the event.</param>
        ''' <param name="KeyDown">Whether to push down or release the key.</param>
        ''' <param name="HardwareKey">Whether the event should use the key's virtual key code or its hardware scan code.</param>
        ''' <remarks></remarks>
        Public Shared Sub SetKeyState(ByVal Key As Keys, ByVal KeyDown As Boolean, Optional ByVal HardwareKey As Boolean = False)
            Dim InputList As New List(Of NativeMethods.INPUT)
            Dim Modifiers As Keys() = InputHelper.ExtractModifiers(Key)

            For Each Modifier As Keys In Modifiers
                InputList.Add(Keyboard.GetKeyboardInputStructure(Modifier, KeyDown, HardwareKey))
            Next
            InputList.Add(Keyboard.GetKeyboardInputStructure(Key, KeyDown, True))

            NativeMethods.SendInput(CType(InputList.Count, UInteger), InputList.ToArray(), Marshal.SizeOf(GetType(NativeMethods.INPUT)))
        End Sub
#End Region

#Region "IsKeyDown()"
        ''' <summary>
        ''' Checks whether a key is currently held down.
        ''' </summary>
        ''' <param name="Key">The key to check.</param>
        ''' <remarks></remarks>
        Public Shared Function IsKeyDown(ByVal Key As Keys) As Boolean
            Dim Modifiers As Keys() = InputHelper.ExtractModifiers(Key)
            For Each Modifier As Keys In Modifiers
                If (NativeMethods.GetAsyncKeyState(Modifier) And InputHelper.KeyDownBit) <> InputHelper.KeyDownBit Then
                    Return False
                End If
            Next

            If Key = Keys.None Then Return True 'All modifiers are held down, no more keys left to check.

            Return (NativeMethods.GetAsyncKeyState(Key) And InputHelper.KeyDownBit) = InputHelper.KeyDownBit
        End Function
#End Region

#Region "IsKeyUp()"
        ''' <summary>
        ''' Checks whether a key is up (returns True for any key that isn't currently held down).
        ''' </summary>
        ''' <param name="Key">The key to check.</param>
        ''' <remarks></remarks>
        Public Shared Function IsKeyUp(ByVal Key As Keys) As Boolean
            Return Keyboard.IsKeyDown(Key) = False
        End Function
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
#End Region

#Region "Mouse"
    ''' <summary>
    ''' A static class for handling and simulating physical mouse input.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Mouse
        Private Sub New()
        End Sub

#Region "Public methods"

#Region "PressButton()"
        ''' <summary>
        ''' Virtually presses a mouse button.
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
        ''' Virtually sends a mouse button event.
        ''' </summary>
        ''' <param name="Button">The button of the event.</param>
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
#End Region

#Region "WindowMessages"
    ''' <summary>
    ''' A static class for handling and simulating input via Window Messages.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class WindowMessages
        Private Sub New()
        End Sub

#Region "Public methods"

#Region "SendKeyPress()"
        ''' <summary>
        ''' Sends a keystroke message to the active window.
        ''' </summary>
        ''' <param name="Key">The key to send.</param>
        ''' <param name="HardwareKey">Whether to send the hardware scan code along with the virtual key code (recommended to be 'True'!).</param>
        ''' <param name="SendAsynchronously">Whether or not to wait for the window to handle the message before continuing.</param>
        ''' <remarks></remarks>
        Public Shared Sub SendKeyPress(ByVal Key As Keys, Optional ByVal HardwareKey As Boolean = True, Optional ByVal SendAsynchronously As Boolean = False)
            WindowMessages.SendKey(Key, True, HardwareKey, SendAsynchronously)
            WindowMessages.SendKey(Key, False, HardwareKey, SendAsynchronously)
        End Sub

        ''' <summary>
        ''' Sends a keystroke message to a window.
        ''' </summary>
        ''' <param name="WindowHandle">The handle of the window to send the keystroke message to.</param>
        ''' <param name="Key">The key to send.</param>
        ''' <param name="HardwareKey">Whether to send the hardware scan code along with the virtual key code (recommended to be 'True'!).</param>
        ''' <param name="SendAsynchronously">Whether or not to wait for the window to handle the message before continuing.</param>
        ''' <remarks></remarks>
        Public Shared Sub SendKeyPress(ByVal WindowHandle As IntPtr, ByVal Key As Keys, Optional ByVal HardwareKey As Boolean = True, Optional ByVal SendAsynchronously As Boolean = False)
            WindowMessages.SendKey(WindowHandle, Key, True, HardwareKey, SendAsynchronously)
            WindowMessages.SendKey(WindowHandle, Key, False, HardwareKey, SendAsynchronously)
        End Sub
#End Region

#Region "SendKey()"
        ''' <summary>
        ''' Sends a KeyDown/KeyUp message to the active window.
        ''' </summary>
        ''' <param name="Key">The key to send.</param>
        ''' <param name="KeyDown">Whether to send a KeyDown or KeyUp message.</param>
        ''' <param name="HardwareKey">Whether to send the hardware scan code along with the virtual key code (recommended to be 'True'!).</param>
        ''' <param name="SendAsynchronously">Whether or not to wait for the window to handle the message before continuing.</param>
        ''' <remarks></remarks>
        Public Shared Sub SendKey(ByVal Key As Keys, ByVal KeyDown As Boolean, Optional ByVal HardwareKey As Boolean = True, Optional ByVal SendAsynchronously As Boolean = False)
            Dim ActiveWindow As IntPtr = InputHelper.WindowMessages.GetAbsoluteActiveWindow()
            WindowMessages.SendKey(ActiveWindow, Key, KeyDown, HardwareKey, SendAsynchronously)
        End Sub

        ''' <summary>
        ''' Sends a KeyDown/KeyUp message to a window.
        ''' </summary>
        ''' <param name="WindowHandle">The handle of the window to send the key message to.</param>
        ''' <param name="Key">The key to send.</param>
        ''' <param name="KeyDown">Whether to send a KeyDown or KeyUp message.</param>
        ''' <param name="HardwareKey">Whether to send the hardware scan code along with the virtual key code (recommended to be 'True'!).</param>
        ''' <param name="SendAsynchronously">Whether or not to wait for the window to handle the message before continuing.</param>
        ''' <remarks></remarks>
        Public Shared Sub SendKey(ByVal WindowHandle As IntPtr, ByVal Key As Keys, ByVal KeyDown As Boolean, Optional ByVal HardwareKey As Boolean = True, Optional ByVal SendAsynchronously As Boolean = False)
            Dim Modifiers As Keys() = InputHelper.ExtractModifiers(Key)

            'Send the modifiers first.
            For Each Modifier As Keys In Modifiers
                WindowMessages.InternalSendKey(WindowHandle, Modifier, KeyDown, HardwareKey, SendAsynchronously)
            Next

            If Key = Keys.None Then Return 'We only sent modifiers.

            WindowMessages.InternalSendKey(WindowHandle, Key, KeyDown, HardwareKey, SendAsynchronously)
        End Sub
#End Region

#Region "SendAltCode()"
        ''' <summary>
        ''' Sends an ALT code message to the active window.
        ''' </summary>
        ''' <param name="NumpadKeys">The numpad combination to send.</param>
        ''' <remarks></remarks>
        Public Shared Sub SendAltCode(ByVal NumpadKeys As Keys())
            Dim ActiveWindow As IntPtr = WindowMessages.GetAbsoluteActiveWindow()
            WindowMessages.SendAltCode(ActiveWindow, NumpadKeys)
        End Sub

        ''' <summary>
        ''' Sends an ALT code message to a window.
        ''' </summary>
        ''' <param name="WindowHandle">The handle of the window to send the ALT code message to.</param>
        ''' <param name="NumpadKeys">The numpad combination to send.</param>
        ''' <remarks></remarks>
        Public Shared Sub SendAltCode(ByVal WindowHandle As IntPtr, ByVal NumpadKeys As Keys())
            If NumpadKeys.Any(Function(k As Keys) k < Keys.NumPad0 OrElse k > Keys.NumPad9) = True Then _
                Throw New ArgumentOutOfRangeException("NumpadKeys", "Only Keys.NumPad0 through Keys.NumPad9 are allowed to be passed to this method.")

            Dim LoopTo As Integer = If(NumpadKeys.Length > 4, 4, NumpadKeys.Length) - 1

            If NumpadKeys(0) = Keys.NumPad0 Then 'Windows-1252 ALT code.

                WindowMessages.SendKey(WindowHandle, Keys.Alt, True)
                For x = 0 To LoopTo
                    WindowMessages.SendKeyPress(WindowHandle, NumpadKeys(x))
                Next
                WindowMessages.SendKey(WindowHandle, Keys.Alt, False)

            Else 'Codepage 437 ALT code.

                WindowMessages.SendKey(WindowHandle, Keys.Alt, True)
                For x = 0 To LoopTo
                    WindowMessages.SendKey(WindowHandle, NumpadKeys(x), True)
                    If x = LoopTo Then
                        WindowMessages.SendKey(WindowHandle, Keys.Alt, False) 'Release ALT before sending KeyUp of the last key.
                    End If
                    WindowMessages.SendKey(WindowHandle, NumpadKeys(x), False)
                Next

            End If
        End Sub
#End Region

#Region "ToggleKey()"
        ''' <summary>
        ''' Toggles a specific key in InputHelper's internal keyboard state (not related to the physical keyboard!). Use only with toggleable keys!
        ''' </summary>
        ''' <param name="Key">The key to toggle (e.g. Keys.CapsLock, Keys.NumLock, etc.).</param>
        ''' <param name="SetToggled">Whether to mark the key as toggled or untoggled (Nothing/Null = toggle automatically).</param>
        ''' <remarks></remarks>
        Public Shared Sub ToggleKey(ByVal Key As Keys, Optional ByVal SetToggled As Nullable(Of Boolean) = Nothing)
            Dim Modifiers As Keys() = InputHelper.ExtractModifiers(Key)

            If SetToggled Is Nothing Then 'Toggle the key.
                For Each Modifier As Keys In Modifiers
                    Dim ModifierIndex As Integer = CType(Modifier, Integer) And 255 'Cap index between 0-255.
                    InputHelper.KeyboardState(ModifierIndex) = InputHelper.KeyboardState(ModifierIndex) Xor (1 << 0) 'Toggle the least significant bit.
                Next

                If Key = Keys.None Then Return 'We toggled modifiers only.

                Dim Index As Integer = CType(Key, Integer) And 255 'Cap index between 0-255.
                InputHelper.KeyboardState(Index) = InputHelper.KeyboardState(Index) Xor (1 << 0) 'Toggle the least significant bit.

            Else 'Set the key to a specific value (toggled/untoggled)
                For Each Modifier As Keys In Modifiers
                    Dim ModifierIndex As Integer = CType(Modifier, Integer) And 255 'Cap index between 0-255.
                    If SetToggled = True Then
                        InputHelper.KeyboardState(ModifierIndex) = InputHelper.KeyboardState(ModifierIndex) Or (1 << 0) 'Set the least significant bit.
                    Else
                        InputHelper.KeyboardState(ModifierIndex) = InputHelper.KeyboardState(ModifierIndex) And Not (1 << 0) 'Unset the least significant bit.
                    End If
                Next

                If Key = Keys.None Then Return 'We set modifiers only.

                Dim Index As Integer = CType(Key, Integer) And 255 'Cap index between 0-255.
                If SetToggled = True Then
                    InputHelper.KeyboardState(Index) = InputHelper.KeyboardState(Index) Or (1 << 0) 'Set the least significant bit.
                Else
                    InputHelper.KeyboardState(Index) = InputHelper.KeyboardState(Index) And Not (1 << 0) 'Unset the least significant bit.
                End If

            End If
        End Sub
#End Region

#Region "IsKeyToggled()"
        ''' <summary>
        ''' Checks whether a specific key is toggled or not in InputHelper's internal keyboard state (not related to the physical keyboard!).
        ''' </summary>
        ''' <param name="Key">The key to check (e.g. Keys.CapsLock, Keys.NumLock, etc.).</param>
        ''' <remarks></remarks>
        Public Shared Function IsKeyToggled(ByVal Key As Keys) As Boolean
            If InputHelper.ExtractModifiers(Key).Length > 0 Then _
                Throw New ArgumentOutOfRangeException("Key", "Keys.Control, Keys.Shift and Keys.Alt is not valid in this context." & Environment.NewLine & _
                                                             "Use Keys.ControlKey, Keys.ShiftKey and Keys.Menu instead.")
            Return (InputHelper.KeyboardState(CType(Key, Integer) And 255) And (1 << 0)) = (1 << 0)
        End Function
#End Region

#Region "IsKeyDown()"
        ''' <summary>
        ''' Checks whether a specific key is down in InputHelper's internal keyboard state (not related to the physical keyboard!).
        ''' </summary>
        ''' <param name="Key"></param>
        ''' <remarks></remarks>
        Public Shared Function IsKeyDown(ByVal Key As Keys) As Boolean
            If InputHelper.ExtractModifiers(Key).Length > 0 Then _
                Throw New ArgumentOutOfRangeException("Key", "Keys.Control, Keys.Shift and Keys.Alt is not valid in this context." & Environment.NewLine & _
                                                             "Use Keys.ControlKey, Keys.ShiftKey and Keys.Menu instead.")
            Return (InputHelper.KeyboardState(CType(Key, Integer) And 255) And (1 << 7)) = (1 << 7)
        End Function
#End Region

#Region "IsModifierDown()"
        ''' <summary>
        ''' Checks whether any Left or Right version of a modifier is down in InputHelper's internal keyboard state (not related to the physical keyboard!).
        ''' </summary>
        ''' <param name="Modifier">The modifier to check.</param>
        ''' <remarks></remarks>
        Public Shared Function IsModifierDown(ByVal Modifier As ModifierKey) As Boolean
            Select Case Modifier
                Case ModifierKey.Control
                    Return WindowMessages.IsKeyDown(Keys.ControlKey) OrElse _
                           WindowMessages.IsKeyDown(Keys.LControlKey) OrElse _
                           WindowMessages.IsKeyDown(Keys.RControlKey)
                Case ModifierKey.Shift
                    Return WindowMessages.IsKeyDown(Keys.ShiftKey) OrElse _
                           WindowMessages.IsKeyDown(Keys.LShiftKey) OrElse _
                           WindowMessages.IsKeyDown(Keys.RShiftKey)
                Case ModifierKey.Alt
                    Return WindowMessages.IsKeyDown(Keys.Menu) OrElse _
                           WindowMessages.IsKeyDown(Keys.LMenu) OrElse _
                           WindowMessages.IsKeyDown(Keys.RMenu)
            End Select
            Throw New ArgumentOutOfRangeException("Modifier", CType(Modifier, Integer) & " is not a valid modifier key!")
        End Function
#End Region

#Region "SendMouseClick()"
        ''' <summary>
        ''' Sends a Window Message-based mouse click to the window at the specified coordinates of the screen.
        ''' </summary>
        ''' <param name="Button">The button to press.</param>
        ''' <param name="Location">The position where to send the click (in screen coordinates).</param>
        ''' <remarks></remarks>
        Public Shared Sub SendMouseClick(ByVal Button As MouseButtons, ByVal Location As Point)
            WindowMessages.SendMouseClick(Button, Location, True, False)
            WindowMessages.SendMouseClick(Button, Location, False, False)
        End Sub

        ''' <summary>
        ''' Sends a Window Message-based mouse click to the window at the specified coordinates of the screen.
        ''' </summary>
        ''' <param name="Button">The button to press.</param>
        ''' <param name="Location">The position where to send the click (in screen coordinates).</param>
        ''' <param name="MouseDown">Whether to push down or release the mouse button.</param>
        ''' <param name="SendAsynchronously">Whether or not to wait for the window to handle the message before continuing.</param>
        ''' <remarks></remarks>
        Public Shared Sub SendMouseClick(ByVal Button As MouseButtons, ByVal Location As Point, ByVal MouseDown As Boolean, Optional ByVal SendAsynchronously As Boolean = False)
            Dim hWnd As IntPtr = NativeMethods.WindowFromPoint(New NativeMethods.NATIVEPOINT(Location.X, Location.Y)) 'Get the window at the specified click point.
            Dim ButtonMessage As NativeMethods.MouseButtonMessage 'A variable holding which Window Message to use.

            Select Case Button 'Set the appropriate mouse button Window Message.
                Case MouseButtons.Left : ButtonMessage = If(MouseDown, NativeMethods.MouseButtonMessage.WM_LBUTTONDOWN, NativeMethods.MouseButtonMessage.WM_LBUTTONUP)
                Case MouseButtons.Right : ButtonMessage = If(MouseDown, NativeMethods.MouseButtonMessage.WM_RBUTTONDOWN, NativeMethods.MouseButtonMessage.WM_RBUTTONUP)
                Case MouseButtons.Middle : ButtonMessage = If(MouseDown, NativeMethods.MouseButtonMessage.WM_MBUTTONDOWN, NativeMethods.MouseButtonMessage.WM_MBUTTONUP)
                Case MouseButtons.XButton1, MouseButtons.XButton2
                    ButtonMessage = NativeMethods.MouseButtonMessage.WM_XBUTTONDOWN
                Case Else
                    Throw New ArgumentException("Invalid mouse button " & Button.ToString() & "!", "Button")
            End Select

            Dim ClickPoint As New NativeMethods.NATIVEPOINT(Location.X, Location.Y) 'Create a native point.

            If NativeMethods.ScreenToClient(hWnd, ClickPoint) = False Then 'Convert the click point to client coordinates relative to the window.
                Throw New Exception("Unable to convert screen coordinates to client coordinates! Win32Err: " & _
                                        Marshal.GetLastWin32Error())
            End If

            Dim wParam As IntPtr = IntPtr.Zero 'Used to specify which X button was clicked (if any).
            Dim lParam As IntPtr = WindowMessages.CreateLWParam(ClickPoint.X, ClickPoint.Y) 'Click point.

            If Button = MouseButtons.XButton1 OrElse _
                Button = MouseButtons.XButton2 Then
                wParam = WindowMessages.CreateLWParam(0, Button / MouseButtons.XButton1) 'Set the correct XButton.
            End If

            InputHelper.WindowMessages.SendMessage(hWnd, ButtonMessage, wParam, lParam, SendAsynchronously)
        End Sub
#End Region

#Region "GetActiveWindow()"
        ''' <summary>
        ''' Gets the active top-level window.
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Function GetActiveWindow() As IntPtr
            Return NativeMethods.GetForegroundWindow()
        End Function
#End Region

#Region "GetAbsoluteActiveWindow()"
        ''' <summary>
        ''' Gets the (absolute) active top-level window or child window (apart from GetActiveWindow() which will only get the active top-level window).
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Function GetAbsoluteActiveWindow() As IntPtr
            Dim CurrentThreadID As UInteger = NativeMethods.GetCurrentThreadId()
            Dim ActiveWindow As IntPtr = NativeMethods.GetForegroundWindow()
            Dim ActiveThread As UInteger = NativeMethods.GetWindowThreadProcessId(ActiveWindow, Nothing)

            If ActiveThread = 0 Then _
                Return ActiveWindow

            If NativeMethods.AttachThreadInput(CurrentThreadID, ActiveThread, True) = False Then _
                Return ActiveWindow

            Dim AbsoluteActiveWindow As IntPtr = NativeMethods.GetFocus()
            Dim DetachAttempts As Integer = 0

            While NativeMethods.AttachThreadInput(CurrentThreadID, ActiveThread, False) = False
                DetachAttempts += 1
                If DetachAttempts >= 10 Then Exit While
                Threading.Thread.Sleep(1)
            End While

            Return AbsoluteActiveWindow
        End Function
#End Region

#End Region

#Region "Internal methods"

#Region "InternalSendKey()"
        ''' <summary>
        ''' Internal method for sending KeyDown/KeyUp messages to a specific window (DOES NOT HANDLE MODIFIERS! USE 'SendKey()' INSTEAD!).
        ''' </summary>
        ''' <param name="WindowHandle">The handle to the window to send the keystrokes to.</param>
        ''' <param name="Key">The key to send.</param>
        ''' <param name="KeyDown">Whether to send a KeyDown or KeyUp message.</param>
        ''' <param name="HardwareKey">Whether to send the hardware scan code along with the virtual key code (recommended to be 'True'!).</param>
        ''' <param name="SendAsynchronously">Whether or not to wait for the window to handle the message before continuing.</param>
        ''' <remarks></remarks>
        Private Shared Sub InternalSendKey(ByVal WindowHandle As IntPtr, ByVal Key As Keys, ByVal KeyDown As Boolean, ByVal HardwareKey As Boolean, ByVal SendAsynchronously As Boolean)
            Dim KeyboardLayout As IntPtr = NativeMethods.GetKeyboardLayout(0)
            Dim Message As NativeMethods.KeyMessage = WindowMessages.GetKeyWindowMessage(Key, KeyDown)
            Dim lParam As IntPtr = WindowMessages.MakeKeyLParam(Key, KeyDown, HardwareKey)
            Dim KeyCode As UInteger = CType(Key And UInteger.MaxValue, UInteger)
            Dim ScanCode As UInteger = If(HardwareKey, NativeMethods.MapVirtualKeyEx(KeyCode, 0, KeyboardLayout), 0)

            Dim CharBuffer As New StringBuilder(16)
            Dim CharCode As UInteger = NativeMethods.MapVirtualKeyEx(KeyCode, 2, KeyboardLayout)
            Dim CharacterConversionResult As Integer = -1

            Dim IsDeadChar As Boolean = (CharCode And (1 << 31)) = (CType(1, UInteger) << 31) 'If the most significant bit is set the key is a dead char.
            Dim IsAltDown As Boolean = WindowMessages.IsModifierDown(ModifierKey.Alt)
            Dim IsAlt As Boolean = InputHelper.IsModifier(Key, ModifierKey.Alt)
            Dim AltUp As Boolean = (IsAlt = True AndAlso IsAltDown = True AndAlso KeyDown = False)
            Dim IsNumpadNumber As Boolean = (KeyCode >= CType(Keys.NumPad0, UInteger) AndAlso KeyCode <= CType(Keys.NumPad9, UInteger))

            'Is this an ALT+numpad combination?
            If InputHelper.IsAltCodeCombination = False AndAlso IsAltDown = True AndAlso IsNumpadNumber = True Then
                InputHelper.IsAltCodeCombination = True 'Begin an ALT+numpad combination.
            End If

            'Check the key's state.
            If KeyDown = True Then
                InputHelper.KeyboardState(CType(Key, Integer)) = (1 << 7) 'Set the most significant bit if the key is down.
            Else
                InputHelper.KeyboardState(CType(Key, Integer)) = 0 'Unset value if key is up.
            End If

            'Must be done AFTER the key state has been changed.
            If KeyDown = True AndAlso InputHelper.IsAltCodeCombination = False Then
                CharacterConversionResult = _
                    NativeMethods.ToUnicodeEx(KeyCode, ScanCode, InputHelper.KeyboardState, CharBuffer, CharBuffer.Capacity, 0, KeyboardLayout)
            End If

            'Verify conversion result.
            If CharacterConversionResult = 0 Then 'Conversion failed. Try MapVirtualKeyEx() instead.
                CharBuffer.Clear()

                Dim MappedCharCode As UInteger = NativeMethods.MapVirtualKeyEx(KeyCode, 2, KeyboardLayout) And Not (1 << 31)
                If MappedCharCode <> 0 Then CharBuffer.Append(Convert.ToChar(MappedCharCode))

            ElseIf CharacterConversionResult > 0 Then 'Conversion succeded.
                CharBuffer.Remove(CharacterConversionResult, CharBuffer.Length - CharacterConversionResult) 'Truncate CharBuffer to the amount of chars we want.

            End If

            'Send the actual key press.
            WindowMessages.SendMessage(WindowHandle, Message, New IntPtr(Key), lParam, SendAsynchronously)



            'Is this the end of an ALT+numpad combination?
            If KeyDown = False AndAlso IsAlt = True AndAlso InputHelper.IsAltCodeCombination = True AndAlso AltCode.Count > 0 Then
                InputHelper.IsAltCodeCombination = False 'ALT is being released, end ALT+numpad combination.

                Dim NumberString As String = ""
                For x = 0 To InputHelper.AltCode.Count - 1
                    If InputHelper.AltCode(x) = Keys.NumPad0 AndAlso x = 0 Then Continue For 'Skip the first key if it's a zero.
                    NumberString &= Convert.ToChar(NativeMethods.MapVirtualKeyEx(InputHelper.AltCode(x), 2, KeyboardLayout)) 'Convert each key into a digit character.
                Next

                Dim AltCodeNumber As Byte = 0 'The ALT code can only range from 0-255.

                If Byte.TryParse(NumberString, AltCodeNumber) = True Then

                    If InputHelper.AltCode(0) = Keys.NumPad0 Then 'Windows-1252 ALT codes start with zero.
                        Dim DecodedBytes As Byte() = Encoding.Convert(InputHelper.Windows1252, Encoding.UTF8, New Byte() {AltCodeNumber})
                        Dim DecodedChars As Char() = Encoding.UTF8.GetChars(DecodedBytes)

                        If DecodedChars.Length >= 1 Then 'Send the resulting char.
                            WindowMessages.SendMessage(WindowHandle, NativeMethods.KeyMessage.WM_CHAR, New IntPtr(Convert.ToInt32(DecodedChars(0))), lParam, SendAsynchronously)
                        End If

                    Else 'Codepage 437 ALT code.
                        WindowMessages.SendMessage(WindowHandle, NativeMethods.KeyMessage.WM_CHAR, New IntPtr(InputHelper.CP437LookupTable(AltCodeNumber)), lParam, SendAsynchronously)
                    End If

                End If

                AltCode.Clear() 'Clear the combination list.
            End If



            'Shall we send a WM_CHAR/WM_DEADCHAR message or is this an ALT+numpad combination?
            If KeyDown = True AndAlso InputHelper.IsAltCodeCombination = False Then
                If IsDeadChar = False Then
                    Dim CharMessage As NativeMethods.KeyMessage = If(IsAltDown, NativeMethods.KeyMessage.WM_SYSCHAR, NativeMethods.KeyMessage.WM_CHAR)

                    'Send WM_CHAR or WM_SYSCHAR.
                    For Each c As Char In CharBuffer.ToString()
                        WindowMessages.SendMessage(WindowHandle, CharMessage, New IntPtr(Convert.ToInt32(c)), lParam, SendAsynchronously)
                    Next
                Else
                    Dim CharMessage As NativeMethods.KeyMessage = If(IsAltDown, NativeMethods.KeyMessage.WM_SYSDEADCHAR, NativeMethods.KeyMessage.WM_DEADCHAR)

                    'Send WM_DEADCHAR or WM_SYSDEADCHAR for dead characters (ex: `, ´, ^, ¨, ~, etc.).
                    WindowMessages.SendMessage(WindowHandle, NativeMethods.KeyMessage.WM_DEADCHAR, New IntPtr(CharCode And Not (1 << 31)), lParam, SendAsynchronously)
                End If

            ElseIf KeyDown = True AndAlso InputHelper.IsAltCodeCombination = True Then
                If IsNumpadNumber = True Then
                    InputHelper.AltCode.Add(Key) 'Add numpad number to ALT+numpad combination.
                Else
                    InputHelper.AltCode.Clear() 'Invalid char, reset combination.
                End If

            End If
        End Sub
#End Region

#Region "SendMessage()"
        ''' <summary>
        ''' Sends or posts a message to a window's message queue.
        ''' </summary>
        ''' <param name="hWnd">The handle of the window to send the message to.</param>
        ''' <param name="Msg">The window message to send.</param>
        ''' <param name="wParam">The wParam of the message.</param>
        ''' <param name="lParam">The lParam of the message.</param>
        ''' <param name="Asynchronous">Whether to post or send the message to the window's message queue.</param>
        ''' <remarks></remarks>
        Private Shared Sub SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByVal Asynchronous As Boolean)
            If Asynchronous = True Then
                NativeMethods.PostMessage(hWnd, Msg, wParam, lParam)
            Else
                NativeMethods.SendMessage(hWnd, Msg, wParam, lParam)
            End If
        End Sub
#End Region

#Region "MakeKeyLParam()"
        ''' <summary>
        ''' Constructs the lParam parameter of a key message.
        ''' </summary>
        ''' <param name="Key">The key of the message.</param>
        ''' <param name="KeyDown">Whether this is a KeyDown or KeyUp message.</param>
        ''' <param name="HardwareKey">Whether to include the key's hardware scan code in the lParam parameter (recommended to be 'True'!).</param>
        ''' <remarks></remarks>
        Private Shared Function MakeKeyLParam(ByVal Key As Keys, ByVal KeyDown As Boolean, ByVal HardwareKey As Boolean) As IntPtr
            Dim ScanCode As UInteger = NativeMethods.MapVirtualKeyEx(CType(Key And UInteger.MaxValue, UInteger), 0, NativeMethods.GetKeyboardLayout(0))
            Dim lParam As Integer = (1 << 0) Or If(HardwareKey, (ScanCode And Integer.MaxValue), 0) << 16 'Set bit 16-23 to the Scan code if HardwareKey = True.

            Dim IsAlt As Boolean = ( _
                (Key And Keys.Alt) = Keys.Alt _
                OrElse Key = Keys.Menu _
                OrElse Key = Keys.LMenu _
                OrElse Key = Keys.RMenu _
            )

            Dim SystemKey As Boolean = ( _
                ( _
                    IsAlt _
                    OrElse Key = Keys.F10 _
                    OrElse WindowMessages.IsKeyDown(Keys.Menu) _
                    OrElse WindowMessages.IsKeyDown(Keys.LMenu) _
                    OrElse WindowMessages.IsKeyDown(Keys.RMenu) _
                ) _
                AndAlso Not (IsAlt = True AndAlso KeyDown = False) _
            )

            Dim IsExtendedKey As Boolean = ( _
                Key = Keys.LControlKey OrElse _
                Key = Keys.RControlKey OrElse _
                Key = Keys.LShiftKey OrElse _
                Key = Keys.RShiftKey OrElse _
                Key = Keys.LMenu OrElse _
                Key = Keys.RMenu _
            )

            If KeyDown = False Then
                'Set bit 30 and 31 to '1' when sending WM_KEYUP or WM_SYSKEYUP:
                'https://msdn.microsoft.com/en-us/library/windows/desktop/ms646281(v=vs.85).aspx
                lParam = (lParam Or (1 << 30)) Or (1 << 31)
            End If

            If SystemKey = True Then
                'Set bit 29 to '1', indicates that ALT is held down:
                'https://msdn.microsoft.com/en-us/library/windows/desktop/ms646286(v=vs.85).aspx
                lParam = lParam Or (1 << 29)
            End If

            If IsExtendedKey = True Then
                'Set bit 24 to '1' if this is an extended key:
                'https://msdn.microsoft.com/en-us/library/windows/desktop/ms646280(v=vs.85).aspx
                lParam = lParam Or (1 << 24)
            End If

            Return New IntPtr(lParam)
        End Function
#End Region

#Region "GetKeyWindowMessage()"
        ''' <summary>
        ''' Gets the window message type of the specified key.
        ''' </summary>
        ''' <param name="Key">The key to use for the message.</param>
        ''' <param name="KeyDown">Whether this is a KeyDown or KeyUp message.</param>
        ''' <remarks></remarks>
        Private Shared Function GetKeyWindowMessage(ByVal Key As Keys, ByVal KeyDown As Boolean) As NativeMethods.KeyMessage
            Dim IsAlt As Boolean = ( _
                (Key And Keys.Alt) = Keys.Alt _
                OrElse Key = Keys.Menu _
                OrElse Key = Keys.LMenu _
                OrElse Key = Keys.RMenu _
            )

            Dim SystemKey As Boolean = ( _
                ( _
                    IsAlt _
                    OrElse Key = Keys.F10 _
                    OrElse WindowMessages.IsKeyDown(Keys.Menu) _
                    OrElse WindowMessages.IsKeyDown(Keys.LMenu) _
                    OrElse WindowMessages.IsKeyDown(Keys.RMenu) _
                ) _
                AndAlso Not (IsAlt = True AndAlso KeyDown = False)
            )

            If SystemKey = True Then _
                Return If(KeyDown, NativeMethods.KeyMessage.WM_SYSKEYDOWN, NativeMethods.KeyMessage.WM_SYSKEYUP)

            Return If(KeyDown, NativeMethods.KeyMessage.WM_KEYDOWN, NativeMethods.KeyMessage.WM_KEYUP)
        End Function
#End Region

#Region "CreateLWParam()"
        ''' <summary>
        ''' Creates a wParam or lParam value.
        ''' </summary>
        ''' <param name="LoWord">The low word (UShort) of the lParam/wParam.</param>
        ''' <param name="HiWord">The high word (UShort) of the lParam/wParam.</param>
        ''' <remarks></remarks>
        Private Shared Function CreateLWParam(ByVal LoWord As Integer, ByVal HiWord As Integer) As IntPtr
            Return New IntPtr((HiWord << 16) Or (LoWord And &HFFFF))
        End Function
#End Region

#End Region

    End Class
#End Region

#End Region

#Region "Internal methods"

#Region "ExtractModifiers()"
    ''' <summary>
    ''' Extracts any .NET modifiers from the specified key combination and returns them as native virtual key code keys.
    ''' </summary>
    ''' <param name="Key">The key combination to extract the modifiers from (if any).</param>
    ''' <remarks></remarks>
    Private Shared Function ExtractModifiers(ByRef Key As Keys) As Keys()
        Dim Modifiers As New List(Of Keys)

        If (Key And Keys.Control) = Keys.Control Then
            Key = Key And Not Keys.Control
            Modifiers.Add(Keys.ControlKey)
        End If

        If (Key And Keys.Shift) = Keys.Shift Then
            Key = Key And Not Keys.Shift
            Modifiers.Add(Keys.ShiftKey)
        End If

        If (Key And Keys.Alt) = Keys.Alt Then
            Key = Key And Not Keys.Alt
            Modifiers.Add(Keys.Menu)
        End If

        Return Modifiers.ToArray()
    End Function
#End Region

#Region "IsModifier()"
    ''' <summary>
    ''' Checks whether the specified key is any Left or Right version of the specified modifier.
    ''' </summary>
    ''' <param name="Key">The key to check.</param>
    ''' <param name="Modifier">The modifier to check for.</param>
    ''' <remarks></remarks>
    Private Shared Function IsModifier(ByVal Key As Keys, ByVal Modifier As ModifierKey) As Boolean
        Select Case Modifier
            Case ModifierKey.Control
                Return _
                    Key = Keys.Control OrElse _
                    Key = Keys.ControlKey OrElse _
                    Key = Keys.LControlKey OrElse _
                    Key = Keys.RControlKey
            Case ModifierKey.Shift
                Return _
                    Key = Keys.Shift OrElse _
                    Key = Keys.ShiftKey OrElse _
                    Key = Keys.LShiftKey OrElse _
                    Key = Keys.RShiftKey
        End Select
        Return _
            Key = Keys.Alt OrElse _
            Key = Keys.Menu OrElse _
            Key = Keys.LMenu OrElse _
            Key = Keys.RMenu
    End Function
#End Region

#End Region

#Region "WinAPI P/Invokes"
    Private NotInheritable Class NativeMethods
        Private Sub New()
        End Sub

#Region "Methods"

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
#End Region

#End Region

#Region "Enumerations"
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

        Public Enum MouseButtonMessage As Integer
            WM_LBUTTONDOWN = &H201
            WM_LBUTTONUP = &H202
            WM_MBUTTONDOWN = &H207
            WM_MBUTTONUP = &H208
            WM_RBUTTONDOWN = &H204
            WM_RBUTTONUP = &H205
            WM_XBUTTONDOWN = &H20B
            WM_XBUTTONUP = &H20C
            XBUTTON1 = &H1
            XBUTTON2 = &H2
        End Enum
#End Region

#Region "Structures"
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
            Public X As Integer
            Public Y As Integer

            Public Sub New(ByVal X As Integer, ByVal Y As Integer)
                Me.X = X
                Me.Y = Y
            End Sub
        End Structure
#End Region

    End Class
#End Region

End Class