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
Imports InputHelper.Hooks

Namespace EventArgs
    Public Class KeyboardHookEventArgs
        Inherits System.EventArgs

        Private _extended As Boolean
        Private _keyCode As Keys
        Private _keyState As KeyState
        Private _modifiers As ModifierKeys
        Private _scanCode As UInteger

        ''' <summary>
        ''' Gets or sets a boolean value indicating whether the keystroke should be blocked from reaching any windows.
        ''' </summary>
        ''' <remarks></remarks>
        Public Property Block As Boolean = False

        ''' <summary>
        ''' Gets a boolean value indicating whether the keystroke message originated from one of the additional keys on the enhanced keyboard 
        ''' (see: https://msdn.microsoft.com/en-us/library/windows/desktop/ms646267(v=vs.85).aspx#_win32_Keystroke_Message_Flags).
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Extended As Boolean
            Get
                Return _extended
            End Get
        End Property

        ''' <summary>
        ''' Gets the keyboard code of the key that generated the keystroke.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property KeyCode As Keys
            Get
                Return _keyCode
            End Get
        End Property

        ''' <summary>
        ''' Gets the current state of the key that generated the keystroke.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property KeyState As KeyState
            Get
                Return _keyState
            End Get
        End Property

        ''' <summary>
        ''' Gets the modifier keys that was pressed in combination with the keystroke.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Modifiers As ModifierKeys
            Get
                Return _modifiers
            End Get
        End Property

        ''' <summary>
        ''' Gets the hardware scan code of the key that generated the keystroke.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property ScanCode As UInteger
            Get
                Return _scanCode
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return String.Format("{{KeyCode: {0}, ScanCode: {1}, Extended: {2}, KeyState: {3}, Modifiers: {4}}}", _
                                 Me.KeyCode, Me.ScanCode, Me.Extended, Me.KeyState, Me.Modifiers)
        End Function

        ''' <summary>
        ''' Initializes a new instance of the KeyboardHookEventArgs class.
        ''' </summary>
        ''' <param name="KeyCode">The keyboard code of the key that generated the keystroke.</param>
        ''' <param name="ScanCode">The hardware scan code of the key that generated the keystroke.</param>
        ''' <param name="Extended">Whether the keystroke message originated from one of the additional keys on the enhanced keyboard 
        ''' (see: https://msdn.microsoft.com/en-us/library/windows/desktop/ms646267(v=vs.85).aspx#_win32_Keystroke_Message_Flags). </param>
        ''' <param name="KeyState">The current state of the key that generated the keystroke.</param>
        ''' <param name="Modifiers">The modifier keys that was pressed in combination with the keystroke.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal KeyCode As Keys, _
                       ByVal ScanCode As UInteger, _
                       ByVal Extended As Boolean, _
                       ByVal KeyState As KeyState, _
                       ByVal Modifiers As ModifierKeys)
            Me._keyCode = KeyCode
            Me._scanCode = ScanCode
            Me._extended = Extended
            Me._keyState = KeyState
            Me._modifiers = Modifiers
        End Sub
    End Class
End Namespace