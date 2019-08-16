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
    Public Class MouseHookEventArgs
        Inherits System.EventArgs

        Private _button As MouseButtons
        Private _buttonState As KeyState
        Private _delta As Integer
        Private _doubleClick As Boolean
        Private _location As Point
        Private _scrollDirection As ScrollDirection

        ''' <summary>
        ''' Gets or sets a boolean value indicating whether the mouse event should be blocked from reaching any windows.
        ''' </summary>
        ''' <remarks></remarks>
        Public Property Block As Boolean = False

        ''' <summary>
        ''' Gets which mouse button was pressed or released.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Button As MouseButtons
            Get
                Return _button
            End Get
        End Property

        ''' <summary>
        ''' Gets the current state of the button that generated the mouse event.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property ButtonState As KeyState
            Get
                Return _buttonState
            End Get
        End Property

        ''' <summary>
        ''' Gets a signed count of the number of detents the mouse wheel has rotated. A detent is one notch of the mouse wheel.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Delta As Integer
            Get
                Return _delta
            End Get
        End Property

        ''' <summary>
        ''' Gets a boolean value indicating whether the event was caused by a double click.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property DoubleClick As Boolean
            Get
                Return _doubleClick
            End Get
        End Property

        ''' <summary>
        ''' Gets the location of the mouse (in screen coordinates).
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Location As Point
            Get
                Return _location
            End Get
        End Property

        ''' <summary>
        ''' Gets which direction the mouse wheel was scrolled in.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property ScrollDirection As ScrollDirection
            Get
                Return _scrollDirection
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return String.Format("{{Button: {0}, State: {1}, DoubleClick: {2}, Location: {3}, Scroll: {4}, Delta: {5}}}", _
                                 Me.Button, Me.ButtonState, Me.DoubleClick, Me.Location, Me.ScrollDirection, Me.Delta)
        End Function

        ''' <summary>
        ''' Initializes a new instance of the MouseHookEventArgs class.
        ''' </summary>
        ''' <param name="Button">Which mouse button was pressed or released.</param>
        ''' <param name="ButtonState">The current state of the button that generated the mouse event.</param>
        ''' <param name="DoubleClick">Whether the event was caused by a double click.</param>
        ''' <param name="Location">The location of the mouse (in screen coordinates).</param>
        ''' <param name="ScrollDirection">Which direction the mouse wheel was scrolled in.</param>
        ''' <param name="Delta">A signed count of the number of detents the mouse wheel has rotated.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Button As MouseButtons, _
                       ByVal ButtonState As KeyState, _
                       ByVal DoubleClick As Boolean, _
                       ByVal Location As Point, _
                       ByVal ScrollDirection As ScrollDirection, _
                       ByVal Delta As Integer)
            Me._button = Button
            Me._buttonState = ButtonState
            Me._doubleClick = DoubleClick
            Me._location = Location
            Me._scrollDirection = ScrollDirection
            Me._delta = Delta
        End Sub
    End Class
End Namespace