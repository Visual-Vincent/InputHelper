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

''' <summary>
''' A class holding internal methods and fields related to InputHelper.
''' </summary>
''' <remarks></remarks>
Friend Class Internal

#Region "ExtractModifiers()"
    ''' <summary>
    ''' Extracts any .NET modifiers from the specified key combination and returns them as native virtual key code keys.
    ''' </summary>
    ''' <param name="Key">The key combination to extract the modifiers from (if any).</param>
    ''' <remarks></remarks>
    Public Shared Function ExtractModifiers(ByRef Key As Keys) As Keys()
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
    Public Shared Function IsModifier(ByVal Key As Keys, ByVal Modifier As ModifierKeys) As Boolean
        Select Case Modifier
            Case ModifierKeys.Control
                Return _
                    Key = Keys.Control OrElse _
                    Key = Keys.ControlKey OrElse _
                    Key = Keys.LControlKey OrElse _
                    Key = Keys.RControlKey
            Case ModifierKeys.Shift
                Return _
                    Key = Keys.Shift OrElse _
                    Key = Keys.ShiftKey OrElse _
                    Key = Keys.LShiftKey OrElse _
                    Key = Keys.RShiftKey
            Case ModifierKeys.Alt
                Return _
                    Key = Keys.Alt OrElse _
                    Key = Keys.Menu OrElse _
                    Key = Keys.LMenu OrElse _
                    Key = Keys.RMenu
            Case ModifierKeys.Windows
                Return _
                    Key = Keys.LWin OrElse _
                    Key = Keys.RWin
        End Select
        Throw New ArgumentOutOfRangeException("Modifier", CType(Modifier, Integer) & " is not a valid modifier key!")
    End Function
#End Region

End Class
