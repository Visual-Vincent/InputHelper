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

Imports System.Text
Imports System.Windows.Forms

''' <summary>
''' A class holding various InputHelper constants.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Constants
    Private Sub New()
    End Sub

    ''' <summary>
    ''' The Alt Gr key.
    ''' </summary>
    ''' <remarks></remarks>
    Public Const AltGr As Keys = Keys.Control Or Keys.Alt

    ''' <summary>
    ''' The bit value used to check if a key is held down (masked with the return value from GetAsyncKeyState()).
    ''' </summary>
    ''' <remarks></remarks>
    Public Const KeyDownBit As Integer = &H8000

    ''' <summary>
    ''' Windows-1252 encoding.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared ReadOnly Windows1252 As Encoding = Encoding.GetEncoding("windows-1252")

End Class
