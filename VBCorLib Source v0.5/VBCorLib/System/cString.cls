VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: cString
'

''
' Provides static functions to format and manipulate string values.

Option Explicit

Private mBuilder As StringBuilder
Private mString As WordBuffer
Private mTrimChars As WordBuffer


''
' Returns if the string is null.
'
' @param s The string to be tested for null.
' @return Indication if the string is null.
' @remarks If the string is null it will return True. An empty string
' will return false. A null string is equal to StrPtr(s) = 0.
'
Public Function IsNull(ByRef s As String) As Boolean
    IsNull = (StrPtr(s) = 0)
End Function

''
' Returns a string from the characters in an integer array.
'
' @param Chars The array to create the string from.
' @param vIndex The starting index in the array.
' @param vCount The number of characters to use.
' @return A string created from the chars in an integer array.
' @remarks This does minimal error checking.
'
Public Function FromCharArray(ByRef Chars() As Integer, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As String
    Dim Index As Long
    Dim Count As Long
    
    If cArray.IsNull(Chars) Then Exit Function
    If IsMissing(vIndex) Then
        Index = LBound(Chars)
    Else
        Index = vIndex
    End If
    If IsMissing(vCount) Then
        Count = UBound(Chars) - Index + 1
    Else
        Count = vCount
    End If
    FromCharArray = SysAllocStringLen(VarPtr(Chars(Index)), Count)
End Function


''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param fmt The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
'
Public Function Format(ByRef fmt As String, ParamArray args() As Variant) As String
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16)
    Format = FormatArrayEx(Nothing, fmt, vArgs)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param provider A custom format provider that is used to format the argument instead of the default formatting.
' @param fmt The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
'
Public Function FormatEx(ByVal provider As IFormatProvider, ByRef fmt As String, ParamArray args() As Variant) As String
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(20)
    FormatEx = FormatArrayEx(provider, fmt, vArgs)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param fmt The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
'
Public Function FormatArray(ByRef fmt As String, ByRef args() As Variant) As String
    FormatArray = FormatArrayEx(Nothing, fmt, args)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param provider A custom format provider that is used to format the argument instead of the default formatting.
' @param fmt The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
'
Public Function FormatArrayEx(ByVal provider As IFormatProvider, ByRef fmt As String, ByRef args() As Variant) As String
    mBuilder.Length = 0
    mBuilder.InternalAppendFormat provider, fmt, args
    FormatArrayEx = mBuilder.ToString
End Function

''
' Creates and array of chars (Integers) from the specified string.
'
' @param s The string to create the chars from.
' @param startindex The start index in the string to begin converting to a char array.
' @param length The number of characters to convert to chars.
' @return An array containing the converted characters from the string.
'
Public Function ToCharArray(ByRef s As Variant, Optional ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As Integer()
    Dim ret() As Integer
    Dim sl As Long
    
    If VarType(s) <> vbString Then _
        Throw Cor.NewArgumentException("A string is required.", "s")
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "startindex"), "startindex", StartIndex)
    sl = Len(s)
    If Length < 0 Then Length = sl - StartIndex
    If StartIndex + Length > sl Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "length")
    
    If sl > 0 Then
        SAPtr(ret) = SafeArrayCreateVector(vbInteger, 0, Length)
        CopyMemory ret(0), ByVal StrPtr(s) + StartIndex * 2, Length * 2
    Else
        ret = cArray.CreateInstance(ciInteger)
    End If
    ToCharArray = ret
End Function

''
' Pads a string's left side.
'
' @param s The string to be padded.
' @param totalWidth The total length of the final string.
' @param paddingChar The character to pad the left of the string with.
' @return The padded string.
' @remarks The paddingChar parameter can accept either a number or chr$ value.
'
Public Function PadLeft(ByRef s As String, ByVal totalWidth As Long, Optional ByRef paddingChar As Variant = 32) As String
    Dim l As Long
    l = Len(s)
    If totalWidth > l Then
        PadLeft = String$(totalWidth - l, paddingChar) & s
    Else
        PadLeft = s
    End If
End Function

''
' Pads a string's right side.
'
' @param s The string to be padded.
' @param totalWidth The total length of the final string.
' @param paddingChar The character to pad the right of the string with.
' @return The padded string.
' @remarks The paddingChar parameter can accept either a number or chr$ value.
'
Public Function PadRight(ByRef s As String, ByVal totalWidth As Long, Optional ByRef paddingChar As Variant = 32) As String
    Dim l As Long
    l = Len(s)
    If totalWidth > l Then
        PadRight = s & String$(totalWidth - l, paddingChar)
    Else
        PadRight = s
    End If
End Function

''
' Trims the beginning of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's LTrim function, this can trim multiple character types,
' not just spaces.
' <p> The trimChars can be either a string of the characters or an integer array.</p>
'
Public Function TrimStart(ByRef s As String, Optional ByRef trimChars As Variant) As String
    Dim Chars() As Integer
    
    Select Case VarType(trimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(trimChars)
            mTrimChars.SA.cElements = Len(trimChars)
            TrimStart = TrimStartChars(s, mTrimChars.data)
        Case INTEGER_ARRAY
            SAPtr(Chars) = cArray.GetValidArrayPointer(trimChars)
            TrimStart = TrimStartChars(s, Chars)
            SAPtr(Chars) = 0
        Case vbError
            TrimStart = TrimStartDefault(s)
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
End Function

''
' Trims the end of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's RTrim function, this can trim multiple character types,
' not just spaces.
' <p> The trimChars can be either a string of the characters or an integer array.</p>
'
Public Function TrimEnd(ByRef s As String, Optional ByRef trimChars As Variant) As String
    Dim Chars() As Integer
    
    Select Case VarType(trimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(trimChars)
            mTrimChars.SA.cElements = Len(trimChars)
            TrimEnd = TrimEndChars(s, mTrimChars.data)
        Case INTEGER_ARRAY
            SAPtr(Chars) = cArray.GetValidArrayPointer(trimChars)
            TrimEnd = TrimEndChars(s, Chars)
            SAPtr(Chars) = 0
        Case vbError
            TrimEnd = TrimEndDefault(s)
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
End Function

''
' Trims both the beginning and end of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's Trim function, this can trim multiple character types,
' not just spaces.
' <p> The trimChars can be either a string of the characters or an integer array.</p>
'
Public Function Trim(ByRef s As String, Optional ByRef trimChars As Variant) As String
    Dim Chars() As Integer
    
    Select Case VarType(trimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(trimChars)
            mTrimChars.SA.cElements = Len(trimChars)
            Trim = TrimEndChars(TrimStartChars(s, mTrimChars.data), mTrimChars.data)
        Case INTEGER_ARRAY
            SAPtr(Chars) = cArray.GetValidArrayPointer(trimChars)
            Trim = TrimEndChars(TrimStartChars(s, Chars), Chars)
            SAPtr(Chars) = 0
        Case vbError
            Trim = TrimEndDefault(TrimStartDefault(s))
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TrimStartChars(ByRef s As String, ByRef trimChars() As Integer) As String
    Dim i As Long
    Dim lb As Long
    Dim ub As Long
    Dim ch As Integer
    Dim pos As Long
    Dim l As Long
    
    lb = LBound(trimChars)
    ub = UBound(trimChars)
    l = Len(s)
    With mString
        .SA.pvData = StrPtr(s)
        Do While pos < l
            ch = .data(pos)
            For i = lb To ub
                If ch = trimChars(i) Then Exit For
            Next i
            If i > ub Then Exit Do
            pos = pos + 1
        Loop
        If pos = l Then Exit Function
        TrimStartChars = Mid$(s, pos + 1)
    End With
End Function

Private Function TrimStartDefault(ByRef s As String) As String
    Dim i As Long
    
    With mString
        .SA.pvData = StrPtr(s)
        For i = 0 To Len(s) - 1
            Select Case .data(i)
                Case &H20, &HD, &H9, &HA, &HB, &HC, &HA0, &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, &H200B, &H3000, &HFEFF
                Case Else
                    TrimStartDefault = Mid$(s, i + 1)
                    Exit Function
            End Select
        Next i
    End With
End Function

Private Function TrimEndChars(ByRef s As String, ByRef trimChars() As Integer) As String
    Dim i As Long
    Dim lb As Long
    Dim ub As Long
    Dim ch As Integer
    Dim pos As Long
    
    lb = LBound(trimChars)
    ub = UBound(trimChars)
    pos = Len(s) - 1
    With mString
        .SA.pvData = StrPtr(s)
        Do While pos >= 0
            ch = .data(pos)
            For i = lb To ub
                If ch = trimChars(i) Then Exit For
            Next i
            If i > ub Then Exit Do
            pos = pos - 1
        Loop
        If pos < 0 Then Exit Function
        TrimEndChars = left$(s, pos + 1)
    End With
End Function

Private Function TrimEndDefault(ByRef s As String) As String
    Dim i As Long
    
    With mString
        .SA.pvData = StrPtr(s)
        For i = Len(s) - 1 To 0 Step -1
            Select Case .data(i)
                Case &H20, &HD, &H9, &HA, &HB, &HC, &HA0, &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, &H200B, &H3000, &HFEFF
                Case Else
                    TrimEndDefault = left$(s, i + 1)
                    Exit Function
            End Select
        Next i
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBuilder = New StringBuilder
    InitWordBuffer mString, 0, &H7FFFFFFF
    InitWordBuffer mTrimChars, 0, 0
End Sub

