VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: UTF8Encoding
'

''
' A set of functions to be used to convert character arrays to and from byte arrays.
'
' <H4>Remarks</H4> <p>UTF-8 is an encoding scheme that is fully compatible with the existing
' ASCII set of characters (0-127). It also supports Unicode characters, including the
' UTF-16 character set.</p>
' <p>UTF-8 is useful because it will have a size ratio of 1:1 for pure ASCII characters.
' The ratio will increase during the encoding for characters outside ASCII.</p>
' <br><br>
' <br><H4>Constructors:</H4><br>
' <font face="Courier New">
' Set obj = New UTF8Encoding<br>
' Set obj = NewUTF8Encoding([ShouldEmitUTF8Identifier], [ShouldThrowException])<br>
' Set obj = Cor.NewUTF8Encoding([ShouldEmitUTF8Identifier], [ShouldThrowException])<br>
' </font><br>
' <b>ShouldEmitUTF8Identifier</b>: If <b>True</b> then GetPreamble returns an array containing the 3 byte identifier,
' otherwise an empty array is returned (not a null array).<br>
' <b>ShouldThrowException</b>: If <b>True</b> then exceptions will be thrown when data cannot be encoded
' or decoded, otherwise, the bad data will be encoded as is.
'
Option Explicit
Implements cObject
Implements Encoding

Private Const CODE_PAGE As Long = 65001

Private mThrowException As Boolean
Private mShouldEmitUTF8Identifier As Boolean
Private mChars() As Integer
Private mCharsSA As SafeArray1d


''
' Returns the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
'
Public Property Get BodyName() As String
    BodyName = "utf-8"
End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = "utf-8"
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = "utf-8"
End Property

''
' Returns the code page identifier for this encoding.
'
' @return Code page identifier.
'
Public Property Get CodePage() As Long
    CodePage = CODE_PAGE
End Property

''
' Returns the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
'
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = 1200
End Property

''
' The name of this encoding.
'
' @return The encoding's name.
'
Public Property Get EncodingName() As String
    EncodingName = "Unicode (UTF-8)"
End Property

''
' Indicates if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
'
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = True
End Property

''
' Indicates if this encoding can be used to save data with this encoding.
'
' @return The indication for the browser saving capable.
'
Public Property Get IsBrowserSave() As Boolean
    IsBrowserSave = True
End Property

''
' Indicates if this encoding can be used to display mail and news by
' mail and news clients.
'
' @return Indication for mail and news client capable.
'
Public Property Get IsMailNewsDisplay() As Boolean
    IsMailNewsDisplay = True
End Property

''
' Indicates if this encoding can be used to save date by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = True
End Property

''
' Returns the number of bytes that would be produced from the set of characters using this encoding.
'
' @param Chars A set of characters used to calculated the number of bytes once encoded with this encoding.
' @param vIndex The starting index of the first character to be used.
' @param vCount The number of characters to include in the calculation.
' @return The number of bytes that would be produced by encoding the set of characters.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array. The vIndex
' is zero based for both the String and Integer array, not 1 based like Mid$.
'
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Long
    Dim Index As Long
    Dim Count As Long
    AttachChars Chars, vIndex, Index, vCount, Count
    GetByteCount = InternalGetByteCount(mChars, Index, Count, Nothing)
End Function

''
' Encodes a set of characters into an array of bytes.
'
' @param Chars The set of characters to be encoded into an array of bytes. This can
' be either an Integer array or a String.
' @param vIndex The index of the first character to begin encoding from.
' @param vCount The number of characters to be encoded.
' @return A byte array containing the encoded characters as bytes.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array. The vIndex
' is zero based for both the String and Integer array, not 1 based like Mid$.
'
Public Function GetBytes(ByRef Chars As Variant, Optional ByVal vIndex As Variant, Optional ByVal vCount As Variant) As Byte()
    Dim Index As Long
    Dim Count As Long
    Dim ret() As Byte
    
    AttachChars Chars, vIndex, Index, vCount, Count
    ReDim ret(0 To InternalGetByteCount(mChars, Index, Count, Nothing) - 1)
    InternalGetBytes mChars, Index, Count, ret, 0, Nothing
    GetBytes = ret
End Function

''
' Encodes a set of characters into an array of bytes, returning the number of bytes produced.
'
' @param Chars the set of characters to be encoded into bytes. This can
' be either an Integer array or a String.
' @param CharIndex The index of the first character to begin encoding from.
' @param CharCount The number of characters to be encoded.
' @param Bytes The destination for the encoded characters.
' @param ByteIndex The index of the first byte stored in the destination array.
' @return The number of bytes produce by the encoding.
' @remarks The <i>Bytes</i> array cannot be null and must be large enough to hold the encoded
' set of characters. To obtain the size required, use GetByteCount to find the necessary size.
' Or, use GetMaxByteCount to get the largest size that could be produced. The largest size is not
' the perfect size, only the maximum number of bytes that could ever be produced by a number of
' characters. Resizing of the resultant <i>Bytes</i> array may be necessary.
'
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    AttachChars Chars, CharIndex, 0, CharCount, 0
    GetBytesEx = InternalGetBytes(mChars, CharIndex, CharCount, Bytes, ByteIndex, Nothing)
End Function

''
' Returns the number of characters that would be produced by decoding a byte array.
'
' @param Bytes The bytes to use when calculating the number of characters.
' @param vIndex The starting index in <i>Bytes</i> to being calculating from.
' @param vCount The number of bytes to be used in the calculation.
' @return The number of characters that would be decoded from the set of bytes.
'
Public Function GetCharCount(ByRef Bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Long
    Dim Index As Long
    Dim Count As Long
    Dim nullch() As Integer
    
    AssignByteIndexAndcount Bytes, vIndex, Index, vCount, Count
    GetCharCount = InternalGetChars(Bytes, Index, Count, nullch, 0, Nothing)
End Function

''
' Decodes a set of bytes into a set of characters.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param vIndex The index of the first byte to be used in decoding.
' @param vCount The number of bytes to be used in decoding.
' @return An array of characters decoded from the set of bytes.
'
Public Function GetChars(ByRef Bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Integer()
    Dim Index As Long
    Dim Count As Long
    Dim ret() As Integer
    Dim size As Long
    
    size = GetMaxCharCount(cArray.GetLength(Bytes))
    AssignByteIndexAndcount Bytes, vIndex, Index, vCount, Count
    ReDim ret(0 To size - 1)
    size = InternalGetChars(Bytes, Index, Count, ret, 0, Nothing)
    ReDim Preserve ret(0 To size - 1)
    GetChars = ret
End Function

''
' Decodes a set of bytes into the supplied Integer array.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param ByteIndex The index of the first byte to begin decoding from.
' @param ByteCount The number of bytes to be used in decoding.
' @param Chars The destination character array for the decoded bytes.
' @param CharIndex The first index to begin storing decoded bytes in <i>Chars</i>.
' @return The number of characters decoded from the array of bytes.
' @remarks The <i>Chars</i> array must be large enough to handle all the bytes that will
' be decoded. To ensure the <i>Chars</i> array is large enough, use either GetCharCount or
' GetMaxCharCount to determine a size that will hold the decoded bytes.
'
Public Function GetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    GetCharsEx = InternalGetChars(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Nothing)
End Function

''
' Decodes a set of bytes into a String.
'
' @param Bytes The set of bytes to be decoded into a string.
' @param vIndex The index of the first byte to be decoded.
' @param vCount The number of bytes to be used in the decoding.
' @return A string containing the decoded set of bytes.
'
Public Function GetString(ByRef Bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As String
    Dim ret As String
    Dim size As Long
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndcount Bytes, vIndex, Index, vCount, Count
    size = GetMaxCharCount(Count)
    ret = SysAllocStringLen(0, size)
    With mCharsSA
        .pvData = StrPtr(ret)
        .lLbound = 0
        .cElements = size
    End With
    SAPtr(mChars) = VarPtr(mCharsSA)
    size = InternalGetChars(Bytes, Index, Count, mChars, 0, Nothing)
    GetString = left$(ret, size)
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks <p>If the encoding was created setting the <i>ShouldEmitUTF8Identifier</i> paramter
' to true, then this will return the UTF-8 identifer &HFFEF already encoded into the 3
' bytes. The bytes are &HEF, &HBB, &HBF. They will decode to &HFFEF. Since &HFF and &HFE
' are not valid in a UTF-8 encoding, these will unique and can be used to identify the
' encoding used in the stream.</p>
' If <i>ShouldEmitUTF8Identifier</i> was set to False, then an empty array will be returned.
' The array will not be null. It will have a UBound of -1.
'
Public Function GetPreamble() As Byte()
    If mShouldEmitUTF8Identifier Then
        GetPreamble = cArray.NewArray(ciByte, &HEF, &HBB, &HBF)
    Else
        GetPreamble = cArray.CreateInstance(ciByte)     ' creates a 0-length array instead of a null array.
    End If
End Function

''
' Returns the maximum number of bytes that can be created from a specific number of characters.
'
' @param CharCount The number of characters to be encoded.
' @return The maximum number of bytes that can be generated from the specified number of characters.
'
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long
    If CharCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "CharCount"), "CharCount", CharCount)
    GetMaxByteCount = CharCount * 4
End Function

''
' Returns the maximum number of characters than can be decoded from the number of bytes specified.
'
' @param ByteCount The number of bytes to be decoded.
' @return The maximum number of characters that can be decoded from the specified number of bytes.
'
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long
    If ByteCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "ByteCount"), "ByteCount", ByteCount)
    GetMaxCharCount = ByteCount
End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
' @remarks Unlike UTF8Encoding, the encoder maintains state between encoding calls. This allows
' for a set of characters to be split up but still be encoded as a single set of characters. This
' will only be effective when encoding UTF-16 surrogate pairs of characters. Otherwise, a single
' characters is all that is required to be encodable. the UTF-16 surrogate pairs can be split to
' the last character of one call and the first character of the next and the encoding will
' continue as if a single call was made.
'
Public Function GetEncoder() As Encoder
    Dim ret As New UTF8Encoder
    ret.Init Me
    Set GetEncoder = ret
End Function

''
' Returns a decoder that maintains state.
'
' @return The stateful decoder.
' @remarks Unlike UTF8Encoding, the decoder maintains state between decoding calls. this allows
' for a sequence of bytes to not have to be in the same call in order to be decoded to the
' representive character. Other than the standard ASCII set of characters (0-127), all other
' characters are represented by 2-4 bytes. Using this decoder allows for some of those bytes to be
' at the end of one call and the rest be at the beginning of the next call. The decoder will
' continue decoding as if all the required bytes were in the same call.
'
Public Function GetDecoder() As Decoder
    Dim ret As New UTF8Decoder
    ret.Init Me
    Set GetDecoder = ret
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef value As Variant) As Boolean
    If IsObject(value) Then
        If TypeOf value Is UTF8Encoding Then
            Dim e As UTF8Encoding
            Set e = value
            Equals = (mShouldEmitUTF8Identifier = e.EmitUTF8Identifier)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = CODE_PAGE + Abs(mShouldEmitUTF8Identifier) + Abs(mThrowException)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal ShouldEmitUTF8Identifier As Long, ByVal ShouldThrowException As Boolean)
    mShouldEmitUTF8Identifier = ShouldEmitUTF8Identifier
    mThrowException = ShouldThrowException
End Sub

Friend Property Get EmitUTF8Identifier() As Boolean
    EmitUTF8Identifier = mShouldEmitUTF8Identifier
End Property

Friend Function InternalGetBytes(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByRef Encoder As UTF8Encoder) As Long
    Const HIGH_OFFSET As Long = &HD800&
    Const LOW_OFFSET As Long = &HDC00&
    
    Dim uch As Long
    Dim maxindex As Long
    Dim maxByteIndex As Long
    Dim StartIndex As Long
    Dim HighSurrogateChar As Long
    
    modPublicFunctions.Encoding.GetBytesArgumentCheck Chars, CharIndex, CharCount, Bytes, ByteIndex
    
    If Not Encoder Is Nothing Then HighSurrogateChar = Encoder.HighSurrogateChar
    StartIndex = ByteIndex
    maxByteIndex = UBound(Bytes) + 1
    maxindex = CharIndex + CharCount
    Do While CharIndex < maxindex
        AsWord(uch) = Chars(CharIndex)
        If HighSurrogateChar <> 0 Then
            Select Case uch
                Case &HDC00& To &HDFFF&
                    ' we have a surrogate pair.
                    Dim pair As Long
                    If ByteIndex + 4 > maxByteIndex Then SmallBufferError "Bytes"
                    
                    pair = ((HighSurrogateChar - HIGH_OFFSET) * &H400) + (uch - LOW_OFFSET) + &H10000
                    Bytes(ByteIndex + 3) = &H80 Or (pair And &H3F)
                    Bytes(ByteIndex + 2) = &H80 Or ((pair \ &H40&) And &H3F)
                    Bytes(ByteIndex + 1) = &H80 Or ((pair \ &H1000&) And &H3F)
                    Bytes(ByteIndex) = &HF0 Or (pair \ &H40000)
                    ByteIndex = ByteIndex + 4
                    HighSurrogateChar = 0
                Case &HD800& To &HDBFF&
                    ' we have two high surrogates in a row
                    If mThrowException Then Throw Cor.NewArgumentException("A high surrogate was followed by a second high surrogate.", "Chars")
                    If ByteIndex + 3 > maxByteIndex Then SmallBufferError "Bytes"
                    
                    ByteIndex = EncodeSurrogate(HighSurrogateChar, Bytes, ByteIndex)
                    HighSurrogateChar = uch
                Case Else
                    If mThrowException Then Throw Cor.NewArgumentException("A high surrogate was not followed by a low surrogate.", "Chars")
                    If ByteIndex + 3 >= maxByteIndex Then SmallBufferError "Bytes"
                    
                    ByteIndex = EncodeSurrogate(HighSurrogateChar, Bytes, ByteIndex)
                    CharIndex = CharIndex - 1
                    HighSurrogateChar = 0
            End Select
        Else

            Select Case uch
                Case Is < &H80     ' ascii
                    If ByteIndex + 1 > maxByteIndex Then SmallBufferError "Bytes"
                    Bytes(ByteIndex) = uch
                    ByteIndex = ByteIndex + 1
                Case Is < &H800
                    If ByteIndex + 2 > maxByteIndex Then SmallBufferError "Bytes"
                    Bytes(ByteIndex) = &HC0 Or (uch \ 64)
                    Bytes(ByteIndex + 1) = &H80 Or (uch And &H3F)
                    ByteIndex = ByteIndex + 2
                Case &HD800& To &HDBFF&     ' high surrogate
                    HighSurrogateChar = uch
                Case &HDC00& To &HDFFF&     ' low surrogate
                    If mThrowException Then Throw Cor.NewArgumentException("A low surrogate was not preceeded by a high surrogate.", "Chars")
                    If ByteIndex + 3 > maxByteIndex Then SmallBufferError "Bytes"
                    ByteIndex = EncodeSurrogate(uch, Bytes, ByteIndex)
                Case Else
                    If ByteIndex + 3 > maxByteIndex Then SmallBufferError "Bytes"
                    ByteIndex = EncodeSurrogate(uch, Bytes, ByteIndex)
            End Select
        End If
        CharIndex = CharIndex + 1
    Loop
    
    If Not Encoder Is Nothing Then
        If Not Encoder.Flush Then
            Encoder.HighSurrogateChar = HighSurrogateChar
        Else
            If mThrowException And (HighSurrogateChar <> 0) Then Throw Cor.NewArgumentException("A high surrogate was not followed by a low surrogate.", "Chars")
            Encoder.HighSurrogateChar = 0
        End If
    ElseIf HighSurrogateChar <> 0 Then
        If mThrowException Then Throw Cor.NewArgumentException("A high surrogate was not followed by a low surrogate.", "Chars")
        ByteIndex = EncodeSurrogate(HighSurrogateChar, Bytes, ByteIndex)
    End If
    
    InternalGetBytes = ByteIndex - StartIndex
End Function

Friend Function InternalGetChars(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByRef Decoder As UTF8Decoder) As Long
    Dim BytesInSequence As Long
    Dim b As Byte
    Dim ch As Long
    Dim CharCount As Long
    Dim BytesLeft As Long
    Dim putChars As Boolean
    Dim ub As Long
    Dim lb As Long
    Dim maxindex As Long
    Dim maxCharIndex As Long
    
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    If ByteIndex < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "ByteIndex", ByteIndex)
    If ByteCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "ByteCount"), "ByteCount", ByteCount)
    If ByteIndex + ByteCount - 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "ByteCount")
    
    putChars = Not cArray.IsNull(Chars)
    If putChars Then
        maxCharIndex = UBound(Chars)
        If CharIndex < LBound(Chars) Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "CharIndex", CharIndex)
        If CharIndex > maxCharIndex Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_UBound))
    End If
    If Not Decoder Is Nothing Then
        BytesLeft = Decoder.BytesLeft
        ch = Decoder.Char
        BytesInSequence = Decoder.BytesInSequence
    End If
    
    ub = UBound(Bytes)
    lb = LBound(Bytes)
    maxindex = ByteIndex + ByteCount

    If ByteIndex < lb Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "ByteIndex", ByteIndex)
    If maxindex - 1 > ub Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "ByteCount")

    Do While ByteIndex < maxindex
        b = Bytes(ByteIndex)
        ' only ascii an character will be 1 byte.   0xxxxxxx format
        If b < &H80 Then
            ' we were expecting more bytes to complete the sequence.
            If BytesInSequence > 0 Then ByteStreamError ByteIndex
            If putChars Then
                If CharIndex > maxCharIndex Then SmallBufferError "Chars"
                Chars(CharIndex) = b
                CharIndex = CharIndex + 1
            End If
            BytesLeft = 0
            BytesInSequence = 0
            CharCount = CharCount + 1
        ElseIf BytesInSequence = 0 Then
            ch = b
            ' we count off the number of bits set starting from
            ' the highest bit. No more than 4 should be found.
            ' This will tell us how many bytes need to be put together
            ' to form the output character, or surrogate pair.
            Do While (b And &H80)
                BytesInSequence = BytesInSequence + 1
                b = (b And &H7F) * 2
            Loop
            BytesLeft = BytesInSequence - 1
            Select Case BytesInSequence
                Case 2      ' 110xxxxx format
                    ch = ch And &H1F
                    If ch < 2 Then ByteStreamError ByteIndex
                Case 3      ' 1110xxxx format
                    ch = ch And &HF
                Case 4      ' 11110xxx format
                    ch = ch And &H7
                    If ch <> 0 Then
                        If ch <> &H4 Then ByteStreamError ByteIndex
                    End If
                Case Else
                    ByteStreamError ByteIndex
            End Select
        Else
            ' if we get in here then we must be on bytes 2-4 in a byte sequence,
            ' so they all will look like 10xxxxxx.
            If (b And &HC0) <> &H80 Then ByteStreamError ByteIndex
            ch = ch * &H40
            ch = ch + (b And &H3F)
            BytesLeft = BytesLeft - 1
            Select Case BytesInSequence
                Case 2      ' 110xxxxx 10xxxxxx -  char from &h80 to &h7ff
                    If ch = 0 Then ByteStreamError ByteIndex
                    If BytesLeft = 0 Then
                        If putChars Then
                            If CharIndex > maxCharIndex Then SmallBufferError "Chars"
                            Chars(CharIndex) = ch
                            CharIndex = CharIndex + 1
                        End If
                        BytesInSequence = 0
                        CharCount = CharCount + 1
                    End If
                Case 3      ' 1110xxxx 10xxxxxx 10xxxxxx - char from &h800 to &hffff, including &hd800-&hdfff for bad surrogates
                    If ch = 0 Then ByteStreamError ByteIndex
                    Select Case BytesLeft
                        Case 0
                            If putChars Then
                                If CharIndex > maxCharIndex Then SmallBufferError "Chars"
                                Chars(CharIndex) = AsWord(ch)
                                CharIndex = CharIndex + 1
                            End If
                            BytesInSequence = 0
                            CharCount = CharCount + 1
                        Case 1  ' check for shortest form conformity
                            If ch < &H20 Then ByteStreamError ByteIndex
                    End Select
                Case 4      ' 11110XXX 10XXxxxx 10xxxxxx 10xxxxxx - used for surrogate pairs
                    If BytesLeft = 0 Then
                        Dim hs As Long
                        Dim ls As Long
                        
                        ls = (ch And &H3FF) + &HDC00&
                        hs = (ch \ &H400) + &HD7C0&
                        If ls < &HDC00& Or ls > &HDFFF& Then ByteStreamError ByteIndex
                        If hs < &HD800& Or hs > &HDBFF& Then ByteStreamError ByteIndex
                        If putChars Then
                            If CharIndex + 1 > maxCharIndex Then SmallBufferError "Chars"
                            Chars(CharIndex + 1) = AsWord(ls)   ' AsWord is used to shove the highest bit in
                            Chars(CharIndex) = AsWord(hs)       ' without causing an overflow. The char will
                            CharIndex = CharIndex + 2           ' be negative afterwards.
                        End If
                        CharCount = CharCount + 2
                        BytesInSequence = 0
                    Else
                        ' check for shortest form conformity
                        If ch = 0 Then ByteStreamError ByteIndex
                    End If
                Case Else
                    ByteStreamError ByteIndex
                    BytesInSequence = 0
                    BytesLeft = 0
            End Select
        End If
        ByteIndex = ByteIndex + 1
    Loop
        
    If Not Decoder Is Nothing Then
        Decoder.BytesLeft = BytesLeft
        If BytesLeft = 0 Then
            Decoder.Char = 0
        Else
            Decoder.Char = ch
        End If
        Decoder.BytesInSequence = BytesInSequence
    End If
    InternalGetChars = CharCount
End Function

Friend Function InternalGetByteCount(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef Encoder As UTF8Encoder) As Long
    Dim ret As Long
    Dim HasHighSurrogate As Boolean
    
    If Index < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "vIndex", Index)
    If Index + Count - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "vCount")

    If Not Encoder Is Nothing Then HasHighSurrogate = (Encoder.HighSurrogateChar <> 0)
    
    Do While Count > 0
        Select Case Chars(Index)
            Case 0 To &H7F      ' ascii value
                If HasHighSurrogate Then
                    If mThrowException Then Throw Cor.NewArgumentException("An incorrect low surrogate was found immediately after a high surrogate.", "Chars")
                End If
                ret = ret + 1
            Case &H80 To &H7FF     ' 2 byte UTF-8 encoding
                If HasHighSurrogate Then
                    If mThrowException Then Throw Cor.NewArgumentException("An incorrect low surrogate was found immediately after a high surrogate.", "Chars")
                End If
                ret = ret + 2
            Case &HD800 To &HDBFF    ' high surrogate value
                If HasHighSurrogate Then
                    If mThrowException Then Throw Cor.NewArgumentException("Second high surrogate in a row was found.", "Chars")
                Else
                    HasHighSurrogate = True
                End If
                ret = ret + 3
            Case &HDC00 To &HDFFF   ' low surrogate value
                If HasHighSurrogate Then
                    ret = ret + 1
                    HasHighSurrogate = False
                Else
                    If mThrowException Then Throw Cor.NewArgumentException("Low surrogate was found without a precedding high surrogate.", "Chars")
                    ret = ret + 3
                End If
            Case Else
                ret = ret + 3
        End Select
        Index = Index + 1
        Count = Count - 1
    Loop
    
    If HasHighSurrogate Then
        If Not Encoder Is Nothing Then
            If Encoder.Flush Then
                If mThrowException Then Throw Cor.NewArgumentException("A high surrogate character was not followed by a low surrogate character.", "Chars")
            End If
        Else
            If mThrowException Then Throw Cor.NewArgumentException("Missing Low Surrogage.", "Chars")
        End If
    End If
    InternalGetByteCount = ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SmallBufferError(ByVal ParamName As String)
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), ParamName)
End Sub

Private Function EncodeSurrogate(ByVal Char As Long, ByRef Bytes() As Byte, ByVal Index As Long) As Long
    Bytes(Index) = &HE0 Or ((Char \ &H1000&) And &HF)
    Bytes(Index + 1) = &H80 Or ((Char \ &H40&) And &H3F)
    Bytes(Index + 2) = &H80 Or (Char And &H3F)
    EncodeSurrogate = Index + 3
End Function

Private Sub AttachChars(ByRef Chars As Variant, ByRef vIndex As Variant, ByRef IndexOut As Long, ByRef vCount As Variant, ByRef CountOut As Long)
    Dim lb As Long
    Dim l As Long
    
    Select Case VarType(Chars)
        Case vbString
            l = Len(Chars)
            mCharsSA.cElements = l
            mCharsSA.pvData = StrPtr(Chars)
            mCharsSA.cbElements = 2
            
            SAPtr(mChars) = VarPtr(mCharsSA)
        Case INTEGER_ARRAY
            SAPtr(mChars) = cArray.GetValidArrayPointer(Chars)
            lb = LBound(mChars)
            l = UBound(mChars) - lb + 1
        Case Else
            Throw Cor.NewArgumentException("Chars must be a String or Integer array.", "Chars")
    End Select
    
    If IsMissing(vIndex) Then
        IndexOut = lb
    Else
        IndexOut = vIndex
    End If
    If IsMissing(vCount) Then
        CountOut = l
    Else
        CountOut = vCount
    End If
End Sub

Private Sub ByteStreamError(ByVal Index As Long)
    If Not mThrowException Then Exit Sub
    Throw Cor.NewArgumentException(cString.Format("An invalid byte was found at index {0}.", Index), "Bytes")
End Sub

Private Sub AssignByteIndexAndcount(ByRef Bytes() As Byte, ByRef vIndex As Variant, ByRef Index As Long, ByRef vCount As Variant, ByRef Count As Long)
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
        
    If IsMissing(vIndex) Then
        Index = LBound(Bytes)
    Else
        Index = vIndex
    End If
    If IsMissing(vCount) Then
        Count = cArray.GetLength(Bytes)
    Else
        Count = vCount
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    With mCharsSA
        .cbElements = 2
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mThrowException = .ReadProperty("ThrowException", True)
        mShouldEmitUTF8Identifier = .ReadProperty("EmitBOM", True)
    End With
End Sub

Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub


Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ThrowException", mThrowException
        .WriteProperty "EmitBOM", mShouldEmitUTF8Identifier
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(value As Variant) As Boolean
    cObject_Equals = Equals(value)
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoding Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Encoding_BodyName() As String
    Encoding_BodyName = BodyName
End Property

Private Property Get Encoding_CodePage() As Long
    Encoding_CodePage = CodePage
End Property

Private Property Get Encoding_EncodingName() As String
    Encoding_EncodingName = EncodingName
End Property

Private Function Encoding_Equals(value As Variant) As Boolean
    Encoding_Equals = Equals(value)
End Function

Private Function Encoding_GetByteCount(value As Variant, Optional vIndex As Variant, Optional vCount As Variant) As Long
    Encoding_GetByteCount = GetByteCount(value, vIndex, vCount)
End Function

Private Function Encoding_GetBytes(Source As Variant, Optional vIndex As Variant, Optional vCount As Variant) As Byte()
    Encoding_GetBytes = GetBytes(Source, vIndex, vCount)
End Function

Private Function Encoding_GetBytesEx(Source As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Encoding_GetBytesEx = GetBytesEx(Source, CharIndex, CharCount, Bytes, ByteIndex)
End Function

Private Function Encoding_GetCharCount(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Long
    Encoding_GetCharCount = GetCharCount(Bytes, Index, Count)
End Function

Private Function Encoding_GetChars(Bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As Integer()
    Encoding_GetChars = GetChars(Bytes, vIndex, vCount)
End Function

Private Function Encoding_GetCharsEx(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Encoding_GetCharsEx = GetCharsEx(Bytes, ByteIndex, ByteCount, Chars, CharIndex)
End Function

Private Function Encoding_GetDecoder() As Decoder
    Set Encoding_GetDecoder = GetDecoder
End Function

Private Function Encoding_GetEncoder() As Encoder
    Set Encoding_GetEncoder = GetEncoder
End Function

Private Function Encoding_GetHashCode() As Long
    Encoding_GetHashCode = GetHashCode
End Function

Private Function Encoding_GetMaxByteCount(ByVal CharCount As Long) As Long
    Encoding_GetMaxByteCount = GetMaxByteCount(CharCount)
End Function

Private Function Encoding_GetMaxCharCount(ByVal ByteCount As Long) As Long
    Encoding_GetMaxCharCount = GetMaxCharCount(ByteCount)
End Function

Private Function Encoding_GetPreamble() As Byte()
    Encoding_GetPreamble = GetPreamble
End Function

Private Function Encoding_GetString(Bytes() As Byte, Optional vIndex As Variant, Optional vCount As Variant) As String
    Encoding_GetString = GetString(Bytes, vIndex, vCount)
End Function

Private Property Get Encoding_HeaderName() As String
    Encoding_HeaderName = HeaderName
End Property

Private Property Get Encoding_IsBrowserDisplay() As Boolean
    Encoding_IsBrowserDisplay = IsBrowserDisplay
End Property

Private Property Get Encoding_IsBrowserSave() As Boolean
    Encoding_IsBrowserSave = IsBrowserSave
End Property

Private Property Get Encoding_IsMailNewsDisplay() As Boolean
    Encoding_IsMailNewsDisplay = IsMailNewsDisplay
End Property

Private Property Get Encoding_IsMailNewsSave() As Boolean
    Encoding_IsMailNewsSave = IsMailNewsSave
End Property

Private Function Encoding_ToString() As String
    Encoding_ToString = ToString
End Function

Private Property Get Encoding_WebName() As String
    Encoding_WebName = WebName
End Property

Private Property Get Encoding_WindowsCodePage() As Long
    Encoding_WindowsCodePage = WindowsCodePage
End Property
