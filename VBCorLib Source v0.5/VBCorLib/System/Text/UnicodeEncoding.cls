VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: UnicodeEncoding
'
''
' A set of functions used to convert unicode characters to and from byte arrays.
'
' <h4>Remarks</h4><p>This encoding can encode characters using big-endian or little-endian
' notation. The default is little-endian (least significant byte first).</p>
' <p>A Byte order Mark will be emitted if selected using the Cor.NewUnicodeEncoding constructor.</p>
'
Option Explicit
Implements cObject
Implements Encoding

Private Const CHAR_SIZE As Long = 2

Private mChars() As Integer
Private mCharsSA As SafeArray1d
Private mIsBigEndian As Boolean
Private mEmitBOM As Boolean
Private mCodePage As Long



''
' Returns the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
'
Public Property Get BodyName() As String
    BodyName = "utf-16"
End Property

''
' Returns the code page identifier for this encoding.
'
' @return Code page identifier.
'
Public Property Get CodePage() As Long
    CodePage = mCodePage
End Property

''
' The name of this encoding.
'
' @return The encoding's name.
'
Public Property Get EncodingName() As String
    Encoding = "Unicode"
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = "utf-16"
End Property

''
' Indicates if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
'
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = False
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
    IsMailNewsDisplay = False
End Property

''
' Indicates if this encoding can be used to save date by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = False
End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = "utf-16"
End Property

''
' Returns the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
'
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = mCodePage
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
    If cArray.IsNull(mChars) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Chars")
    If Index < LBound(mChars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "vIndex", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Count"), "vCount", Count)
    If Index + Count - 1 > UBound(mChars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset))
    
    GetByteCount = Count * CHAR_SIZE
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
Public Function GetBytes(ByRef Chars As Variant, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Byte()
    Dim ret() As Byte
    Dim Index As Long
    Dim Count As Long
    
    AttachChars Chars, vIndex, Index, vCount, Count
    ReDim ret(0 To GetByteCount(Chars, vIndex, vCount) - 1)
    GetBytesEx Chars, Index, Count, ret, 0
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
    modPublicFunctions.Encoding.GetBytesArgumentCheck mChars, CharIndex, CharCount, Bytes, ByteIndex
    If mIsBigEndian Then
        Dim i As Long
        Dim ch As Long
        For i = 0 To CharCount - 1
            AsWord(ch) = mChars(CharIndex + i)
            Bytes(ByteIndex) = ch \ &H100
            Bytes(ByteIndex + 1) = ch And &HFF
            ByteIndex = ByteIndex + 2
        Next i
    Else
        CopyMemory Bytes(ByteIndex), mChars(CharIndex), CharCount * CHAR_SIZE
    End If
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
    
    AssignByteIndexAndcount Bytes, vIndex, Index, vCount, Count
    If Index < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "vIndex", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Count"), "vCount", Count)
    If Index + Count - 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "vCount")
    
    GetCharCount = Count \ CHAR_SIZE
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
    Dim ret() As Integer
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndcount Bytes, vIndex, Index, vCount, Count
    ReDim ret(0 To GetCharCount(Bytes, Index, Count) - 1)
    GetCharsEx Bytes, Index, Count, ret, 0
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
    Dim i As Long
    Dim ch As Long
    
    modPublicFunctions.Encoding.GetCharsArgumentCheck Bytes, ByteIndex, ByteCount, Chars, CharIndex
    If CharIndex + GetMaxCharCount(ByteCount) - 1 > UBound(Chars) Then SmallBufferError
    
    If mIsBigEndian Then
        For i = 0 To ByteCount - 1 Step 2
            ch = Bytes(ByteIndex + i) * &H100
            ch = ch Or Bytes(ByteIndex + i + 1)
            Chars(CharIndex) = AsWord(ch)
            CharIndex = CharIndex + 1
        Next i
    Else
        CopyMemory Chars(CharIndex), Bytes(ByteIndex), ByteCount
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

    GetMaxByteCount = CharCount * CHAR_SIZE
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

    GetMaxCharCount = ByteCount \ 2
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks <p>If the encoding was created setting the <i>BigEndian</i> paramter
' to true, then this will return the Unicode byte array identifer &HFEFF otherwise
' &hFFFE for a non-bigEndian encoding. If no identifier is to be emitted, then an
' empty byte array is returned (not a null array).
'
Public Function GetPreamble() As Byte()
    If mEmitBOM Then
        If mIsBigEndian Then
            GetPreamble = cArray.NewArray(ciByte, &HFF, &HFE)
        Else
            GetPreamble = cArray.NewArray(ciByte, &HFE, &HFF)
        End If
    Else
        GetPreamble = cArray.CreateInstance(ciByte)
    End If
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
    Dim size As Long
    Dim Index As Long
    Dim Count As Long
    
    AssignByteIndexAndcount Bytes, vIndex, Index, vCount, Count
    size = GetCharCount(Bytes, Index, Count)
    GetString = SysAllocStringLen(0, size)
    With mCharsSA
        .pvData = StrPtr(GetString)
        .lLbound = 0
        .cElements = size
    End With
    SAPtr(mChars) = VarPtr(mCharsSA)
    GetCharsEx Bytes, Index, Count, mChars, 0
End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
' @remarks Since encoding Unicode characters equates to exactly 2 bytes per
' character, there is no state that is maintained between calls. This method
' functions identically to GetBytes.
'
Public Function GetEncoder() As Encoder
    Dim ret As New DefaultEncoder
    ret.Init Me
    Set GetEncoder = ret
End Function

''
' Returns a decoder that maintains state.
'
' @return The stateful decoder.
' @remarks Unlike UnicodeEncoding, the decoder maintains state between decoding calls.
' Since a single unicode character required exactly 2 bytes to be decoded, it is possible
' that an odd number of bytes may be attempted to be decoding, creating an orphaned byte
' which represents only half of a unicode character. The remaining byte is held until
' the next decoding call, and the byte is then inserted at the beginning of the next set
' of bytes to be decoded, picking up byte decoding where the orphaned byte left off.
'
Public Function GetDecoder() As Decoder
    Dim ret As New UnicodeDecoder
    ret.Init mIsBigEndian
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
        Equals = (value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal BigEndian As Boolean, ByVal ByteOrderMark As Boolean)
    mIsBigEndian = BigEndian
    mEmitBOM = ByteOrderMark
    If BigEndian Then mCodePage = 1201
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AttachChars(ByRef Chars As Variant, ByRef vIndex As Variant, ByRef IndexOut As Long, ByRef vCount As Variant, ByRef CountOut As Long)
    Dim lb As Long
    Dim l As Long
    
    Select Case VarType(Chars)
        Case vbString
            l = Len(Chars)
            mCharsSA.cElements = l
            mCharsSA.pvData = StrPtr(Chars)
            mCharsSA.cbElements = 2
            mCharsSA.cDims = 1
            
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

Private Sub SmallBufferError()
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mCodePage = 1200
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Init .ReadProperty("BigEndian", False), .ReadProperty("EmitBOM", True)
    End With
End Sub

Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BigEndian", mIsBigEndian
        .WriteProperty "EmitBOM", mEmitBOM
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

Private Function Encoding_GetByteCount(Chars As Variant, Optional vIndex As Variant, Optional vCount As Variant) As Long
    Encoding_GetByteCount = GetByteCount(Chars, vIndex, vCount)
End Function

Private Function Encoding_GetBytes(Chars As Variant, Optional vIndex As Variant, Optional vCount As Variant) As Byte()
    Encoding_GetBytes = GetBytes(Chars, vIndex, vCount)
End Function

Private Function Encoding_GetBytesEx(Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Encoding_GetBytesEx = GetBytesEx(Chars, CharIndex, CharCount, Bytes, ByteIndex)
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

