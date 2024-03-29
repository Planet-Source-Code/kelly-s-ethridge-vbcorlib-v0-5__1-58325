VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: EncodingStatic
'

''
' Provides static methods used to retrieve existing encodings and convert between encodings.
'
Option Explicit

Private Const BASE64_BYTES As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"


Private mUTF8 As UTF8Encoding
Private mUTF7 As UTF7Encoding
Private mASCII As ASCIIEncoding
Private mUnicode As UnicodeEncoding
Private mBigEndianUnicode As UnicodeEncoding



''
' Returns a UnicodeEncoding object that encodes using BigEndian byte ordering.
'
' @return A single instance of UnicodeEncoding with BigEndian encoding.
' @remarks This encoding is declared as Cor.NewUnicodeEncoding(True, True).
'
Public Property Get BigEndianUnicode() As UnicodeEncoding
    If mBigEndianUnicode Is Nothing Then Set mBigEndianUnicode = Cor.NewUnicodeEncoding(True, True)
    Set BigEndianUnicode = mBigEndianUnicode
End Property

''
' Returns a UnicodeEncoding object that encodes using LittleEndian byte ordering.
'
' @return A single instance of UnicodeEncoding with LittleEndian encoding.
' @remarks This encoding is declared as Cor.NewUnicodeEncoding(False, True).
'
Public Property Get Unicode() As UnicodeEncoding
    If mUnicode Is Nothing Then Set mUnicode = Cor.NewUnicodeEncoding(False, True)
    Set Unicode = mUnicode
End Property

''
' Returns a shared ASCIIEncoding object.
'
' @return A single instance of an ASCIIEncoding.
' @remarks This encoding is declared as New ASCIIEncoding.
'
Public Property Get ASCII() As ASCIIEncoding
    If mASCII Is Nothing Then Set mASCII = New ASCIIEncoding
    Set ASCII = mASCII
End Property

''
' Returns a shared UTF7Encoding object.
'
' @return A single instance of a UTF7Encoding object
' @remarks This encoding is declared as New UTF7Encoding.
'
Public Property Get UTF7() As UTF7Encoding
    If mUTF7 Is Nothing Then Set mUTF7 = New UTF7Encoding
    Set UTF7 = mUTF7
End Property

''
' Returns a shared UTF8Encoding object.
'
' @return A single instance of UTF8Encoding object.
' @remarks This encoding is declared as Cor.NewUTF8Encoding(True)
'
Public Property Get UTF8() As UTF8Encoding
    If mUTF8 Is Nothing Then Set mUTF8 = Cor.NewUTF8Encoding(True)
    Set UTF8 = mUTF8
End Property

''
' Converts a set of bytes from one encoding to another encoding.
'
' @param srcEncoding The encoding to be used to decode the bytes.
' @param dstEncoding The encoding to be used to re-encode the decoded characters.
' @param vIndex The starting byte to begin decoding at.
' @param vCount The number of bytes to be included in the decoding process.
' @return The encoded bytes.
'
Public Function Convert(ByVal srcEncoding As Encoding, ByVal dstEncoding As Encoding, ByRef Bytes() As Byte, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant) As Byte()
    Dim ch() As Integer
    Dim Index As Long
    Dim Count As Long
    
    If srcEncoding Is Nothing Then _
        Throw Cor.NewArgumentNullException("Source Encoding cannot be Nothing.", "srcEncoding")
    If dstEncoding Is Nothing Then _
        Throw Cor.NewArgumentNullException("Destination Encoding cannot be Nothing.", "dstEncoding")
    
    AssignByteIndexAndcount Bytes, vIndex, Index, vCount, Count
    ch = srcEncoding.GetChars(Bytes, Index, Count)
    Convert = dstEncoding.GetBytes(ch)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetCharsArgumentCheck(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long)
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    If cArray.IsNull(Chars) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Chars")
    If ByteIndex < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "ByteIndex", ByteIndex)
    If ByteCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "ByteCount"), "ByteCount", ByteCount)
    If ByteIndex + ByteCount - 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "ByteCount")
    If CharIndex < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "CharIndex", CharIndex)
    If CharIndex > UBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_UBound))
End Sub

Friend Sub GetBytesArgumentCheck(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal Count As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long)
    If cArray.IsNull(Chars) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Chars")
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    If CharIndex < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "CharIndex", CharIndex)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Count"), "Count", Count)
    If CharIndex + Count - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    If ByteIndex < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "ByteIndex", ByteIndex)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitEncodingHelpers()
    Dim i As Long
    
    ReDim Base64CharToBits(0 To 127)
    For i = 0 To 127
        Base64CharToBits(i) = -1
    Next i
    For i = 0 To 25
        Base64CharToBits(65 + i) = i
        Base64CharToBits(97 + i) = i + 26
    Next i
    For i = 0 To 9
        Base64CharToBits(48 + i) = i + 52
    Next i
    Base64CharToBits(43) = 62
    Base64CharToBits(47) = 63
    
    ReDim Base64Bytes(63)
    For i = 0 To Len(BASE64_BYTES) - 1
        Base64Bytes(i) = Asc(Mid$(BASE64_BYTES, i + 1, 1))
    Next i
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
        Count = cArray.GetLength(Bytes, 1)
    Else
        Count = vCount
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitEncodingHelpers
End Sub
