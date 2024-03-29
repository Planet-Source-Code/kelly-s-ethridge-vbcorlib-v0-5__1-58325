VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryWriter"
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
'    Module: BinaryWriter
'

''
' Provides methods to write typed information to a stream.
'
' <h4>Remarks</h4>
' <p>The datatypes that are writting to the stream are written in their binary
' form. The value values as they are represented in memory. For example, as datatype
' of vbLong is written as 4 bytes, regardless of the actual value of the vbLong
' variable.</p>
'
Option Explicit
Implements cObject

Private Const VT_BYREF As Long = &H4000


Private mStream As Stream
Private mEncoding As Encoding
Private mIsOpen As Boolean
Private mBuffer() As Byte
Private mBufferSA As SafeArray1d
Private mBytes() As Byte


''
' Returns the underlying stream this writer is writing to.
'
' @return The underlying stream.
' @remarks The stream is flushed before it is returned.
'
Public Property Get BaseStream() As Stream
    Flush
    Set BaseStream = mStream
End Property

''
' Closes the writer, flushing any data.
'
Public Sub CloseWriter()
    If mIsOpen Then
        Flush
        mStream.CloseStream
        mIsOpen = False
    End If
End Sub

''
' Tells the underlying stream to store its data to its repository.
'
Public Sub Flush()
    mStream.Flush
End Sub

''
' Sets the underlying stream's current position.
'
' @param Offset The number of bytes to move the position.
' @param Origin The starting point to move from in the stream.
' @return The final position in the stream.
'
Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    mStream.SeekPosition Offset, Origin
End Function

''
' Writes the binary representation of a datatype to the underlying stream.
'
' @param value The datatype to be written.
' @param vIndex The starting index of arrays or string characters to being writing from.
' @param vCount The number of elements or characters to be written to the stream.
' @remarks <i>vIndex</i> and <i>vCount</i> are only valid when writing Byte arrays,
' Integer array, and Strings. For all other types, these parameters are ignored.
'<p>Simple datatypes such as vbLong and vbDouble are written out exactly as they
' appear in memory. Each individual byte of that datatype is written. For example, a vbLong
' is 4 bytes in memory, even if its value is 0. The 4 bytes are written to the stream, not
' one byte with a value of 0. A vbDouble is 8 bytes, so all 8 bytes are written as is to
' the stream.</p>
' <p>The one exception is when writing a String value. Both the number of bytes from an encoded
' string and the encoded bytes are written, however, the length is not written as the individual bytes, like the
' other datatypes. The length of the string is written as a series of 7bit values. Each byte
' for the length is calculated by starting out with the actual value, then ANDing it with &h3f
' to take the first 7 bits. The value is ORed with &h80 to signify there are more bytes to
' follow for the length. Once that byte is written, the original value is shifted right
' 7 bits and the process starts over until the value reaches less than &h80. The final byte
' value is not ORed with &h80, it is written as is without the high bit set. When reading
' the bytes back, the byte without the high bit set will signal as the last byte in the value.</p><br>
' <font face="Courier New">
'    Do While value > &H80
'        WriteValue CByte((value And &H7F) Or &H80)
'        value = value \ &H80
'    Loop
'    WriteValue CByte(value)
' </font><br>
' The characters of the string are written as the encoded bytes generated by the current Encoding.<br>
' <p>Byte arrays are written as is. There is no description written along with the byte values.
' The number of bytes will have to be known in advance or manually stored in the stream.</p>
' <p>Integer arrays are treated as characters and are converted to the encoded bytes using the
' current encoding. Like the byte array, the number of bytes is not stored with the encoded
' characters and will have to be manually stored.</p>
' <br>
' <p>The Decimal datatype is converted to 16 bytes with the same
' layout as used in .NET. The layout is different than that of VB.<br><br>
' VB Decimal layout<br>
' bytes 0-1: Variant datatype information<br>
' bytes 2: precision<br>
' bytes 3: sign (&h80 is negative)<br>
' bytes 4-7: the 32 highest bits in the 96bit value<br>
' bytes 8-11: the 32 lowest bits in the 96bit value<br>
' bytes 12-15: the 32 middle bits in the 96bit value<br>
' <br>
' .NET Decimal layout<br>
' bytes 0-3: the 32 lowest bits in the 96bit value<br>
' bytes 4-7: the 32 middle bits in the 96bit value<br>
' bytes 8-11: the 32 highest bits in the 96bit value<br>
' bytes 12-13: unused (zero)<br>
' bytes 14: precision<br>
' bytes 15: sign (&h80 is negative)<br>
'
Public Sub WriteValue(ByRef Value As Variant, Optional ByRef vIndex As Variant, Optional ByRef vCount As Variant)
    If Not mIsOpen Then _
        Throw Cor.NewObjectDisposedException("BinaryWriter", "Cannot write to a closed BinaryWriter.")
    
    If IsArray(Value) Then
        WriteArray Value, vIndex, vCount
    Else
        WriteIntrinsic Value, vIndex, vCount
    End If
End Sub


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
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
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
Friend Sub Init(ByVal Stream As Stream, ByVal Encoding As Encoding)
    If Stream Is Nothing Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Stream), "Stream")
    If Not Stream.CanWrite Then _
        Throw Cor.NewArgumentException("Cannot write to a stream that does not support writing.", "Stream")
    
    Set mStream = Stream
    If Encoding Is Nothing Then
        Set mEncoding = Cor.NewUTF8Encoding(False, True)
    Else
        Set mEncoding = Encoding
    End If
    mIsOpen = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Private Sub WriteStringLength(ByVal Value As Long)
    Do While Value > &H80
        WriteValue CByte((Value And &H7F) Or &H80)
        Value = Value \ &H80
    Loop
    WriteValue CByte(Value)
End Sub

Private Sub WriteIntrinsic(ByRef Value As Variant, ByRef vIndex As Variant, ByRef vCount As Variant)
    Dim num As Long
    Dim Index As Long
    Dim Count As Long

    Select Case VarType(Value) And &HFF
        Case vbBoolean:                     mStream.WriteByte IIf(Value, 1, 0)
        Case vbByte:                        mStream.WriteByte Value
        Case vbLong, vbSingle:              num = 4
        Case vbInteger:                     num = 2
        Case vbDouble, vbDate, vbCurrency:  num = 8
        Case vbString
            WriteStringLength mEncoding.GetByteCount(Value, vIndex, vCount)
            WriteValue mEncoding.GetBytes(Value, vIndex, vCount)
        Case vbDecimal
            mStream.WriteBlock BitConverter.GetBytes(Value), 0, 16
        Case Else
            Throw Cor.NewArgumentException("Can only write intrinsice datatypes to the stream.", "Value")
    End Select
    If num > 0 Then
        mBufferSA.pvData = VarPtr(Value) + 8
        If VariantType(Value) And VT_BYREF Then mBufferSA.pvData = MemLong(mBufferSA.pvData)
        mStream.WriteBlock mBuffer, 0, num
    End If
End Sub

Private Sub WriteArray(ByRef Value As Variant, ByRef vIndex As Variant, ByRef vCount As Variant)
    Dim Index As Long
    Dim Count As Long
    
    If cArray.IsNull(Value) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Value")
        
    Select Case VarType(Value) And &HFF
        Case vbByte
            SAPtr(mBytes) = cArray.GetValidArrayPointer(Value)
            AssignByteIndexAndcount mBytes, vIndex, Index, vCount, Count
            mStream.WriteBlock mBytes, Index, Count
        Case vbInteger
            ' we call back into WriteValue passing a byte array of the
            ' encoded integer values instead.
            WriteValue mEncoding.GetBytes(Value, vIndex, vCount)
        Case Else
            Throw Cor.NewArgumentException("Can only write Byte and Integer arrays to the stream.", "Value")
    End Select
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    With mBufferSA
        .cbElements = 1
        .cDims = 1
        .cElements = &H10
    End With
    SAPtr(mBuffer) = VarPtr(mBufferSA)
End Sub

Private Sub Class_Terminate()
    SAPtr(mBuffer) = 0
    SAPtr(mBytes) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    cObject_Equals = Equals(Value)
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function
