VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stream"
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
'    Module: Stream
'
Option Explicit

Private Const FILE_BEGIN As Long = 0
Private Const FILE_END As Long = 2
Private Const FILE_CURRENT As Long = 1

''
' The reference point in which to begin seeking from.
'
' @param FromBeginning The new position will be offset from the beginning of the file.
' A negative value will cause an exception.
' @param FromCurrent The position will be offset from the current file pointer location.
' The value can be negative to seek backwards from the current position.
' @param Fromend The new position will be offset from the end of the file. A positive
' number will cause an exception. The offset should be 0 or negative.
'
Public Enum SeekOrigin
    FromBeginning = FILE_BEGIN
    FromCurrent = FILE_CURRENT
    FromEnd = FILE_END
End Enum

''
' Returns if the stream supports reading.
'
' @return The indication for the support of reading.
'
Public Property Get CanRead() As Boolean: End Property

''
' Returns if the stream supports seeking.
'
' @return The indication for the support to seek.
' @remarks Seeking is the ability to change position within the
' stream from a specified reference point.'
'
Public Property Get CanSeek() As Boolean: End Property

''
' Returns if the stream supports being written to.
'
' @return The indication for the support of being written to.
'
Public Property Get CanWrite() As Boolean: End Property

''
' Returns the length of the stream in bytes.
'
' @return The length of the stream.
' @remarks A currency datatype is used to support stream greater than 2gig.<br>
' The stream generally needs to support seeking to be able to return the length.
'
Public Property Get Length() As Currency: End Property

''
' Returns the current position within the stream.
'
' @return The current position within the stream.
' @remarks A currency datatype is used to support stream greater than 2gig.<br>
'
Public Property Get Position() As Currency: End Property

''
' Sets the position in the stream where the next read or write will take place.
'
' @param RHS The position within the stream offset from the beginning of the stream.
' @remarks A currency datatype is used to support stream greater than 2gig.<br>
' The stream generally needs to support seeking to be able to return the length.
'
Public Property Let Position(ByVal RHS As Currency): End Property

''
' Begins an Asynchronous reading process.
'
' @param Buffer The buffer to read into.
' @param Offset The starting index in <i>Buffer</i> to begin reading into.
' @param Count The number of bytes to be read into the buffer.
' @param Callback Object used to receive notification of the finish of the read.
' @param State Information the caller can use to distinguish from other reads.
' @return Represents the current read state.
'
Public Function BeginRead(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: End Function

''
' Begins an Asynchronous writing process.
'
' @param Buffer The source of the data to write to the stream.
' @param Offset The position in the buffer to begin writing from.
' @param Count The number of bytes to be written to the stream.
' @param Callback An object used to receive notification of the completion for the writing.
' @param State A user-supplied information used to distinguish this from other writings.
' @return Represents the current write state.
'
Public Function BeginWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: End Function

''
' Closes the current stream.
'
Public Sub CloseStream(): End Sub

''
' Ends an asynchronous read operation returning the total number of bytes read.
'
' @param AsyncResult The object that represents the read process.
' @return The number of bytes actuall read from the stream.
'
Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long: End Function

''
' Used to end an asynchronous write operation.
'
' @param AsyncResult The object that represents the write process.
'
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult): End Sub

''
' Determines if this instance is equal to the value.
'
' @param value The value to check this instance against for equality.
' @return Indication of equality.
'
Public Function Equals(ByRef value As Variant) As Boolean: End Function

''
' Flushes any buffers in the object to the underlying stream.
'
Public Sub Flush(): End Sub

''
' Returns a semi-unique number representing this instance.
'
Public Function GetHashCode() As Long: End Function

''
' Reads a block of bytes from the stream.
'
' @param Buffer The array to read the bytes into.
' @param Offset The starting index in <i>Buffer</i> to begin reading to.
' @param Count The number of bytes to read from the stream.
' @return The number of bytes actually read from the stream.
' @remarks The .NET version of this will block the thread until the requested
' number of bytes are actually read. This is not supported in this version.
'
Public Function ReadBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long: End Function

''
' Reads a single byte from the stream.
'
' @return A byte value or -1 if passed the end of the stream.
'
Public Function ReadByte() As Long: End Function

''
' Moves the current stream pointer to a new position.
'
' @param Offset The number of bytes to move the pointer (can be negative.)
' @param Origin The reference to begin moving from.
' @return The final position the pointer moved to.
'
Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency: End Function

''
' Sets the length of the current stream.
'
' @param value The new length to set the stream to.
' @remarks Generally the stream must support seeking to set the length.
'
Public Sub SetLength(ByVal value As Currency): End Sub

''
' Returns a string representation of the this instance.
'
' @return A string representation of the this instance.
'
Public Function ToString() As String: End Function

''
' Writes an array of bytes to the stream.
'
' @param Buffer The array of bytes to be written.
' @param Offset The starting index in <i>Buffer</i> to begin writing from.
' @param Count The number of bytes to be written.
'
Public Sub WriteBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long): End Sub

''
' Writes a single byte to the stream.
'
' @param value The byte to be written.
'
Public Sub WriteByte(ByVal value As Byte): End Sub


