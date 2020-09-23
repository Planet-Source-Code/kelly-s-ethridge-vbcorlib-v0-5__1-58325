VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
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
'    Module: File
'

''
' Provides methods for manipulating and retrieving file information.
'
Option Explicit

Private mDirectorySeparators() As Integer


''
' Opens a file to be written to starting at the end of the file.
'
' @param Path The name of the file to be appended to.
' @return A StreamWriter object used to write to the file.
' @remarks The StreamWriter uses a UTF8Encoding declared as Cor.NewUTF8Encoding(False, True).
'
Public Function AppendText(ByVal Path As String) As StreamWriter
    Set AppendText = Cor.NewStreamWriter(Path, Append:=True)
End Function

''
' Copies a file.
'
' @param SourceFileName The file to be copied.
' @param DestinationFileName The location and name of the copied file.
' @param OverWrite Indicates if the file already exists it should be overwritten.
' @remarks The source and destincation must be on the same volume or an exception will be thrown.
'
Public Sub Copy(ByVal SourceFileName As String, ByVal DestinationFileName As String, Optional ByVal OverWrite As Boolean = False)
    SourceFileName = cString.TrimEnd(SourceFileName, mDirectorySeparators)
    DestinationFileName = cString.TrimEnd(DestinationFileName, mDirectorySeparators)
    
    If Len(SourceFileName) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyFileName), "SourceFileName")
    If Len(DestinationFileName) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyFileName), "DestinationFileName")
    
    Path.VerifyPath SourceFileName
    Path.VerifyPath DestinationFileName
    
    If StrComp(SourceFileName, DestinationFileName, vbTextCompare) = 0 Then Exit Sub
    If CopyFile(SourceFileName, DestinationFileName, Not OverWrite) = BOOL_FALSE Then
        Dim e As Long
        e = Err.LastDllError
        If e <> ERROR_FILE_EXISTS Then
            If Not file.Exists(SourceFileName) Then _
                IOError ERROR_FILE_NOT_FOUND, SourceFileName
            If Directory.Exists(DestinationFileName) Then _
                Throw Cor.NewIOException(Environment.GetResourceString(IOException_DirectoryExists))
        End If
        IOError e, SourceFileName
    End If
End Sub

''
' Creates a file. If the file already exists, it is overwritten.
'
' @param Path The new file to be created.
' @param BufferSize The size of the internal buffer used byte the FileStream object.
' @return A FileStream object to be used for reading and writing to the file.
' @remarks The FileStream is declared as Cor.NewFileStream(Path, FileMode.Create, FileAccess.ReadWriteAccess, FileShare.None, BufferSize).
'
Public Function Create(ByVal Path As String, Optional ByVal BufferSize As Long = 4096) As FileStream
    Set Create = Cor.NewFileStream(Path, FileMode.Create, FileAccess.ReadWriteAccess, FileShare.None, BufferSize)
End Function

''
' Creates a file using a UTF8Encoding. If the file already exists, then it is overwritten.
'
' @param Path The new file to be created.
' @return A StreamWriter to be used to write to the new file.
' @remarks The StreamWriter is declared as Cor.NewStreamWriter(Path).
'
Public Function CreateText(ByVal Path As String) As StreamWriter
    Set CreateText = Cor.NewStreamWriter(Path)
End Function

''
' Deletes a file.
'
' @param sPath The file to be deleted.
' @remarks If the file did not exist, nothing happens.
'
Public Sub Delete(ByVal sPath As String)
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "Path")
    
    Path.VerifyPath sPath
    If DeleteFile(sPath) = BOOL_FALSE Then
        Dim e As Long
        e = Err.LastDllError
        If e = ERROR_FILE_NOT_FOUND Then Exit Sub
        IOError e, sPath
    End If
End Sub

''
' Determines if a file exists.
'
' @param sPath The file to check for existence.
' @return Whether the file exists or not.
'
Public Function Exists(ByVal sPath As String) As Boolean
    Dim data As WIN32_FIND_DATA
    
    If GetFileData(sPath, data) = NO_ERROR Then
        Exists = Not CBool(data.dwFileAttributes And FileAttributes.DirectoryAttr)
    End If
End Function

''
' Returns the file attributes (ReadOnly, Hidden, ...)
'
' @param sPath The file to retrieve the attributes for.
' @return A value with individual bits representing if an attribute is applied to the file.
' @remarks To determine if a specific attribute is being used, the return value will
' need to be ANDed with the specific FileAttributes flag.<br>
' <font face="Courier New">
' If Attr And FileAttributes.ReadOnly Then
' ... file is readonly
' End If
' </font>
'
Public Function GetAttributes(ByVal sPath As String) As FileAttributes
    sPath = cString.Trim(sPath)
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    GetAttributes = GetFileAttributes(sPath)
    If GetAttributes = INVALID_FILE_ATTRIBUTES Then IOError Err.LastDllError, sPath
End Function

''
' Returns the time the file was created in local time.
'
' @param sPath The file to retrieve the creation time of.
' @return A cDateTime object representing the creation time.
'
Public Function GetCreationTime(ByVal sPath As String) As cDateTime
    Dim data As WIN32_FIND_DATA
    Dim e As Long
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    e = GetFileData(sPath, data)
    If e <> NO_ERROR Then IOError e, sPath
    Set GetCreationTime = cDateTime.FromFileTime(data.ftCreationTime)
End Function

''
' Returns the time the file was created in UTC time.
'
' @param sPath The file to retrieve the creation time of.
' @return A cDateTime object representing the creation time.
'
Public Function GetCreationTimeUtc(ByVal sPath As String) As cDateTime
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    Set GetCreationTimeUtc = GetCreationTime(sPath).ToUniversalTime
End Function

''
' Returns the time the file was accessed in local time.
'
' @param sPath The file to retrieve the last access time for.
' @return A cDateTime object that represents the last time the file was accessed.
'
Public Function GetLastAccessTime(ByVal sPath As String) As cDateTime
    Dim data As WIN32_FIND_DATA
    Dim e As Long
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    e = GetFileData(sPath, data)
    If e <> NO_ERROR Then IOError e, sPath
    Set GetLastAccessTime = cDateTime.FromFileTime(data.ftLastAccessTime)
End Function

''
' Returns The last time a file was accessed in UTC time.
'
' @param sPath The file to get the last accessed time for.
' @return A cDateTime object representing the last time the file was accessed.
'
Public Function GetLastAccessTimeUtc(ByVal sPath As String) As cDateTime
    Set GetLastAccessTimeUtc = GetLastAccessTime(sPath).ToUniversalTime
End Function

''
' Returns the last time the file was written to in local time.
'
' @param sPath The file to retrieve the last written time for.
' @return A cDateTime object representing the last time the file was written to.
'
Public Function GetLastWriteTime(ByVal sPath As String) As cDateTime
    Dim data As WIN32_FIND_DATA
    Dim e As Long
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    e = GetFileData(sPath, data)
    If e <> NO_ERROR Then IOError e, sPath
    Set GetLastWriteTime = cDateTime.FromFileTime(data.ftLastWriteTime)
End Function

''
' Returns the last time the file was written to in UTC time.
'
' @param sPath The file to retrieve the last written time for.
' @return A cDateTime object representing the last time a file was written to.
'
Public Function GetLastWriteTimeUtc(ByVal sPath As String) As cDateTime
    Set GetLastWriteTimeUtc = GetLastWriteTime(sPath).ToUniversalTime
End Function

''
' Moves a file from one location to another.
'
' @param SourceFileName The file to be moved.
' @param DestinationFileName The location and filename the file is to be moved to.
' @remarks The source and destination must exist on the same volume, or an exception will be thrown.
'
Public Sub Move(ByVal SourceFileName As String, ByVal DestinationFileName As String)
    Path.VerifyPath SourceFileName
    Path.VerifyPath DestinationFileName
    
    If Len(SourceFileName) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "SourceFileName")
    If Len(DestinationFileName) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "DestinationFileName")
    
    If Not file.Exists(SourceFileName) Then IOError ERROR_FILE_NOT_FOUND, SourceFileName
    If MoveFile(SourceFileName, DestinationFileName) = BOOL_FALSE Then IOError Err.LastDllError, DestinationFileName
End Sub

''
' Opens a file with a FileStream
'
' @param Path The file to be opened.
' @param Mode The mode in which to open the file.
' @param Access The type of access allowed by the FileStream.
' @param Share The access allowed for other processes.
' @return A FileStream object of the specified file.
'
Public Function OpenFile(ByVal Path As String, ByVal Mode As FileMode, Optional ByVal Access As FileAccess = -1, Optional ByVal Share As FileShare = FileShare.None) As FileStream
    Set OpenFile = Cor.NewFileStream(Path, Mode, Access, Share)
End Function

''
' Opens a file for reading as a FileStream.
'
' @param Path The file to be opened in read mode.
' @return A FileStream used to read from the file.
' @remarks The FileStream object is declared as Cor.NewFileStream(Path, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare).
'
Public Function OpenRead(ByVal Path As String) As FileStream
    Set OpenRead = Cor.NewFileStream(Path, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare)
End Function

''
' Opens a file for reading through a StreamReader.
'
' @param Path The file to be opened for reading.
' @return A StreamReader used to read from the file with the default encoding (UTF8Encoding).
' @remarks The StreamReader is declared as Cor.NewStreamReader(Path).
'
Public Function OpenText(ByVal Path As String) As StreamReader
    Set OpenText = Cor.NewStreamReader(Path)
End Function

''
' Opens a file for writting as a FileStream.
'
' @param Path The file to be opend for writing to.
' @return A FileStream object used to write to the specified file.
' @remarks The FileStream object is declared as Cor.NewFileStream(Path, FileMode.OpenOrCreate, FileAccess.WriteAccess, FileShare.None).
'
Public Function OpenWrite(ByVal Path As String) As FileStream
    Set OpenWrite = Cor.NewFileStream(Path, FileMode.OpenOrCreate, FileAccess.WriteAccess, FileShare.None)
End Function

''
' Sets the file attributes (ReadOnly, Hidden, ...)
'
' @param sPath The file to set the attributes of.
' @param Attributes The new attributes to set the file to.
' @remarks In order to set a single attribute without unsetting an existing
' attributes, they will first need to be read from the file (GetAttributes), then
' bit manipulated to create the resulting set of desired attributes before setting.
'
Public Sub SetAttributes(ByVal sPath As String, ByVal Attributes As FileAttributes)
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    Path.VerifyPath sPath
    If SetFileAttributes(sPath, Attributes) = BOOL_FALSE Then IOError Err.LastDllError, sPath
End Sub

''
' Sets the time of creation for a file in local time.
'
' @param sPath The file to set the time of creation for.
' @param CreationTime A Date or cDateTime object of the new time.
'
Public Sub SetCreationTime(ByVal sPath As String, ByVal CreationTime As Variant)
    InternalSetFileTime sPath, cDateTime.GetcDateTime(CreationTime).ToFileTime
End Sub

''
' Sets the time of creation for a file in UTC time.
'
' @param sPath The file to set the time of creation for.
' @param CreationTime A Date or cDateTime object of the new time.
'
Public Sub SetCreationTimeUtc(ByVal sPath As String, ByVal CreationTime As Variant)
    InternalSetFileTime sPath, cDateTime.GetcDateTime(CreationTime).ToFileTimeUtc
End Sub

''
' Sets the last time the file was accessed in local time.
'
' @param sPath The file to set the last access time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
'
Public Sub SetLastAccessTime(ByVal sPath As String, ByVal LastAccessTime As Variant)
    InternalSetFileTime sPath, , cDateTime.GetcDateTime(LastAccessTime).ToFileTime
End Sub

''
' Sets the last time the file was accessed in UTC time.
'
' @param sPath The file to set the last access time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
'
Public Sub SetLastAccessTimeUtc(ByVal sPath As String, ByVal LastAccessTime As Variant)
    InternalSetFileTime sPath, , cDateTime.GetcDateTime(LastAccessTime).ToFileTimeUtc
End Sub

''
' Sets the last time the file was written to in local time.
'
' @param sPath The file to set the last written time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTime(ByVal sPath As String, ByVal LastWriteTime As Variant)
    InternalSetFileTime sPath, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTime
End Sub

''
' Sets the last time the file was written to in UTC time.
'
' @param sPath The file to set the last written time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTimeUtc(ByVal sPath As String, ByVal LastWriteTime As Variant)
    InternalSetFileTime sPath, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTimeUtc
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetFileData(ByVal sPath As String, ByRef data As WIN32_FIND_DATA) As Long
    Dim h As Handle
    
    sPath = cString.TrimEnd(sPath, mDirectorySeparators)
    If Len(sPath) = 0 Then Exit Function
    
    Path.VerifyPath sPath
    h = FindFirstFile(sPath, data)
    If h <> INVALID_HANDLE Then
        FindClose h
    Else
        GetFileData = Err.LastDllError
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InternalSetFileTime(ByRef sPath As String, Optional ByVal CreationTime As Variant, Optional ByVal LastAccessTime As Variant, Optional ByVal LastWriteTime As Variant)
    Dim ct As Long
    Dim lat As Long
    Dim lwt As Long
    Dim h As Long
    
    If Not IsMissing(CreationTime) Then ct = VarPtr(CreationTime) + 8
    If Not IsMissing(LastAccessTime) Then lat = VarPtr(LastAccessTime) + 8
    If Not IsMissing(LastWriteTime) Then lwt = VarPtr(LastWriteTime) + 8
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    Path.VerifyPath sPath
    
    h = GetFileHandle(sPath)
    If h = INVALID_HANDLE Then IOError Err.LastDllError, sPath
    If SetFileTime(h, ByVal ct, ByVal lat, ByVal lwt) = BOOL_FALSE Then
        CloseHandle h
        IOError Err.LastDllError, sPath
    End If
    CloseHandle h
End Sub

Private Function GetFileHandle(ByRef sPath As String) As Handle
    GetFileHandle = CreateFile(sPath, FileAccess.WriteAccess, FileShare.None, ByVal 0&, FileMode.OpenExisting, FILE_ATTRIBUTE_NORMAL, 0)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mDirectorySeparators = cArray.NewArray(ciInteger, Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
End Sub
