VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicFunctions"
Attribute VB_GlobalNameSpace = True
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
'    Module: PublicFunctions
'

''
'
Option Explicit

' Returns accessors to categorized functions.
Public Function Cor() As Constructors
    Set Cor = modPublicFunctions.Cor
End Function

Public Function cArray() As cArray
    Set cArray = modPublicFunctions.cArray
End Function

Public Function cString() As cString
    Set cString = modPublicFunctions.cString
End Function

Public Function comparer() As ComparerStatic
    Set comparer = modPublicFunctions.comparer
End Function

Public Function Environment() As Environment
    Set Environment = modPublicFunctions.Environment
End Function

Public Function BitArray() As BitArrayStatic
    Set BitArray = modPublicFunctions.BitArray
End Function

Public Function Buffer() As Buffer
    Set Buffer = modPublicFunctions.Buffer
End Function

Public Function NumberFormatInfo() As NumberFormatInfoStatic
    Set NumberFormatInfo = modPublicFunctions.NumberFormatInfo
End Function

Public Function BitConverter() As BitConverter
    Set BitConverter = modPublicFunctions.BitConverter
End Function

Public Function Version() As VersionStatic
    Set Version = modPublicFunctions.Version
End Function

Public Function TimeSpan() As TimeSpanStatic
    Set TimeSpan = modPublicFunctions.TimeSpan
End Function

Public Function cDateTime() As cDateTimeStatic
    Set cDateTime = modPublicFunctions.cDateTime
End Function

Public Function DateTimeFormatInfo() As DateTimeFormatInfoStatic
    Set DateTimeFormatInfo = modPublicFunctions.DateTimeFormatInfo
End Function

Public Function CultureInfo() As CultureInfoStatic
    Set CultureInfo = modPublicFunctions.CultureInfo
End Function

Public Function TimeZone() As TimeZoneStatic
    Set TimeZone = modPublicFunctions.TimeZone
End Function

Public Function Path() As Path
    Set Path = modPublicFunctions.Path
End Function

Public Function Encoding() As EncodingStatic
    Set Encoding = modPublicFunctions.Encoding
End Function

Public Function TextReader() As TextReaderStatic
    Set TextReader = modPublicFunctions.TextReader
End Function

Public Function Directory() As Directory
    Set Directory = modPublicFunctions.Directory
End Function

Public Function file() As file
    Set file = modPublicFunctions.file
End Function

Public Function Stream() As StreamStatic
    Set Stream = modPublicFunctions.Stream
End Function

Public Function Catch(ByRef ex As Exception, Optional ByVal Err As ErrObject) As Boolean
    Catch = modExceptionMethods.Catch(ex, Err)
End Function

Public Sub Throw(ByVal ex As Exception)
    modExceptionMethods.Throw ex
End Sub


