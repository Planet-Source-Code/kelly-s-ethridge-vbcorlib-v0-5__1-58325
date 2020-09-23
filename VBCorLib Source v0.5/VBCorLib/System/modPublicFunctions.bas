Attribute VB_Name = "modPublicFunctions"
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
'    Module: modPublicFunctions
'
'   This mirrors PublicFunctions.cls to allow the project access to the
'   same set of public functions.
Option Explicit

Public Cor As Constructors
Public cArray As cArray
Public cString As cString
Public comparer As ComparerStatic
Public Environment As Environment
Public BitArray As BitArrayStatic
Public Buffer As Buffer
Public NumberFormatInfo As NumberFormatInfoStatic
Public BitConverter As BitConverter
Public Version As VersionStatic
Public TimeSpan As TimeSpanStatic
Public cDateTime As cDateTimeStatic
Public DateTimeFormatInfo As DateTimeFormatInfoStatic
Public CultureTable As CultureTable
Public CultureInfo As CultureInfoStatic
Public TimeZone As TimeZoneStatic
Public Path As Path
Public Encoding As EncodingStatic
Public TextReader As TextReaderStatic
Public Directory As Directory
Public File As File
Public Stream As StreamStatic


Public Powers(31) As Long



Public Sub InitPublicFunctions()
    InitPowers
    
    Set comparer = New ComparerStatic
    Set Cor = New Constructors
    Set cArray = New cArray
    Set cString = New cString
    Set Environment = New Environment
    Set BitArray = New BitArrayStatic
    Set Buffer = New Buffer
    Set CultureTable = New CultureTable
    Set CultureInfo = New CultureInfoStatic
    Set NumberFormatInfo = New NumberFormatInfoStatic
    Set BitConverter = New BitConverter
    Set Version = New VersionStatic
    Set TimeSpan = New TimeSpanStatic
    Set cDateTime = New cDateTimeStatic
    Set DateTimeFormatInfo = New DateTimeFormatInfoStatic
    Set TimeZone = New TimeZoneStatic
    Set Path = New Path
    Set Encoding = New EncodingStatic
    Set TextReader = New TextReaderStatic
    Set Directory = New Directory
    Set File = New File
    Set Stream = New StreamStatic
    
End Sub

Public Function FuncAddr(ByVal pfn As Long) As Long
    FuncAddr = pfn
End Function

Public Function CObj(ByRef value As Variant) As Object
    Set CObj = value
End Function

Public Function Modulus(ByVal x As Currency, ByVal y As Currency) As Currency
  Modulus = x - (y * Fix(x / y))
End Function


Private Sub InitPowers()
    Dim i As Long
    For i = 0 To 30
        Powers(i) = 2 ^ i
    Next i
    Powers(31) = &H80000000
End Sub

