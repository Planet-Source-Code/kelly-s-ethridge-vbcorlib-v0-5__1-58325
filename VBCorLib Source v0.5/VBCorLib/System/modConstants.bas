Attribute VB_Name = "modConstants"
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
'    Module: modConstants
'
Option Explicit
Public Enum BucketStateEnum
    bsEmpty
    bsOccupied
    bsDeleted
End Enum

Public Type STRINGREF
    Length As Long
    SA As SafeArray1d
    Chars() As Integer
End Type
Public Type Bucket
    Key As Variant
    value As Variant
    hashcode As Long
    State As BucketStateEnum
End Type


Public Const LOWER_A_CHAR       As Integer = 97
Public Const LOWER_Z_CHAR       As Integer = 122
Public Const UPPER_A_CHAR       As Integer = 65
Public Const UPPER_Z_CHAR       As Integer = 90
Public Const CHAR_0             As Integer = 48
Public Const CHAR_9             As Integer = 57
Public Const CHAR_PLUS_SIGN     As Integer = 43
Public Const CHAR_MINUS_SIGN    As Integer = 45
Public Const CHAR_UPPER_A       As Long = 65
Public Const CHAR_UPPER_Z       As Long = 90
Public Const CHAR_LOWER_A       As Long = 97
Public Const CHAR_LOWER_Z       As Long = 122
Public Const CHAR_BACKSLASH     As Long = 92
Public Const CHAR_FORSLASH      As Long = 47
Public Const CHAR_COLON         As Long = 58

Public Const INTEGER_ARRAY As Long = vbArray Or vbInteger

Public Const MAX_PATH               As Long = 260
Public Const MAX_DIRECTORY_PATH     As Long = 260
Public Const NO_ERROR               As Long = 0


Public Const FILE_FLAG_OVERLAPPED As Long = &H40000000
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const FILE_TYPE_DISK As Long = &H1
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const INVALID_FILE_ATTRIBUTES As Long = -1

' File manipulation function attributes
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const OPEN_EXISTING As Long = 3
Public Const PAGE_READONLY As Long = &H2
Public Const SECTION_MAP_READ As Long = &H4
Public Const FILE_MAP_READ As Long = SECTION_MAP_READ
Public Const INVALID_HANDLE As Long = -1

Public Const ERROR_PATH_NOT_FOUND       As Long = 3
Public Const ERROR_ACCESS_DENIED        As Long = 5
Public Const ERROR_FILE_NOT_FOUND       As Long = 2
Public Const ERROR_FILE_EXISTS          As Long = 80

    

' Exception HResults
Public Const E_POINTER                  As Long = &H5B
Public Const COR_E_EXCEPTION            As Long = &H80131500
Public Const COR_E_SYSTEM               As Long = &H80131501
Public Const COR_E_RANK                 As Long = &H9
Public Const COR_E_INVALIDOPERATION     As Long = &H5
Public Const COR_E_INVALIDCAST          As Long = &HD
Public Const COR_E_INDEXOUTOFRANGE      As Long = &H9
Public Const COR_E_ARGUMENT             As Long = &H5
Public Const COR_E_ARGUMENTOUTOFRANGE   As Long = &H5
Public Const COR_E_OUTOFMEMORY          As Long = &H7
Public Const COR_E_FORMAT               As Long = &H80131537
Public Const COR_E_NOTSUPPORTED         As Long = &H1B6
Public Const COR_E_SERIALIZATION        As Long = &H14A
Public Const COR_E_ARRAYTYPEMISMATCH    As Long = &HD
Public Const COR_E_IO                   As Long = &H39
Public Const COR_E_FILENOTFOUND         As Long = &H35
Public Const COR_E_PLATFORMNOTSUPPORTED As Long = &H80131539
Public Const COR_E_PATHTOOLONG          As Long = &H800700CE
Public Const COR_E_DIRECTORYNOTFOUND    As Long = &H35
Public Const COR_E_ENDOFSTREAM          As Long = &H80070026
Public Const COR_E_ARITHMETIC           As Long = &H80070216
Public Const COR_E_OVERFLOW             As Long = &H6



' Resource Strings
' ArrayTypeMismatch
Public Const ArrayTypeMismatch_Conversion               As Long = 101
Public Const ArrayTypeMismatch_Incompatible             As Long = 102
Public Const ArrayTypeMismatch_Exception                As Long = 103
Public Const ArrayTypeMismatch_Compare                  As Long = 104

' Rank
Public Const Rank_MultiDimension                        As Long = 200

' IndexOutOfRange
Public Const IndexOutOfRange_Dimension                  As Long = 300

' IOException
Public Const IOException_Exception                      As Long = 400
Public Const IOException_DirectoryExists                As Long = 401

' FileNotFound
Public Const FileNotFound_Exception                     As Long = 500

' ArgumentOutOfRange
Public Const ArgumentOutOfRange_MustBeNonNegNum         As Long = 1000
Public Const ArgumentOutOfRange_SmallCapacity           As Long = 1001
Public Const ArgumentOutOfRange_NeedNonNegNum           As Long = 1002
Public Const ArgumentOutOfRange_ArrayListInsert         As Long = 1003
Public Const ArgumentOutOfRange_Index                   As Long = 1004
Public Const ArgumentOutOfRange_LargerThanCollection    As Long = 1005
Public Const ArgumentOutOfRange_LBound                  As Long = 1006
Public Const ArgumentOutOfRange_Exception               As Long = 1007
Public Const ArgumentOutOfRange_Range                   As Long = 1008
Public Const ArgumentOutOfRange_UBound                  As Long = 1009
Public Const ArgumentOutOfRange_MinMax                  As Long = 1010
Public Const ArgumentOutOfRange_VersionFieldCount       As Long = 1011
Public Const ArgumentOutOfRange_ValidValues             As Long = 1012
Public Const ArgumentOutOfRange_NeedPosNum              As Long = 1013

' Argument
Public Const Argument_InvalidCountOffset                As Long = 2000
Public Const Argument_ArrayPlusOffTooSmall              As Long = 2001
Public Const Argument_Exception                         As Long = 2002
Public Const Argument_ArrayRequired                     As Long = 2003
Public Const Argument_MatchingBounds                    As Long = 2004
Public Const Argument_IndexPlusTypeSize                 As Long = 2005
Public Const Argument_VersionRequired                   As Long = 2006
Public Const Argument_TimeSpanRequired                  As Long = 2007
Public Const Argument_DateRequired                      As Long = 2008
Public Const Argument_InvalidHandle                     As Long = 2009
Public Const Argument_EmptyPath                         As Long = 2010
Public Const Argument_SmallConversionBuffer             As Long = 2011
Public Const Argument_EmptyFileName                     As Long = 2012
Public Const Argument_ReadableStreamRequired            As Long = 2013

' ArgumentNull
Public Const ArgumentNull_Array                         As Long = 2100
Public Const ArgumentNull_Exception                     As Long = 2101
Public Const ArgumentNull_Stream                        As Long = 2102

' NotSupported
Public Const NotSupported_ReadOnlyCollection            As Long = 3000
Public Const NotSupported_FixedSizeCollection           As Long = 3001

' InvalidOperation
Public Const InvalidOperation_EmptyStack                As Long = 4000
Public Const InvalidOperation_EnumNotStarted            As Long = 4001
Public Const InvalidOperation_EnumFinished              As Long = 4002
Public Const InvalidOperation_VersionError              As Long = 4003
Public Const InvalidOperation_EmptyQueue                As Long = 4004
Public Const InvalidOperation_Comparer_Arg              As Long = 4005
Public Const InvalidOperation_ReadOnly                  As Long = 4006

' Constants used by CultureInfo and related classes when
' utilizing the CultureTable class.
Public Const LCID_INSTALLED As Long = &H1
Public Const LCID_SUPPORTED As Long = &H2
Public Const INVARIANT_LCID As Long = 127
             
Public Const ILCID                          As Long = 0
Public Const IPARENTLCID                    As Long = 1
Public Const IFIRSTWEEKOFYEAR               As Long = 2
Public Const IFIRSTDAYOFWEEK                As Long = 3
Public Const ICURRENCYDECIMALDIGITS         As Long = 4
Public Const ICURRENCYNEGATIVEPATTERN       As Long = 5
Public Const ICURRENCYPOSITIVEPATTERN       As Long = 6
Public Const INUMBERDECIMALDIGITS           As Long = 7
Public Const INUMBERNEGATIVEPATTERN         As Long = 8
Public Const IPERCENTDECIMALDIGITS          As Long = 9
Public Const IPERCENTNEGATIVEPATTERN        As Long = 10
Public Const IPERCENTPOSITIVEPATTERN        As Long = 11
'Public Const ICALENDARTYPE As Long = 14


Public Const SENGLISHNAME                   As Long = 0
Public Const SDISPLAYNAME                   As Long = 1
Public Const SNAME                          As Long = 2
Public Const SNATIVENAME                    As Long = 3
Public Const STHREELETTERISOLANGUAGENAME    As Long = 4
Public Const STWOLETTERISOLANGUAGENAME      As Long = 5
Public Const STHREELETTERWINDOWSLANGUAGENAME As Long = 6
Public Const SABBREVIATEDDAYNAMES           As Long = 7
Public Const SABBREVIATEDMONTHNAMES         As Long = 8
Public Const SAMDESIGNATOR                  As Long = 9
Public Const SDATESEPARATOR                 As Long = 10
Public Const SDAYNAMES                      As Long = 11
Public Const SLONGDATEPATTERN               As Long = 12
Public Const SLONGTIMEPATTERN               As Long = 13
Public Const SMONTHDAYPATTERN               As Long = 14
Public Const SMONTHNAMES                    As Long = 15
Public Const SPMDESIGNATOR                  As Long = 16
Public Const SSHORTDATEPATTERN              As Long = 17
Public Const SSHORTTIMEPATTERN              As Long = 18
Public Const STIMESEPARATOR                 As Long = 19
Public Const SYEARMONTHPATTERN              As Long = 20
Public Const SALLLONGDATEPATTERNS           As Long = 21
Public Const SALLSHORTDATEPATTERNS          As Long = 22
Public Const SALLLONGTIMEPATTERNS           As Long = 23
Public Const SALLSHORTTIMEPATTERNS          As Long = 24
Public Const SALLMONTHDAYPATTERNS           As Long = 25
Public Const SCURRENCYGROUPSIZES            As Long = 26
Public Const SNUMBERGROUPSIZES              As Long = 27
Public Const SPERCENTGROUPSIZES             As Long = 28
Public Const SCURRENCYDECIMALSEPARATOR      As Long = 29
Public Const SCURRENCYGROUPSEPARATOR        As Long = 30
Public Const SCURRENCYSYMBOL                As Long = 31
Public Const SNANSYMBOL                     As Long = 32
Public Const SNEGATIVEINFINITYSYMBOL        As Long = 33
Public Const SNEGATIVESIGN                  As Long = 34
Public Const SNUMBERDECIMALSEPARATOR        As Long = 35
Public Const SNUMBERGROUPSEPARATOR          As Long = 36
Public Const SPERCENTDECIMALSEPARATOR       As Long = 37
Public Const SPERCENTGROUPSEPARATOR         As Long = 38
Public Const SPERCENTSYMBOL                 As Long = 39
Public Const SPERMILLESYMBOL                As Long = 40
Public Const SPOSITIVEINFINITYSYMBOL        As Long = 41
Public Const SPOSITIVESIGN                  As Long = 42




