VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfo"
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
'    Module: CultureInfo
'

''
' Represents information about a culture, such as number and date formatting.
'
Option Explicit
Implements cObject
Implements ICloneable
Implements IFormatProvider

Private Const PROP_LCID                             As String = "LCID"
Private Const PROP_PARENTLCID                       As String = "ParentLCID"
Private Const PROP_NAME                             As String = "Name"
Private Const PROP_ENGLISHNAME                      As String = "EnglishName"
Private Const PROP_DISPLAYNAME                      As String = "DisplayName"
Private Const PROP_NATIVENAME                       As String = "NativeName"
Private Const PROP_THREELETTERISOLANGUAGENAME       As String = "ThreeLetterISOLanguageName"
Private Const PROP_THREELETTERWINDOWSLANGUAGENAME   As String = "ThreeLetterWindowsLanguageName"
Private Const PROP_TWOLETTERISOLANGUAGENAME         As String = "TwoLetterISOLanguageName"
Private Const PROP_USEUSEROVERRIDE                  As String = "UseUserOverride"
Private Const PROP_ISREADONLY                       As String = "IsReadOnly"
Private Const PROP_DATETIMEFORMAT                   As String = "DateTimeFormat"
Private Const PROP_NUMBERFORMAT                     As String = "NumberFormat"


Private Type PropsType
    LCID As Long
    ParentLCID As Long
    Name As String
    EnglishName As String
    DisplayName As String
    NativeName As String
    ThreeLetterISOLanguageName As String
    ThreeLetterWindowLanguageName As String
    TwoLetterISOLanguageName As String
    UseUserOverride As Boolean
    
    IsReadOnly As Boolean
End Type

Private mProps As PropsType
Private mNumberFormat As NumberFormatInfo
Private mDateTimeFormat As DateTimeFormatInfo
Private mLoaded As Boolean


''
' Returns the Locale ID for this culture instance.
'
' @return The Locale ID.
' @remarks A 32 bit LCID is in the format of:
'        31 - 20 19 18 17 16 15 14 13 12 11 10 9 8 7 6 5 4 3 2 1 0
'        +-----+ +---------+ +---------------+ +-----------------+
'           |         |           |            Primary language ID (10 bits)
'           |         |           +----------- Sublanguage ID (6 its)
'           |         +----------------------- Sort ID (4 bits)
'           +--------------------------------- Reserved (12 bits)
'
Public Property Get LCID() As Long
    VerifyLoaded
    LCID = mProps.LCID
End Property

''
' Returns the name of the culture.
'
' @return The name of the culture.
' @remarks The culture name follows the RFC 1766 standard in that there are
' 2 parts. A 2 letter culture and 2 letter region:
' en-US = English (United States)
'
Public Property Get Name() As String
    VerifyLoaded
    Name = mProps.Name
End Property

''
' Returns the english translation for the culture name.
'
' @return The name in the english language.
'
Public Property Get EnglishName() As String
    VerifyLoaded
    EnglishName = mProps.EnglishName
End Property

''
' Returns a displayable name for the culture name.
'
' @return The name as it should be displayed.
'
Public Property Get DisplayName() As String
    VerifyLoaded
    DisplayName = mProps.DisplayName
End Property

''
' Returns if this culture is neutral.
'
' @return Indication of the culture's neutrality.
' @remarks A neutral culture is a culture that has not region specified.
' en is a neutral culture, where as en-US is not.
'
Public Property Get IsNeutralCulture() As Boolean
    VerifyLoaded
    If mProps.LCID = INVARIANT_LCID Then Exit Property
    IsNeutralCulture = CultureTable.IsNeutral(mProps.LCID)
End Property

''
' Returns if this instance is ReadOnly.
'
' @return The readonly status of this instance.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mProps.IsReadOnly
End Property

Friend Property Let IsReadOnly(ByVal RHS As Boolean)
    mProps.IsReadOnly = RHS
    If Not mDateTimeFormat Is Nothing Then mDateTimeFormat.IsReadOnly = RHS
    If Not mNumberFormat Is Nothing Then mNumberFormat.IsReadOnly = RHS
End Property

''
' Returns the native name for this culture instance.
'
' @return The name using native characters to spell the culture name.
'
Public Property Get NativeName() As String
    VerifyLoaded
    NativeName = mProps.NativeName
End Property

''
' Returns the 3 letter ISO 639-2 standard of the culture name.
'
' @return The name of the culture in three letters.
'
Public Property Get ThreeLetterISOLanguageName() As String
    VerifyLoaded
    ThreeLetterISOLanguageName = mProps.ThreeLetterISOLanguageName
End Property

''
' Returns the 3 letter Windows name of the culture name.
'
' @return The name of the culture found in windows.
'
Public Property Get ThreeLetterWindowsLanguageName() As String
    VerifyLoaded
    ThreeLetterWindowsLanguageName = mProps.ThreeLetterWindowLanguageName
End Property

''
' Returns the 2 letter  ISO 639-1 standard of the culture name.
'
' @return The name of the culture in 2 letters.
'
Public Property Get TwoLetterISOLanguageName() As String
    VerifyLoaded
    TwoLetterISOLanguageName = mProps.TwoLetterISOLanguageName
End Property

''
' Returns whether the culture instance is using the user settings.
'
' @return Indication as to if the culture is using the user setting
' defined in the control panel.
'
Public Property Get UseUserOverride() As Boolean
    UseUserOverride = mProps.UseUserOverride
End Property

''
' Returns a clone of this culture instance.
'
' @return An instance of this culture. ReadOnly status is retained.
'
Public Function Clone() As Object
    Dim ret As CultureInfo
    Dim dtInfo As DateTimeFormatInfo
    Dim nInfo As NumberFormatInfo
    
    Set ret = New CultureInfo
    If Not mDateTimeFormat Is Nothing Then Set dtInfo = mDateTimeFormat.Clone
    If Not mNumberFormat Is Nothing Then Set nInfo = mNumberFormat.Clone
    
    ret.CloneHelper mProps, dtInfo, nInfo
    Set Clone = ret
End Function

''
' Returns a format provider of the specified type.
'
' @param FormatType The type of format provider that is requested.
' @return The format provider, or Nothing if a provider was not found.
'
Public Function GetFormat(ByVal FormatType As String) As Object
    Select Case LCase$(FormatType)
        Case "numberformatinfo"
            Set GetFormat = NumberFormat
        Case "datetimeformatinfo"
            Set GetFormat = DateTimeFormat
    End Select
End Function

''
' Returns the NumberFormatInfo associated with this culture.
'
' @return The NumberFormatInfo for this culture that can be used
' to format numbers specific to this culture.
'
Public Property Get NumberFormat() As NumberFormatInfo
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
    
    If mNumberFormat Is Nothing Then
        Set mNumberFormat = New NumberFormatInfo
        mNumberFormat.Load mProps.LCID
        mNumberFormat.IsReadOnly = mProps.IsReadOnly
    End If
    Set NumberFormat = mNumberFormat
End Property

''
' Sets the NumberFormatInfo for the specific culture object.
'
' @param RHS The NumberFormatInfo to associate with this object instance.
'
Public Property Set NumberFormat(ByVal RHS As NumberFormatInfo)
    VerifyWritable
    If RHS Is Nothing Then _
        Throw Cor.NewArgumentNullException("Cannot set NumberFormat to Nothing.", "NumberFormat")
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
        
    Set mNumberFormat = RHS
End Property

''
' Returns the DateTimeFormatInfo associated with this culture.
'
' @return The DateTimeFormatInfo for this culture that can be used
' to format dates and times specific to this culture.
'
Public Property Get DateTimeFormat() As DateTimeFormatInfo
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
        
    If mDateTimeFormat Is Nothing Then
        Set mDateTimeFormat = New DateTimeFormatInfo
        mDateTimeFormat.Load mProps.LCID
        mDateTimeFormat.IsReadOnly = mProps.IsReadOnly
    End If
    Set DateTimeFormat = mDateTimeFormat
End Property

''
' Sets the DateTimeFormatInfo for the specific culture object.
'
' @param RHS The DateTimeFormatInfo object to associate with this object instance.
'
Public Property Set DateTimeFormat(ByVal RHS As DateTimeFormatInfo)
    VerifyWritable
    If RHS Is Nothing Then _
        Throw Cor.NewArgumentNullException("Cannot set DateTimeFormat to Nothing.", "DateTimeFormat")
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
    
    Set mDateTimeFormat = RHS
End Property

''
' Returns the parent culture for this culture.
'
' @return The parent culture for this culture. If this culture is
' invariant, then invariant is returned.
'
Public Property Get Parent() As CultureInfo
    Set Parent = Cor.NewCultureInfo(mProps.ParentLCID)
End Property

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
        Dim c As CultureInfo
        If TypeOf value Is CultureInfo Then
            Set c = value
            Equals = (c.LCID = mProps.LCID)
        End If
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
Friend Sub Init(ByRef NameOrLCID As Variant, ByVal UseUserOverride As Boolean)
    Dim ID As Long
    
    Select Case VarType(NameOrLCID)
        Case vbLong, vbInteger, vbByte
            ID = NameOrLCID
        Case vbString
            If CultureTable.IsLoaded Then ID = CultureTable.GetCultureID(NameOrLCID)
        Case Else
            Throw Cor.NewArgumentException("Invalid Culture Identifier.", "NameOrLCID")
    End Select
    Load ID
    mLoaded = True
End Sub

Friend Sub CloneHelper(ByRef props As PropsType, ByVal dtInfo As DateTimeFormatInfo, ByVal nInfo As NumberFormatInfo)
    mProps = props
    Set mDateTimeFormat = dtInfo
    Set mNumberFormat = nInfo
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyWritable()
    If mProps.IsReadOnly Then Throw Cor.NewInvalidOperationException("Culture is Read-Only.")
End Sub

Private Sub VerifyLoaded()
    If mLoaded Then Exit Sub
    If CultureTable.IsLoaded Then
        Load INVARIANT_LCID
    Else
        LoadDefault
    End If
    mLoaded = True
End Sub

Private Sub Load(ByVal LCID As Long)
    If Not CultureTable.IsLoaded Then
        LoadDefault
        Exit Sub
    End If
    
    With mProps
        .LCID = LCID
        .ParentLCID = CultureTable.GetNumber(LCID, IPARENTLCID)
        .Name = CultureTable.GetString(LCID, SNAME)
        .EnglishName = CultureTable.GetString(LCID, SENGLISHNAME)
        .DisplayName = CultureTable.GetString(LCID, SDISPLAYNAME)
        .NativeName = CultureTable.GetString(LCID, SNATIVENAME)
        .ThreeLetterISOLanguageName = CultureTable.GetString(LCID, STHREELETTERISOLANGUAGENAME)
        .ThreeLetterWindowLanguageName = CultureTable.GetString(LCID, STHREELETTERWINDOWSLANGUAGENAME)
        .TwoLetterISOLanguageName = CultureTable.GetString(LCID, STWOLETTERISOLANGUAGENAME)
    End With
End Sub

Private Sub LoadDefault()
    With mProps
        .LCID = INVARIANT_LCID
        .ParentLCID = INVARIANT_LCID
        .Name = ""
        .EnglishName = "Invariant Language (Invariant Country)"
        .DisplayName = "Invariant Language (Invariant Country)"
        .NativeName = "Invariant Language (Invariant Country)"
        .ThreeLetterISOLanguageName = "IVL"
        .ThreeLetterWindowLanguageName = "IVL"
        .TwoLetterISOLanguageName = "iv"
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mProps.LCID = .ReadProperty(PROP_LCID)
        mProps.EnglishName = .ReadProperty(PROP_ENGLISHNAME)
        mProps.DisplayName = .ReadProperty(PROP_DISPLAYNAME)
        mProps.NativeName = .ReadProperty(PROP_NATIVENAME)
        mProps.Name = .ReadProperty(PROP_NAME)
        mProps.ParentLCID = .ReadProperty(PROP_PARENTLCID)
        mProps.ThreeLetterISOLanguageName = .ReadProperty(PROP_THREELETTERISOLANGUAGENAME)
        mProps.ThreeLetterWindowLanguageName = .ReadProperty(PROP_THREELETTERWINDOWSLANGUAGENAME)
        mProps.TwoLetterISOLanguageName = .ReadProperty(PROP_TWOLETTERISOLANGUAGENAME)
        mProps.UseUserOverride = .ReadProperty(PROP_USEUSEROVERRIDE)
        mProps.IsReadOnly = .ReadProperty(PROP_ISREADONLY)
        Set mDateTimeFormat = .ReadProperty(PROP_DATETIMEFORMAT)
        Set mNumberFormat = .ReadProperty(PROP_NUMBERFORMAT)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PROP_LCID, mProps.LCID
        .WriteProperty PROP_PARENTLCID, mProps.ParentLCID
        .WriteProperty PROP_NAME, mProps.Name
        .WriteProperty PROP_ENGLISHNAME, mProps.EnglishName
        .WriteProperty PROP_DISPLAYNAME, mProps.DisplayName
        .WriteProperty PROP_NATIVENAME, mProps.NativeName
        .WriteProperty PROP_THREELETTERISOLANGUAGENAME, mProps.ThreeLetterISOLanguageName
        .WriteProperty PROP_THREELETTERWINDOWSLANGUAGENAME, mProps.ThreeLetterWindowLanguageName
        .WriteProperty PROP_TWOLETTERISOLANGUAGENAME, mProps.TwoLetterISOLanguageName
        .WriteProperty PROP_USEUSEROVERRIDE, mProps.UseUserOverride
        .WriteProperty PROP_ISREADONLY, mProps.IsReadOnly
        .WriteProperty PROP_DATETIMEFORMAT, mDateTimeFormat
        .WriteProperty PROP_NUMBERFORMAT, mNumberFormat
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(value As Variant) As Boolean
    cObject_Equals = Equals(value)
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormatProvider Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormatProvider_GetFormat(ByVal FormatType As String) As Object
    Set IFormatProvider_GetFormat = GetFormat(FormatType)
End Function
