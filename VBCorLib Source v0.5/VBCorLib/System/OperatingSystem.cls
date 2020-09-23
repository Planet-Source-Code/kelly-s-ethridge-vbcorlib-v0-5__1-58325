VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperatingSystem"
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
'    Module: OperatingSystem
'

''
' Provides information about the currently running operating system.
'
Option Explicit
Implements cObject
Implements ICloneable

Public Enum PlatformID
   Win32NT = 2
   Win32S = 0
   Win32Windows = 1
   WinCE = 3
End Enum


Private mPlatformID As PlatformID
Private mVersion As Version
Private mLoaded As Boolean


''
' Returns the operating system type currently running on this machine.
'
' @return The current operating system running.
'
Public Property Get Platform() As PlatformID
    VerifyLoaded
    Platform = mPlatformID
End Property

''
' Returns the version of the current operating system.
'
' @return The operating system version.
'
Public Property Get Version() As Version
    VerifyLoaded
    Set Version = mVersion
End Property

Public Function Clone() As Object
    VerifyLoaded
    Set Clone = Cor.NewOperatingSystem(mPlatformID, mVersion.Clone)
End Function

''
' Returns a string representation of the current operating system.
'
' @return String representing this instance.
Public Function ToString() As String
    Dim ret As String
    
    VerifyLoaded
    
    Select Case mPlatformID
        Case Win32NT
            Select Case mVersion.Major
                Case 5
                    If mVersion.Minor = 0 Then
                        ret = "Microsoft Windows 2000"
                    Else
                        ret = "Microsoft Windows XP"
                    End If
                Case Else
                    ret = "Microsoft Windows NT"
            End Select
        Case Win32S
            ret = "Microsoft Win32s"
        Case Win32Windows
            If mVersion.Major = 4 And mVersion.Minor = 0 Then
                ret = "Microsoft Windows 95"
            Else
                ret = "Microsoft Windows 98/ME"
            End If
        Case WinCE
            ret = "Microsoft WinCE"
    End Select
    
    ToString = ret & " " & mVersion.ToString
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
Friend Sub Init(ByVal Platform As PlatformID, ByVal Version As Version)
    mPlatformID = Platform
    Set mVersion = Version
    mLoaded = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyLoaded()
    Dim info As OsVersionInfo
    
    If mLoaded Then Exit Sub
    
    info.dwOSVersionInfoSize = LenB(info)
    GetVersionEx info
    Init info.dwPlatformId, Cor.NewVersion(info.dwMajorVersion, info.dwMinorVersion)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mPlatformID = .ReadProperty("PlatformID", 0)
        Set mVersion = .ReadProperty("Version", Nothing)
    End With
    mLoaded = True
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "PlateformID", mPlatformID
        .WriteProperty "Version", mVersion
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
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function
