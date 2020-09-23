VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfoStatic"
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
'    Module: CultureInfoStatic
'

''
' Provides a set of static methods retrieve CultureInfo objects.
'
Option Explicit

Public Enum CultureTypes
    NeutralCultures = 1
    SpecificCultures = 2
    InstalledWin32Cultures = 4
    AllCultures = 7
End Enum

Private mCurrentCulture As CultureInfo
Private mInvariantCulture As CultureInfo



''
' Returns a ReadOnly wrapped CultureInfo object.
'
' @param Culture The culture to create a ReadOnly wrapper for.
' @return The wrapped ReadOnly culture.
'
Public Function ReadOnly(ByVal Culture As CultureInfo) As CultureInfo
    If Culture.IsReadOnly Then
        Set ReadOnly = Culture
    Else
        Set ReadOnly = Culture.Clone
        ReadOnly.IsReadOnly = True
    End If
End Function

''
' Returns a set of CultureInfo objects.
'
' @param Types The culture types that are to be retrieved.
' @return An array of CultureInfo objects based on Types specified.
'
Public Function GetCultures(ByVal Types As CultureTypes) As CultureInfo()
    GetCultures = CultureTable.GetCultures(Types)
End Function

''
' Returns the current culture associated with this machine.
'
' @return A ReadOnly version of CultureInfo specific for this machine.
'
Public Function CurrentCulture() As CultureInfo
    Set CurrentCulture = mCurrentCulture
End Function

''
' Returns a ReadOnly version of the invariant culture.
'
' @return A ReadOnly version of the invariant culture.
'
Public Function InvariantCulture() As CultureInfo
    Set InvariantCulture = mInvariantCulture
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mCurrentCulture = Cor.NewCultureInfo(GetSystemDefaultLCID)
    mCurrentCulture.IsReadOnly = True
    
    Set mInvariantCulture = Cor.NewCultureInfo(INVARIANT_LCID)
    mInvariantCulture.IsReadOnly = True
End Sub
