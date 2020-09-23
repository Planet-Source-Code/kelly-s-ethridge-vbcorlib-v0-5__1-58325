VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cObject"
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
'    Module: cObject
'

''
' This is the base interface to allow objects to be utilized throughout most of the system.
'
' <H4>Remarks</H4> If this interface is implements, then the same methods should be
' implemented in the normal public interface to maintain consitency.
'
Option Explicit

''
' Returns a boolean indicating the equality of this instance versus the
' passed in value.
'
' @param value The value to compare this instance against for equality.
' @return Boolean indicating the equality of this instance and the value.
' @remarks This should be overrided to provide a the type of equality
' checking that is appropriate for each class.  Most classes simply
' check to see if the value is the same object instance as this object,
' meaning, they are both the same object in memory. cDateTime checks to
' see if the value has the same date as this instance, and TimeSpan
' checks to see if the value has the same number of milliseconds as this
' instance to determine equality.
'
Public Function Equals(ByRef value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number to identify this object instance.
'
' @return The number representing this instance.
' @remarks This should be override to provide a value to identify
' this object instance. Most objects return their location in memory,
' unless the objects need to be considered the same as other objects.
' This occurs for both cDateTime and TimeSpan objects. All cDateTime
' objects that have the same date will generate the same hashcode. Same
' for TimeSpan. If the spans are the same, so are the hashcodes.
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string representation of the object.
'
' @return A string representation of the object.
' @remarks This should be overriden to return appropriate string values
' for the given object. StringBuilder returns the current state of the
' internal string that is being built.
'
Public Function ToString() As String: End Function

