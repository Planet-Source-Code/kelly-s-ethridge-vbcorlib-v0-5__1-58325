Attribute VB_Name = "modWeakReferenceHelpers"
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
'    Module: modWeakReferenceHelpers
'
Option Explicit

Private Type VTableType
    VTable(2) As Long
End Type
Public Type WeakRefHookType
    VTable As VTableType
    owner As WeakReference
    Target As IVBUnknown
    pOriginalVTable As Long
End Type

Private Type WeakSafeArray
    SA As SafeArray1d
    WeakRef() As WeakRefHookType
End Type

Private mWeak As WeakSafeArray
Private mRelease As Long

Public Sub InitWeakReference(ByRef ref As WeakRefHookType, ByVal owner As WeakReference, ByVal Target As IUnknown)
    If mRelease = 0 Then
        mRelease = FuncAddr(AddressOf Release)
        With mWeak
            With .SA
                .cbElements = LenB(ref)
                .cDims = 1
                .cElements = 1
            End With
            SAPtr(.WeakRef) = VarPtr(.SA)
        End With
    End If
    
    Dim p As Long
    With ref
        p = VTablePtr(Target)
        With .VTable
            .VTable(0) = MemLong(p)
            .VTable(1) = MemLong(p + 4)
            .VTable(2) = mRelease
        End With
        .pOriginalVTable = p
        ObjectPtr(.owner) = ObjectPtr(owner)
        ObjectPtr(.Target) = ObjectPtr(Target)
        p = MemLong(VarPtr(Target))
        Set Target = Nothing
        MemLong(p) = VarPtr(ref)
    End With
End Sub

Public Sub DisposeWeakReference(ByRef ref As WeakRefHookType)
    With ref
        If .pOriginalVTable <> 0 Then
            VTablePtr(.Target) = .pOriginalVTable
            .pOriginalVTable = 0
            ObjectPtr(.Target) = 0
            ObjectPtr(.owner) = 0
        End If
    End With
End Sub

Private Function Release(ByRef this As Long) As Long
    Dim tmpThis As Long
    tmpThis = this
    mWeak.SA.pvData = this
    With mWeak.WeakRef(0)
        this = .pOriginalVTable
        Release = .Target.Release
        .owner.Release Release
        If Release > 0 Then
            this = tmpThis
        Else
            ObjectPtr(.Target) = 0
            ObjectPtr(.owner) = 0
        End If
    End With
End Function

