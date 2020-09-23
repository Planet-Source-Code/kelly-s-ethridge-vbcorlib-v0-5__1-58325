Attribute VB_Name = "modFunctionDelegator"
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
'    Module: modFunctionDelegator
'
Option Explicit

Private Const DELEGATE_ASM As Currency = -368956918007638.6215@     ' from Matt Curland

Private Type DelegatorVTable
    Func(3) As Long
End Type
Public Type FunctionDelegator
    pVTable As Long
    pfn As Long
    cRefs As Long
End Type

Private mDelegateASM As Currency
Private mDelegatorVTable As DelegatorVTable
Private mpDelegatorVTable As Long
Private mUserDelegatorVTable As DelegatorVTable
Private mpUserDelegatorVTable As Long

Public Function InitDelegator(ByRef Delegator As FunctionDelegator, Optional ByVal pfn As Long = 0) As IUnknown
    If mpUserDelegatorVTable = 0 Then InitUserVTable
    With Delegator
        .pfn = pfn
        .pVTable = mpUserDelegatorVTable
    End With
    ObjectPtr(InitDelegator) = VarPtr(Delegator)
End Function
Private Sub InitUserVTable()
    mDelegateASM = DELEGATE_ASM
    With mUserDelegatorVTable
        .Func(0) = FuncAddr(AddressOf InitDelegator_QueryInterface)
        .Func(1) = FuncAddr(AddressOf InitDelegator_AddRefRelease)
        .Func(2) = .Func(1)
        .Func(3) = VarPtr(mDelegateASM)
        mpUserDelegatorVTable = VarPtr(.Func(0))
    End With
End Sub
Private Function InitDelegator_QueryInterface(ByRef This As FunctionDelegator, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = VarPtr(This)
End Function
Private Function InitDelegator_AddRefRelease(ByRef This As FunctionDelegator) As Long
    ' do nothing
End Function



Public Function NewDelegator(ByVal pfn As Long) As IUnknown
    Dim This As Long
    Dim Struct As FunctionDelegator
    
    If mpDelegatorVTable = 0 Then InitVTable
    This = CoTaskMemAlloc(LenB(Struct))
    If This = 0 Then Err.Raise 7
    
    With Struct
        .pVTable = mpDelegatorVTable
        .cRefs = 1
        .pfn = pfn
    End With
    CopyMemory ByVal This, Struct, LenB(Struct)
    ObjectPtr(NewDelegator) = This
End Function
Private Sub InitVTable()
    mDelegateASM = DELEGATE_ASM
    With mDelegatorVTable
        .Func(0) = FuncAddr(AddressOf QueryInterface)
        .Func(1) = FuncAddr(AddressOf AddRef)
        .Func(2) = FuncAddr(AddressOf Release)
        .Func(3) = VarPtr(mDelegateASM)
        mpDelegatorVTable = VarPtr(.Func(0))
    End With
End Sub
Private Function QueryInterface(ByRef This As FunctionDelegator, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = VarPtr(This)
    AddRef This
End Function
Private Function AddRef(ByRef This As FunctionDelegator) As Long
    With This
        .cRefs = .cRefs + 1
        AddRef = .cRefs
    End With
End Function
Private Function Release(ByRef This As FunctionDelegator) As Long
    With This
        .cRefs = .cRefs - 1
        Release = .cRefs
        If .cRefs = 0 Then CoTaskMemFree VarPtr(This)
    End With
End Function

