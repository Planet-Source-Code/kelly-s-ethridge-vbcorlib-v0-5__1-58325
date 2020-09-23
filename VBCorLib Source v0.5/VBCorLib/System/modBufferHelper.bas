Attribute VB_Name = "modBufferHelper"
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
'    Module: modBufferHelper
'
Option Explicit

Public Type WordBuffer
    pVTable As Long
    this As IUnknown
    Data() As Integer
    SA As SafeArray1d
End Type
Private Type WordVTableType
    Func(2) As Long
End Type

Private mpWordVTable As Long
Private mWordVTable As WordVTableType



Public Sub InitWordBuffer(ByRef Buffer As WordBuffer, ByVal pData As Long, ByVal length As Long)
    If mpWordVTable = 0 Then
        With mWordVTable
            .Func(0) = FuncAddr(AddressOf WordBuffer_QueryInterface)
            .Func(1) = FuncAddr(AddressOf WordBuffer_AddRef)
            .Func(2) = FuncAddr(AddressOf WordBuffer_Release)
            mpWordVTable = VarPtr(.Func(0))
        End With
    End If
    With Buffer.SA
        .cbElements = 2
        .cDims = 1
        .cElements = length
        .pvData = pData
    End With
    With Buffer
        .pVTable = mpWordVTable
        SAPtr(.Data) = VarPtr(.SA)
        ObjectPtr(.this) = VarPtr(.pVTable)
    End With
End Sub

Private Function WordBuffer_QueryInterface(ByRef this As WordBuffer, ByRef riid As Long, ByRef pvObj As Long) As Long
    Debug.Assert False
    pvObj = 0
    WordBuffer_QueryInterface = E_NOINTERFACE
End Function

Private Function WordBuffer_AddRef(ByRef this As WordBuffer) As Long
    Debug.Assert False
End Function

Private Function WordBuffer_Release(ByRef this As WordBuffer) As Long
    SAPtr(this.Data) = 0
    this.SA.pvData = 0
End Function

