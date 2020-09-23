VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hashtable"
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
'    Module: Hashtable
'

''
' The Hashtable represents a repository of key/value pairs.
'
Option Explicit
Implements cObject
Implements IDictionary
Implements ICollection
Implements IEnumerable
Implements ICloneable

Private Const DEF_CAPACITY As Long = 32
Private Const LOAD_FACTOR As Double = 0.9

Private mCount As Long
Private mBuckets() As Bucket
Private mCapacity As Long
Private mVersion As Long
Private mLoadThreshold As Long
Private mProvider As IHashcodeProvider
Private mComparer As IComparer

' used in hashing a string key.
Private mStringHashChars As WordBuffer

' used to quickly empty a bucket and mark it deleted.
Private mDeletedBucket As Bucket


''
' Adds a new key/value pair to the table.
'
' @param Key The key used to identify the value.
' @param Value The value to be added to the table.
' @remarks The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' cObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the cObject
' interface and override the GetHashCode function.
'
Public Sub Add(ByRef Key As Variant, ByRef value As Variant)
    InsertItem Key, value, True, False
End Sub

''
' Clears all of the elements from the table.
'
' @remarks When clearing the table, the count is set to 0, but the
' capacity remains unchanged.
'
Public Sub Clear()
    ReDim mBuckets(0 To mCapacity - 1)
    mCount = 0
End Sub

''
' Returns a clone of the current instance.
'
' @return The clone of the current instance.
'
Public Function Clone() As Object
    Dim ret As New Hashtable
    ret.CloneHelper mBuckets, mCount
    Set Clone = ret
End Function

''
' Checks if the current instance contains a specific key.
'
' @param Key The key to check for.
' @return Indicates if the key was found.
' @remarks The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' cObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the cObject
' interface and override the GetHashCode function.
'
Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = ContainsKey(Key)
End Function

''
' Checks if the current instance contains a specific key.
'
' @param Key The key to check for.
' @return Indicates if the key was found.
' @remarks The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' cObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the cObject
' interface and override the GetHashCode function.
'
Public Function ContainsKey(ByRef Key As Variant) As Boolean
    ContainsKey = (GetKeyIndex(Key) >= 0)
End Function

''
' Checks if the current instance contains a specific value.
'
' @param Value The value to search for.
' @return Indicates if the value was found.
' @remarks The table is search linearly checking each value. Once
' a matching value is found, True is returned. If there is more than
' one of the same values, only the first is detected.
'
Public Function ContainsValue(ByRef value As Variant) As Boolean
    Dim i As Long
    
    For i = 0 To mCapacity - 1
        If mBuckets(i).State = bsOccupied Then
            If EqualsVariants(mBuckets(i).value, value) Then
                ContainsValue = True
                Exit Function
            End If
        End If
    Next i
End Function

''
' Copies the contents of the table to an array.
'
' @param dstArray The array to copy the contents to.
' @param index The starting index in dstArray to begin copying to.
' @remarks The elements that are copied to the array are DictionaryEntry
' elements. Each DictionaryEntry contains a key and value from the table.
'
Public Sub CopyTo(ByRef dstArray As Variant, ByVal Index As Long)
    InternalCopyTo dstArray, Index, etEntries
End Sub

''
' Returns the number of elements in the table.
'
' @return The number of elements in the table.
'
Public Property Get Count() As Long
    Count = mCount
End Property

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
' Returns an enumerator for the table.
'
' @return An IEnumerator object that enumerates over the table.
' @remarks The values enumerated are DictionaryEntry object each
' containing a key/value pair.
'
Public Function GetEnumerator() As Object
    Dim ret As New HashtableEnumerator
    ret.Init Me, SAPtr(mBuckets), etEntries
    Set GetEnumerator = ret
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function

''
' Returns if this instance is fixed size.
'
' @return Default is False.
' @remarks This is to satisfy the IList interface.
'
Public Property Get IsFixedSize() As Boolean
    IsFixedSize = False
End Property

''
' Returns if this instance is readonly
'
' @return Default is False.
' @remarks This is to satisfy the IList interface.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Returns a value associated with the specified key.
'
' @param Key The key used to look up a value.
' @return The value associated with the Key.
' @remarks If the key does not exist, then Empty is returned.
' <p>The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' cObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the cObject
' interface and override the GetHashCode function.</p>
'
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim i As Long
    i = GetKeyIndex(Key)
    If i >= 0 Then VariantCopy Item, mBuckets(i).value
End Property

''
' Sets the value associated to the specified key.
'
' @param Key The key used to look up a value.
' @param Value The value to associate with the key.
' @remarks If the key is not found, then the key/value are added to
' the table. If the key IS found, then the associated value with the
' key is replaced with the new value.
' <p>The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' cObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the cObject
' interface and override the GetHashCode function.</p>
'
Public Property Let Item(ByRef Key As Variant, ByRef value As Variant)
    InsertItem Key, value, False, False
End Property

''
' Sets the value associated to the specified key.
'
' @param Key The key used to look up a value.
' @param Value The value to associate with the key.
' @remarks If the key is not found, then the key/value are added to
' the table. If the key IS found, then the associated value with the
' key is replaced with the new value.
' <p>The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' cObject interface to allow for custom hashcode creation. If the
' object does not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the cObject
' interface and override the GetHashCode function.</p>
'
Public Property Set Item(ByRef Key As Variant, ByRef value As Variant)
    InsertItem Key, value, False, False
End Property

''
' Retuns an ICollection object used to access the keys of the table.
'
' @return An ICollection object used to access the keys of the table.
'
Public Property Get Keys() As ICollection
    Dim ret As New HTKeyCollection
    ret.Init Me
    Set Keys = ret
End Property

''
' Returns an enumerator for the table.
'
' @return The enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Removes a key/value pair associated with a specific key.
'
' @param Key The key use to lookup the key/value pair to be removed.
'
Public Sub Remove(ByRef Key As Variant)
    Dim i As Long
    i = GetKeyIndex(Key)
    If i >= 0 Then
        mBuckets(i) = mDeletedBucket
        mCount = mCount - 1
        mVersion = mVersion + 1
    End If
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Retuns an ICollection object used to access the values of the table.
'
' @return An ICollection object used to access the values of the table.
'
Public Property Get Values() As ICollection
    Dim ret As New HTValueCollection
    ret.Init Me
    Set Values = ret
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub CloneHelper(ByRef buckets() As Bucket, ByVal Count As Long)
    mCount = Count
    mBuckets = buckets
    mCapacity = UBound(buckets) + 1
End Sub

Friend Sub CopyKeys(ByRef dstArray As Variant, ByVal Index As Long)
    InternalCopyTo dstArray, Index, etKeys
End Sub

Friend Sub CopyValues(ByRef dstArray As Variant, ByVal Index As Long)
    InternalCopyTo dstArray, Index, etValues
End Sub

Friend Function GetKeyEnumerator() As IEnumerator
    Dim ret As New HashtableEnumerator
    ret.Init Me, SAPtr(mBuckets), etKeys
    Set GetKeyEnumerator = ret
End Function

Friend Function GetValueEnumerator() As IEnumerator
    Dim ret As New HashtableEnumerator
    ret.Init Me, SAPtr(mBuckets), etValues
    Set GetValueEnumerator = ret
End Function

Friend Sub InitCapacity(ByVal capacity As Long, ByVal hcp As IHashcodeProvider, ByVal comparer As IComparer)
    mCapacity = GetPrime(capacity)
    mLoadThreshold = mCapacity * LOAD_FACTOR
    ReDim mBuckets(0 To mCapacity - 1)
    Set mProvider = hcp
    Set mComparer = comparer
End Sub

Friend Sub InitDictionary(ByVal dic As IDictionary, ByVal hcp As IHashcodeProvider, ByVal comparer As IComparer)
    InitCapacity dic.Count, hcp, comparer
    Dim Entry As DictionaryEntry
    For Each Entry In dic
        InsertItem Entry.Key, Entry.value, True, True
    Next Entry
End Sub

Friend Property Get Version() As Long
    Version = mVersion
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExpandTable()
    Dim OldBuckets() As Bucket
    Dim i As Long
    
    mCount = 0
    mCapacity = GetPrime(mCapacity * 2)
    mLoadThreshold = mCapacity * LOAD_FACTOR
    Helper.Swap4 ByVal ArrPtr(OldBuckets), ByVal ArrPtr(mBuckets)
    ReDim mBuckets(0 To mCapacity - 1)
    For i = 0 To UBound(OldBuckets)
        If OldBuckets(i).State = bsOccupied Then
            InsertItem OldBuckets(i).Key, OldBuckets(i).value, True, True
        End If
    Next i
End Sub

Private Function GetKeyIndex(ByRef Key As Variant) As Long
    Dim incr As Long
    Dim Index As Long
    Dim hashcode As Long
    
    hashcode = InitHashCode(Key, incr, Index)
    Do
        Select Case mBuckets(Index).State
            Case bsEmpty
                GetKeyIndex = -1
                Exit Function
            Case bsOccupied
                If mBuckets(Index).hashcode = hashcode Then
                    If mComparer Is Nothing Then
                        If EqualsVariants(Key, mBuckets(Index).Key) Then
                            GetKeyIndex = Index
                            Exit Function
                        End If
                    ElseIf mComparer.Compare(Key, mBuckets(Index).Key) = 0 Then
                        GetKeyIndex = Index
                        Exit Function
                    End If
                End If
        End Select
        Index = (Index + incr) Mod mCapacity
    Loop
End Function

Private Function HashKey(ByRef Key As Variant) As Long
    Select Case VarType(Key)
        Case vbLong, vbInteger, vbByte, vbBoolean
            HashKey = Key
        Case vbString
            HashKey = HashString(Key)
        Case vbDouble, vbDate, vbCurrency
            With AsDLong(ByVal VarPtr(Key) + 8)
                HashKey = .LoDWord Xor .HiDWord
            End With
        Case vbObject, vbDataObject
            If TypeOf Key Is cObject Then
                Dim o As cObject
                Set o = Key
                HashKey = o.GetHashCode
            Else
                HashKey = ObjPtr(Key)
            End If
        Case Else
            Debug.Assert False
    End Select
End Function

Private Function HashString(ByRef Key As Variant) As Long
    Dim i As Long
    Dim l As Long
    
    l = Len(Key)
    With mStringHashChars.SA
        .cElements = l
        .pvData = StrPtr(Key)
    End With
    For i = 0 To l - 1
       HashString = ((HashString * 16) + mStringHashChars.data(i)) And &H3FFFFFF
    Next i
End Function

Private Function InitHashCode(ByRef Key As Variant, ByRef incr As Long, ByRef Index As Long)
    If mProvider Is Nothing Then
        InitHashCode = HashKey(Key) And &H7FFFFFFF
    Else
        InitHashCode = mProvider.GetHashCode(Key) And &H7FFFFFFF
    End If
    incr = 1 + (((InitHashCode \ 32) + 1) Mod (mCapacity - 1))
    Index = InitHashCode Mod mCapacity
End Function

Private Sub InsertItem(ByRef Key As Variant, ByRef value As Variant, ByVal Adding As Boolean, ByVal steal As Boolean)
    Dim incr As Long
    Dim Index As Long
    Dim firstEmpty As Long
    Dim hashcode As Long
    
    If mCount > mLoadThreshold Then ExpandTable
    
    hashcode = InitHashCode(Key, incr, Index)
    firstEmpty = -1
    Do
        Select Case mBuckets(Index).State
            Case bsEmpty
                If firstEmpty = -1 Then firstEmpty = Index
                With mBuckets(firstEmpty)
                    If steal Then
                        Helper.MoveVariant .Key, Key
                        Helper.MoveVariant .value, value
                    Else
                        VariantCopyInd .Key, Key
                        VariantCopyInd .value, value
                    End If
                    .State = bsOccupied
                    .hashcode = hashcode
                End With
                mVersion = mVersion + 1
                mCount = mCount + 1
                Exit Sub
            Case bsOccupied
                If mBuckets(Index).hashcode = hashcode Then
                    If mComparer Is Nothing Then
                        If EqualsVariants(Key, mBuckets(Index).Key) Then
                            If Adding Then Throw Cor.NewArgumentException("An element with the same key already exists in the collection.", "key")
                            VariantCopyInd mBuckets(Index).value, value
                            mVersion = mVersion + 1
                            Exit Sub
                        End If
                    ElseIf mComparer.Compare(Key, mBuckets(Index).Key) = 0 Then
                        If Adding Then Throw Cor.NewArgumentException("An element with the same key already exists in the collection.", "key")
                        VariantCopyInd mBuckets(Index).value, value
                        mVersion = mVersion + 1
                        Exit Sub
                    End If
                End If
            Case bsDeleted
                If firstEmpty = -1 Then firstEmpty = Index
        End Select
        Index = (Index + incr) Mod mCapacity
    Loop
End Sub

Private Sub InternalCopyTo(ByRef dstArray As Variant, ByVal Index As Long, ByVal CopyType As HTEnumTypeEnum)
    If cArray.IsNull(dstArray) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "dstArray")
    If Index < LBound(dstArray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "index", Index)
    If Index + mCount - 1 > UBound(dstArray) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "index")
    If cArray.GetRank(dstArray) <> 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    
    Dim i As Long
    Dim j As Long
    For i = 0 To mCapacity - 1
        If mBuckets(i).State = bsOccupied Then
            Select Case CopyType
                Case etEntries
                    Set dstArray(Index + j) = Cor.NewDictionaryEntry(mBuckets(i).Key, mBuckets(i).value)
                Case etKeys
                    If IsObject(mBuckets(i).Key) Then
                        Set dstArray(Index + j) = mBuckets(i).Key
                    Else
                        dstArray(Index + j) = mBuckets(i).Key
                    End If
                Case etValues
                    If IsObject(mBuckets(i).value) Then
                        Set dstArray(Index + j) = mBuckets(i).value
                    Else
                        dstArray(Index + j) = mBuckets(i).value
                    End If
            End Select
            j = j + 1
            If j = mCount Then Exit Sub
        End If
    Next i
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitCapacity DEF_CAPACITY, Nothing, Nothing
    mDeletedBucket.State = bsDeleted
    InitWordBuffer mStringHashChars, 0, 0
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICollection_CopyTo(arr As Variant, ByVal Index As Long)
    CopyTo arr, Index
End Sub

Private Property Get ICollection_Count() As Long
    ICollection_Count = Count
End Property

Private Function ICollection_GetEnumerator() As IEnumerator
    Set ICollection_GetEnumerator = GetEnumerator
End Function

Private Function ICollection_NewEnum() As stdole.IUnknown
    Set ICollection_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IDictionary Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IDictionary_Add(Key As Variant, value As Variant)
    Add Key, value
End Sub

Private Sub IDictionary_Clear()
    Clear
End Sub

Private Function IDictionary_Contains(Key As Variant) As Boolean
    IDictionary_Contains = Contains(Key)
End Function

Private Sub IDictionary_CopyTo(arr As Variant, ByVal Index As Long)
    CopyTo arr, Index
End Sub

Private Property Get IDictionary_Count() As Long
    IDictionary_Count = Count
End Property

Private Function IDictionary_GetEnumerator() As IEnumerator
    Set IDictionary_GetEnumerator = GetEnumerator
End Function

Private Property Get IDictionary_IsFixedSize() As Boolean
    IDictionary_IsFixedSize = IsFixedSize
End Property

Private Property Get IDictionary_IsReadOnly() As Boolean
    IDictionary_IsReadOnly = IsReadOnly
End Property

Private Property Set IDictionary_Item(Key As Variant, RHS As Variant)
    Set Item(Key) = RHS
End Property

Private Property Let IDictionary_Item(Key As Variant, RHS As Variant)
    Item(Key) = RHS
End Property

Private Property Get IDictionary_Item(Key As Variant) As Variant
    Helper.MoveVariant IDictionary_Item, Item(Key)
End Property

Private Property Get IDictionary_Keys() As ICollection
    Set IDictionary_Keys = Keys
End Property

Private Function IDictionary_NewEnum() As stdole.IUnknown
    Set IDictionary_NewEnum = NewEnum
End Function

Private Sub IDictionary_Remove(Key As Variant)
    Remove Key
End Sub

Private Property Get IDictionary_Values() As ICollection
    Set IDictionary_Values = Values
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function
