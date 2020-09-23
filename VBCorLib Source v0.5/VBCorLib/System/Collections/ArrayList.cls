VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayList"
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
'    Module: ArrayList
'

''
' A list that dynamically increases in size to hold a number of items.
'
' <H4>Remarks</H4> As new items are added to the list, the size will increase
' as necessary. The capacity of the list is the current number of elements
' in the internal-array. To lower the capacity to save memory, the TrimToSize
' method can be called, or Capacity can be set directly. The default capacity is 16.
' The Item index is zero based.
'
' <br><H4>Constructors:</H4><br>
' <font face="Courier New">
' Set obj = New ArrayList<br>
' Set obj = NewArrayList([Comparer], [c])<br>
' Set obj = Cor.NewArrayList([Comparer], [c])<br>
' </font><br>
' <b>Arguments:</b><br>
' <i>Comparer</i>: A user-supplied object that impelments IComparer that the ArrayList
' will use when searching and sorting the list.<br>
' <i>c</i>: An initial collection to place in the ArrayList when being created. The valid
' types of collections are VBA.Collection, objects that implement <b>ICollection</b> and arrays.
'
Option Explicit
Implements cObject
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList

Private Const PROP_COUNT            As String = "Count"
Private Const PROP_CAPACITY         As String = "Capacity"
Private Const PROP_ITEMPREFIX       As String = "Item"
Private Const PROP_DEFAULTCOMPARER  As String = "DefaultComparer"
Private Const PROP_COMPARER         As String = "Comparer"

Private Const DEF_CAPACITY          As Long = 16

Private mItems() As Variant
Private mCount As Long
Private mCapacity As Long
Private mVersion As Long
Private mComparer As IComparer



''
' Adds a new item to the list.
'
' @param value The item to be added to the list.
' @return The index at which the item was added.
' @remarks As items are added, the capacity is increased as necessary.
'
Public Function Add(ByRef value As Variant) As Long
    Add = mCount
    Insert mCount, value
End Function

''
' Adds a collection of items to the list.
'
' @param c The collection of items to add.
' The collection can be a <b>VBA.Collection</b>,<b>ICollection</b> object,
' or an <b>Array</b>.
' @remarks The capacity is increased as necessary to hold the new values.
'
Public Sub AddRange(ByRef c As Variant)
    InsertRange mCount, c
End Sub

''
' Performs a binary search for the value in the internal list.
'
' @param value The value to search for.
' @param startindex  The starting index to begin the search.
' @param count The number of elements to search in from the startindex.
' @param comparer A comparer to perform user-defined comparison logic.
' @return The index at which the value was found, or a negative value indicating
' the value was not found.
' @remarks <p>The search can be limited to a specific range within the list instead
' of over the entire list.</p>
'
' <p>If there is special compare logic necessary, then a comparer object can be supplied.
' The compare object must implement the IComparer interface.</P>
' <p>If the return value is negative, then the value was not found in the list. To
' determine where the value should have been found, negate (Not) the return value.</p>
' <p>The search assumes the array is in sorted order.</p>
'
Public Function BinarySearch(ByRef value As Variant, Optional ByVal StartIndex As Long, Optional ByVal Count As Long = -1, Optional ByVal comparer As IComparer) As Long
    If mCount = 0 Then
        BinarySearch = -1
        Exit Function
    End If
    If Count < 0 Then Count = mCount
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "startindex"), "startindex", StartIndex)
    If StartIndex + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    If comparer Is Nothing Then Set comparer = mComparer
    BinarySearch = cArray.BinarySearch(mItems, value, StartIndex, Count, comparer)
End Function

''
' Returns the total number of elements in the the internal array.
'
' @return The number of elements allocated internally.
'
Public Property Get capacity() As Long
    capacity = mCapacity
End Property

''
' Sets the number of allocated elements in the array.
'
' @param RHS The number of elements to set the array to.
' @remarks The capacity cannot be set less than the number of items in the list.
' If the capacity is set to zero, then it is set to the default capacity (16).
'
Public Property Let capacity(ByVal RHS As Long)
    If RHS < mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_SmallCapacity), "Capacity", capacity)
    If RHS = 0 Then RHS = DEF_CAPACITY
    If RHS = mCapacity Then Exit Property
    mCapacity = RHS
    ReDim Preserve mItems(0 To mCapacity - 1)
End Property

''
' Clears the list of all values.
'
' @remarks The internal array is cleared of values. The count is reset
' to zero, but the capacity remains the same. To lower the number of
' elements in the array, use TrimToSize or set Capacity directly.
'
Public Sub Clear()
    ReDim mItems(0 To mCapacity - 1)
    mCount = 0
    mVersion = mVersion + 1
End Sub

''
' Returns an ArrayList object containing a copy of the array in
' the original ArrayList.
'
Public Function Clone() As Object
    Dim ret As New ArrayList
    ret.CloneHelper mItems, mCount
    Set Clone = ret
End Function

''
' Returns if the list contains a specific value.
'
' @param value The value to search for.
' @param comparer  A user-defined comparer to performs custom
' compare logic on the array elements.
' @return A boolean value indicating the value was found in the list.
' @remarks If there is special compare logic necessary, then a comparer
' object can be supplied. The compare object must implement the IComparer interface.
'
Public Function Contains(ByRef value As Variant, Optional ByVal comparer As IComparer) As Boolean
    Contains = (IndexOf(value, , , comparer) >= 0) ' (cArray.IndexOf(mItems, value, , , comparer) >= 0)
End Function

''
' Copies all of the items to an array.
'
' @param dstArray The array to copy the items to.
' @param arrayindex  The index to start copying items to in the destination array.
' @remarks The destination array must be large enough to hold all of the items. If an
' arrayindex is specified, then the destination array must be large enough to hold all of
' the items from that index on.
'
Public Sub CopyTo(ByRef dstArray As Variant, Optional ByRef arrayindex As Variant)
    If IsMissing(arrayindex) Then arrayindex = LBound(dstArray)
    cArray.CopyEx mItems, 0, dstArray, arrayindex, mCount
End Sub

''
' Copies all of the items to an Array.
'
' @param index The starting index in the list to start copying from.
' @param dstArray The array to copy the items to.
' @param arrayindex The index in the destination array to start copying items to.
' @param count The number of items to copy.
' @remarks The destination array must be large enough to hold the number of elements being copied.
'
Public Sub CopyToEx(ByVal Index As Long, ByRef dstArray As Variant, ByVal arrayindex As Long, ByVal Count As Long)
    cArray.CopyEx mItems, Index, dstArray, arrayindex, Count
End Sub

''
' Returns the number of items in the list.
'
' @return The number of items in the list.
'
Public Property Get Count() As Long
    Count = mCount
End Property

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef value As Variant) As Boolean
    If IsObject(value) Then
        Equals = (value Is Me)
    End If
End Function

''
' Returns an enumerator for an ArrayList.
'
' @param startindex  The index in the list to being enumeration.
' @param count  The number of items in the list to enumerate over.
' @return An ArrayList enumerator
' @remarks This enumerator can be used in For..Each loops. To access the optional
' parameters, the GetEnumerator must be called instead of simply passing the object
' to the For..Each. This allows for a portion of the list to be enumerated.<br><br>
'
' <b>Example:</b><br>
'   Normal enumeration<br>
'   <font face="Courier New">
'   For Each v In list<br>
'   &nbsp;&nbsp&nbsp&nbsp...<br>
'   Next v<br>
'   </font>
'<br>
'   Parameterized enumeration<br>
'   <font face="Courier New">
'   For Each v in list.GetEnumerator(0,10)<br>
'   &nbsp&nbsp&nbsp&nbsp...<br>
'   Next v<br>
'   </font>
'
Public Function GetEnumerator(Optional ByVal StartIndex As Long, Optional ByVal Count As Long = -1) As Object
    If Count < 0 Then Count = mCount
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "startindex"), "startindex", StartIndex)
    If StartIndex + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    
    Dim ret As ArrayListEnumerator
    Set ret = New ArrayListEnumerator
    ret.Init Me, StartIndex, Count
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
' Returns an ArrayList object that is a windowed view into the original ArrayList.
'
' @param index The index in the original list to start the new view.
' @param count The number of elements in the original list to show in the view.
' @return An ArrayList that represents a subset view of the original list.
' @remarks  This windowed view is the only way to then modify the original
' ArrayList. If the original list is modified some other way, then an exception
' is thrown the next time the ranged view is accessed.
'
Public Function GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    If Index < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, IIf(Index < 0, "index", "count")), IIf(Index < 0, "index", "count"), IIf(Index < 0, Index, Count))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Argument_InvalidCountOffset, "count")
    Dim ret As RangedArrayList
    Set ret = New RangedArrayList
    ret.Init Me, Index, Count
    Set GetRange = ret
End Function

''
' Returns the index of the first occurrence of the value in the list.
'
' @param value The value to search the list for.
' @param startindex The index to begin the search in the list.
' @param count The number of items in the list to search.
' @param comparer A custom comparer to perform any special compare logic.
' @return The index of the first occurrence of value, or -1 if the value was not found.
' @remarks If there is special compare logic necessary, then a comparer
' object can be supplied. The compare object must implement the IComparer interface.
'
Public Function IndexOf(ByRef value As Variant, Optional ByVal StartIndex As Long, Optional ByVal Count As Long = -1, Optional ByVal comparer As IComparer) As Long
    If mCount = 0 Then
        IndexOf = -1
        Exit Function
    End If
    If Count < 0 Then Count = mCount
    If StartIndex + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    If comparer Is Nothing Then Set comparer = mComparer
    IndexOf = cArray.IndexOf(mItems, value, StartIndex, Count, comparer)
End Function

''
' Inserts a value into the list at the specified index.
'
' @param index The index in which the value is to be inserted.
' @param value The value to be inserted into the list.
' @remarks When a value is inserted, all items starting at the insertion point
' are moved out by one space. If capacity is reached, the list will be reallocated.
'
Public Sub Insert(ByVal Index As Long, ByRef value As Variant)
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_ArrayListInsert), "index", Index)
    If mCount = mCapacity Then EnsureCapacity mCount + 1
    InsertSpace Index, 1
    VariantCopyInd mItems(Index), value
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

''
' Inserts a collection of items into the list.
'
' @param index The index at which the items will be inserted.
' @param c The collection of items to be inserted into the list.
' @remarks When the items are inserted, the items starting at the
' index will be moved out a number of spaced equal to the number of
' items to be inserted. If the capacity is reached, the list will
' be reallocated.
'
' The collection can be a <b>VBA.Collection</b>,<b>ICollection</b> object,
' or an <b>Array</b>.
'
Public Sub InsertRange(ByVal Index As Long, ByRef c As Variant)
    Dim v As Variant
    Dim size As Long
    'Dim obj As Object
    
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_ArrayListInsert), "index", Index)
    
    size = GetCollectionAndSize(c)
    If size = 0 Then Exit Sub
    EnsureCapacity mCount + size
    InsertSpace Index, size
    For Each v In c
        Helper.MoveVariant mItems(Index), v
        Index = Index + 1
    Next v
    
    mCount = mCount + size
    mVersion = mVersion + 1
End Sub

''
' Returns if the list is fixed-size.
'
' @return Value indicating if the list is fixed-size.
' @remarks All operations on the list can be executed as long as the number
' of items don't change, this includes the capacity of the underlying array.
'
Public Property Get IsFixedSize() As Boolean
    IsFixedSize = False
End Property

''
' Returns if the list is read-only.
'
' @return Value indicating if the list is read-only.
' @remarks All operations on the list can be executed as long as the number
' of items don't change, and the items aren't changed in the list.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Returns an item from the list.
'
' @param index The index in the list from which to return the item.
'
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If Index < 0 Or Index >= mCount Then Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index))
    VariantCopyInd Item, mItems(Index)
End Property

''
' Sets an item in the list to a value.
'
' @param index The index in the list to set the value.
' @param RHS The value to set in the list.
'
Public Property Let Item(ByVal Index As Long, ByRef RHS As Variant)
    If Index < 0 Or Index >= mCount Then Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index))
    mItems(Index) = RHS
    mVersion = mVersion + 1
End Property

''
' Sets an item in the list to a value.
'
' @param index The index in the list to set the value.
' @param RHS The value to set in the list.
'
Public Property Set Item(ByVal Index As Long, ByRef RHS As Variant)
    If Index < 0 Or Index >= mCount Then Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index))
    Set mItems(Index) = RHS
    mVersion = mVersion + 1
End Property

''
' Returns the last occurrence of a value in the list.
'
' @param value The value to search for in the list.
' @param startindex The index to start searching from.
' @param count The number of items to be searched.
' @param comparer A custom comparer to perform any special compare logic.
' @return The index of the last occurrence of the value, or -1 if not found.
' @remarks The search is performed from the end of the list towards the
' beginning. If a startindex
'
Public Function LastIndexOf(ByRef value As Variant, Optional ByVal StartIndex As Long = -1, Optional ByVal Count As Long = -1, Optional ByVal comparer As IComparer) As Long
    If mCount = 0 Then
        LastIndexOf = -1
        Exit Function
    End If
    If Count < 0 Then Count = mCount
    If StartIndex < 0 Then StartIndex = mCount - 1
    If StartIndex >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LargerThanCollection), "startindex", StartIndex)
    If StartIndex - Count + 1 < 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    If comparer Is Nothing Then Set comparer = mComparer
    LastIndexOf = cArray.LastIndexOf(mItems, value, StartIndex, Count, comparer)
End Function

''
' Returns an enumerator
'
' @return The enumerator
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Removes a value from the list.
'
' @param value The value to be removed.
' @param comparer A custom comparer used to find the value.
' @remarks If the value is an object, then it must be the same instance of
' the object to be removed. This behaviour can be overriden with a custom comparer.
'
Public Sub Remove(ByRef value As Variant, Optional ByVal comparer As IComparer)
    Dim i As Long
    i = IndexOf(value, , , comparer)
    If i >= 0 Then RemoveAt i
End Sub

''
' Removes a value from the list at a specific index.
'
' @param index The index in the list to remove the item.
'
Public Sub RemoveAt(ByVal Index As Long)
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    mItems(Index) = Empty
    RemoveSpace Index, 1
    mCount = mCount - 1
    mVersion = mVersion + 1
End Sub

''
' Remove a set of items from the list.
'
' @param index The index to start removing items.
' @param count The number of items to remove.
'
Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)
    If Index < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, IIf(Index < 0, "index", "count")), IIf(Index < 0, "index", "count"), IIf(Index < 0, Index, Count))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    cArray.Clear mItems, Index, Count
    RemoveSpace Index, Count
    mCount = mCount - Count
    mVersion = mVersion + 1
End Sub

''
' Reverses the list of items in the list.
'
' @param index The index to start reversing items.
' @param count The number of items to include in the reverse.
'
Public Sub Reverse(Optional ByVal Index As Long, Optional ByVal Count As Long = -1)
    If Count < 0 Then Count = mCount
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    cArray.Reverse mItems, Index, Count
    mVersion = mVersion + 1
End Sub

''
' Sets the items in the list to a collection of items.
'
' @param index The index at which to start setting items.
' @param c A collection of items to set in the list.
' @remarks The number of items in the list is not changed, nor is the size
' of the list increased. Only existing items in the list can be set to
' the items in the collection. This basically overlays new items onto
' existing items.
'
Public Sub SetRange(ByVal Index As Long, ByRef c As Variant)
    Dim size As Long
    Dim v As Variant
    
    size = GetCollectionAndSize(c)
    If Index < 0 Or Index + size > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index")
    
    For Each v In c
        Helper.MoveVariant mItems(Index), v
        Index = Index + 1
    Next v
    mVersion = mVersion + 1
End Sub

''
' Sorts the items in the list.
'
' @param startindex The index to start sorting the items.
' @param count The number of items to include in the sort.
' @param comparer A custom comparer used to compare items in the list.
' @remarks The list is sorted in ascending order by default. By providing
' a custom comaprer, the list can be sorted in special ways, such as descending order.
'
Public Sub Sort(Optional ByVal StartIndex As Long, Optional ByVal Count As Long = -1, Optional ByVal comparer As IComparer)
    If Count < 0 Then Count = mCount
    If mCount = 0 Then Exit Sub
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "startindex"), "startindex", StartIndex)
    If StartIndex + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    If comparer Is Nothing Then Set comparer = mComparer
    cArray.SortEx mItems, StartIndex, Count, comparer
    mVersion = mVersion + 1
End Sub

''
' Returns an array of the items in the list.
'
' @param arraytype The specific array type to return.
' @return An array of items in the list.
' @remarks <p>If an array type other than variant is specified, then the
' list attempts to convert each item in the list to the specific array type
' requested. If an item cannot be converted, then an exception is thrown.</p>
' <p>No items will be narrowed in the conversion. For example, a Double will
' not be converted to a Long. An exception will be thrown.
'
Public Function ToArray(Optional ByVal arraytype As ciArrayTypes = ciVariant) As Variant
    Dim ret As Variant
    
    If mCount = 0 Then
        ret = cArray.CreateInstance(arraytype)
    Else
        VariantType(ret) = arraytype Or vbArray
        MemLong(VarPtr(ret) + VARIANTDATA_OFFSET) = SafeArrayCreateVector(arraytype, 0, mCount)
        cArray.CopyEx mItems, 0, ret, 0, mCount
    End If
    ToArray = ret
End Function

''
' Sets the capacity to the number of items in the list.
'
' @remarks This will decrease the size of the internal array. This can save
' memory if you know that no more items will be added to the list.
'
Public Sub TrimToSize()
    capacity = mCount
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Returns the current version of the instance.
'
' @return Version count.
' @remarks This is used to track if the list has been modified. Wrappers
' for an ArrayList may need to keep in sync with the list. An example would
' be the ArrayListEnumerator. If the list is modified during enumeration, then
' the enumerator throws an error by checking to see if the version has changed
' since the enumeration began. Since other objects implement the ArrayList
' interface and need to be enumerated using the same ArrayListEnumerator,
' the enumerator can't access a Friend method of ArrayList.
'
Public Property Get Version() As Long
    Version = mVersion
End Property



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub CloneHelper(ByRef arr() As Variant, ByVal Count As Long)
    mItems = arr
    mCount = Count
    mCapacity = UBound(arr) + 1
End Sub

Friend Sub Init(ByRef comparer As IComparer, ByRef c As Variant)
    Set mComparer = comparer
    AddRange c
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureCapacity(ByVal requiredCapacity As Long)
    Dim newCapacity As Long
    If requiredCapacity < mCapacity Then Exit Sub
    newCapacity = mCapacity * 2
    If requiredCapacity > newCapacity Then newCapacity = requiredCapacity
    ReDim Preserve mItems(0 To newCapacity - 1)
    mCapacity = newCapacity
End Sub

Private Sub InsertSpace(ByVal Index As Long, ByVal size As Long)
    If Index < mCount Then
        CopyMemory mItems(Index + size), mItems(Index), (mCount - Index) * 16
        ZeroMemory mItems(Index), size * 16
    End If
End Sub

Private Sub RemoveSpace(ByVal Index As Long, ByVal size As Long)
    If Index < mCount - 1 Then
        CopyMemory mItems(Index), mItems(Index + size), (mCount - Index - size) * 16
        ZeroMemory mItems(mCount - size), size * 16
    End If
End Sub

Friend Function GetCollectionAndSize(ByRef c As Variant) As Long
    If IsArray(c) Then
        If cArray.IsNull(c) Then Exit Function
        GetCollectionAndSize = UBound(c) - LBound(c) + 1
    Else
        If c Is Nothing Then Exit Function
        If TypeOf c Is Collection Then
            Dim col As Collection
            Set col = c
            GetCollectionAndSize = col.Count
            Set c = col
        Else
            Dim icol As ICollection
            Set icol = c
            GetCollectionAndSize = icol.Count
            Set c = icol
        End If
    End If
End Function

Private Sub Class_InitProperties()
    ReDim mItems(0 To DEF_CAPACITY - 1)
    mCapacity = DEF_CAPACITY
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Dim i As Long
    
    On Error GoTo errTrap
    With PropBag
        mCount = .ReadProperty(PROP_COUNT, 0)
        EnsureCapacity .ReadProperty(PROP_CAPACITY, DEF_CAPACITY)
        For i = 0 To mCount - 1
            Helper.MoveVariant mItems(i), .ReadProperty(PROP_ITEMPREFIX & i, Empty)
        Next i
        If .ReadProperty(PROP_DEFAULTCOMPARER, False) Then
            Set mComparer = comparer.Default
        Else
            Set mComparer = .ReadProperty(PROP_COMPARER, Nothing)
        End If
    End With
    Exit Sub
    
errTrap:
    Throw Cor.NewSerializationException(Err.Description)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Dim i As Long
    
    On Error GoTo errTrap
    With PropBag
        For i = 0 To mCount - 1
            .WriteProperty PROP_ITEMPREFIX & i, mItems(i)
        Next i
        .WriteProperty PROP_COUNT, mCount
        .WriteProperty "capacity", mCapacity
        If mComparer Is comparer.Default Then
            .WriteProperty PROP_DEFAULTCOMPARER, True
        Else
            .WriteProperty PROP_COMPARER, mComparer, Nothing
        End If
    End With
    Exit Sub
    
errTrap:
    Throw Cor.NewSerializationException(Err.Description)
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
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IList_Add(value As Variant) As Long
    IList_Add = Add(value)
End Function

Private Sub IList_Clear()
    Clear
End Sub

Private Function IList_Contains(value As Variant, Optional ByVal comparer As IComparer) As Boolean
    IList_Contains = Contains(value, comparer)
End Function

Private Sub IList_CopyTo(arr As Variant, ByVal Index As Long)
    CopyTo arr, Index
End Sub

Private Property Get IList_Count() As Long
    IList_Count = Count
End Property

Private Function IList_GetEnumerator() As IEnumerator
    Set IList_GetEnumerator = GetEnumerator
End Function

Private Function IList_IndexOf(value As Variant, Optional ByVal comparer As IComparer) As Long
    IList_IndexOf = IndexOf(value, , , comparer)
End Function

Private Sub IList_Insert(ByVal Index As Long, value As Variant)
    Insert Index, value
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = IsFixedSize
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = IsReadOnly
End Property

Private Property Set IList_Item(ByVal Index As Long, RHS As Variant)
    Set Item(Index) = RHS
End Property

Private Property Let IList_Item(ByVal Index As Long, RHS As Variant)
    Item(Index) = RHS
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Helper.MoveVariant IList_Item, Item(Index)
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = NewEnum
End Function

Private Sub IList_Remove(value As Variant, Optional ByVal comparer As IComparer)
    Remove value, comparer
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    RemoveAt Index
End Sub
