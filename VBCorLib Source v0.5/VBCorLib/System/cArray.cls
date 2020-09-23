VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArray"
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
'    Module: cArray
'

''
' Provides methods for manipulating, searching and sorting arrays.
'
' <h4>Usage:</h4><br>
' This class cannot be instantiated. The methods are accessed directly throw the
' class name in the form: cArray.*<br><br>
' <font face="Courier New">
' NumOfDims = cArray.GetRank(<i>arr</i>)<br>
' </font>
Option Explicit

Private Const PVDATA_OFFSET As Long = 12


'#If INDEBUG Then
'' this is used in tests with user-defined types.
'Public Type TestType
'    value As Long
'    strValue As String
'    objValue As Object
'End Type
'#End If

' extended safearray descriptor the includes the 16 bytes
' preceding the array descriptor. The 16 bytes contains
' a GUID for objects and user-defined types. This is to
' help in the correct copying of those types of elements.
Private Type GuidSafeArray1d
    Guid As VBGUID
    SA As SafeArray1d
End Type

''
' Represents the available array types that can be created using
' CreateInstance. The values corrispond the vbVarType enum.
'
Public Enum ciArrayTypes
    ciLong = vbLong
    ciInteger = vbInteger
    ciByte = vbByte
    ciDouble = vbDouble
    ciSingle = vbSingle
    ciBoolean = vbBoolean
    ciDate = vbDate
    ciString = vbString
    ciCurrency = vbCurrency
    ciVariant = vbVariant
    ciObject = vbObject
End Enum



Private mBinarySearchDelegator As FunctionDelegator
Private mBinarySearchComparer As ITwoArgReturnLong

Private mIndexOfDelegator As FunctionDelegator
Private mIndexOfComparer As ITwoRefReturnBool

Private mAssignDelegator As FunctionDelegator
Private mAssigner As ITwoArgReturnVoid

Private mSortDelegator As FunctionDelegator
Private mSorter As ISortRoutine

''
' Performs a binary search on a given array. A subportion of the array can
' be searched using the startindex and length parameters. A custom user
' comparer can optionally be supplied to perform special comparisons
' between elements in the array.
'
' @param arr The array to search for a specific value.
' @param value The value to search for in the array.
' @param startindex The starting index in the array to begin searching.
' @param length The number of elements to search, starting at startindex.
' @param comparer A user supplied object to compare elements within the array.
' @return The index at which the value was found.
' @remarks If the returned value is less than the lower bound of the array, then
' the value indicates where the value would have been found in the array. The
' following is how to convert the return value to the array index:
' Lowerbound of 0:  return = Not return
' Other lowerbound: return = (Not (return - lowerbound)) + lowerbound
'
Public Function BinarySearch(ByRef arr As Variant, ByRef value As Variant, Optional ByVal vStartIndex As Variant, Optional ByVal vLength As Variant, Optional ByVal comparer As IComparer) As Long
    Dim md As Long
    Dim lo As Long
    Dim hi As Long
    Dim lb As Long
    Dim pSA As Long
    Dim StartIndex As Long
    Dim Length As Long
    
    pSA = GetValidArrayPointer(arr)
    If SafeArrayGetDim(pSA) > 1 Then Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    
    lb = LBound(arr)
    If IsMissing(vStartIndex) Then StartIndex = lb Else StartIndex = vStartIndex
    If IsMissing(vLength) Then Length = cArray.GetLength(arr) Else Length = vLength
    
    hi = StartIndex + Length - 1
    lo = StartIndex

    ' do some argument checking
    If StartIndex < lb Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "StartIndex", StartIndex)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Length"), "Length", Length)
    If hi > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Length")
        
    On Error GoTo errTrap
    If comparer Is Nothing Then
        If TrySZBinarySearch(pSA, value, StartIndex, Length, BinarySearch) Then Exit Function
        Set comparer = VBCorLib.comparer.Default
    End If
    
    ' perform a search using comparer.
    Do While lo <= hi
        md = (lo + hi) \ 2
        Select Case comparer.Compare(arr(md), value)
            Case 0
                BinarySearch = md
                Exit Function
            Case Is > 0
                hi = md - 1
            Case Else
                lo = md + 1
        End Select
    Loop
    
    BinarySearch = (Not (lo - lb)) + lb
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
'   Clears a portion of the elements in an array.
'
' @param arr The array to clear elements from.
' @param startindex The starting element to being clearing.
' @param length The number of elements to be cleared.
'
Public Sub Clear(ByRef arr As Variant, ByVal StartIndex As Long, ByVal Length As Long)
    Dim pSA As Long
    Dim lb As Long
    
    pSA = GetValidArrayPointer(arr)
    lb = LBound(arr)
    If StartIndex < lb Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "startindex", StartIndex)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Length"), "Length", Length)
    If StartIndex + Length - 1 > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset))
    If SafeArrayGetDim(pSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    
    If Length = 0 Then Exit Sub
    
    Dim SA As GuidSafeArray1d
    Select Case VariantType(arr) And &HFF
        Case vbObject, vbUserDefinedType
            CopyMemory SA, ByVal pSA - SIZEOF_GUID, SIZEOF_GUIDSAFEARRAY1D
        Case Else
            CopyMemory SA.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
    End Select
    
    With SA.SA
        .fFeatures = .fFeatures Or FADF_FIXEDSIZE Or FADF_STATIC Or FADF_AUTO
        .pvData = .pvData + (StartIndex - lb) * .cbElements
        .cElements = Length
    End With
    SafeArrayDestroyData VarPtr(SA.SA)
End Sub

''
' Copies one Array to another Array and performs type casting as necessary.
'
' @param sourcearray The array from which to copy the elements.
' @param destinationarray The array in which to place the elements.
' @param length The number of elements to copy.
Public Sub Copy(ByRef sourceArray As Variant, ByRef destinationarray As Variant, ByVal Length As Long)
    CopyEx sourceArray, LBound(sourceArray), destinationarray, LBound(destinationarray), Length
End Sub

''
' Copies a section of one Array to another Array and performs type casting as necessary.
'
' @param sourcearray The array from which to copy the elements.
' @param sourceindex The starting element in the source array to begin copying from.
' @param destinationarray The array in which to place the elements.
' @param destinationindex The starting element in the destination array to place elements.
' @param length The number of elements to copy.
Public Sub CopyEx(ByRef sourceArray As Variant, ByVal sourceindex As Long, ByRef destinationarray, ByVal destinationindex As Long, ByVal Length As Long)
    Dim srcpSA As Long
    Dim dstpSA As Long
    Dim srcType As VbVarType
    Dim dstType As VbVarType
    
    srcpSA = GetValidArrayPointer(sourceArray)
    dstpSA = GetValidArrayPointer(destinationarray)
    
    If sourceindex < LBound(sourceArray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "sourceindex", sourceindex)
    If destinationindex < LBound(destinationarray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "destinationindex", destinationindex)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Length"), "Length", Length)
    If (sourceindex + Length - 1 > UBound(sourceArray)) Or (destinationindex + Length - 1 > UBound(destinationarray)) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "length")
    
    If Length = 0 Then Exit Sub
    
    srcType = VarType(sourceArray) And &HFF
    dstType = VarType(destinationarray) And &HFF
    
    On Error GoTo errTrap
    If srcType = dstType Then
        Select Case srcType
            Case vbObject, vbDataObject, vbUserDefinedType
                CopyReferenceTypes srcpSA, sourceindex, dstpSA, destinationindex, Length
            Case Else
                CopyValueTypes srcpSA, sourceindex, dstpSA, destinationindex, Length
        End Select
    Else
        Dim pfn As Long
        pfn = GetAssigningFunction(srcType, dstType)
        If pfn = 0 Then
            On Error GoTo 0
            Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
        End If
        mAssignDelegator.pfn = pfn
        CopyWithConversion srcpSA, sourceindex, dstpSA, destinationindex, Length
    End If
    Exit Sub
    
errTrap:
    If Err.Number = 13 Then
        Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Conversion))
    Else
        Throw Cor.NewSystemException(Err.Description)
    End If
End Sub

''
' Initializes a new array of the specified type in up to 3 dimensions.
'
' @param ArrayType The data type the array represents.
' @param length1 The number of elements in the first dimension.
' @param length2 The number of elements in the second dimension.
' @param length3 The number of elements in the third dimension.
' @returns An array of the specified type with the specified dimensions.
' @remarks If a length is zero, then an empty array is returned, not a Null array.
' If one length is defined as zero, then all other defined lengths must be zero.
' This does not mean that lenghts 2 and 3 need to be set to zero if only an empty
' one-dimensional array is being created.
Public Function CreateInstance(ByVal arraytype As ciArrayTypes, Optional ByVal length1 As Long = 0, Optional ByRef length2 As Variant, Optional ByRef length3 As Variant) As Variant
    Dim lengths(2) As Long
    Dim bounds() As SafeArrayBound
    Dim cDims As Long
    
    cDims = 1
    If length1 < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "length1"), "length1", length1)
    lengths(0) = length1
    
    If Not IsMissing(length2) Then
        cDims = 2
        lengths(1) = length2
        If lengths(1) < 0 Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "length2"), "length2", length2)
        If lengths(1) = 0 And lengths(0) <> 0 Then _
            Throw Cor.NewArgumentException("If defining one dimension as 0, then all other defined dimensions must be 0.", "length2")
    End If
    If Not IsMissing(length3) Then
        If cDims <> 2 Then Throw Cor.NewArgumentException("Cannot specifiy dimension without dimension 2.", "length3")
        cDims = 3
        lengths(2) = length3
        If lengths(2) < 0 Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "length3"), "length3", length3)
        If lengths(2) = 0 And lengths(1) <> 0 Then _
            Throw Cor.NewArgumentException("If defining one dimension as 0, then all other defined dimensions must be 0.", "length2")
    End If
    
    ReDim bounds(0 To cDims - 1)
    Dim ret As Variant
    Dim i As Long
    For i = 0 To cDims - 1
        bounds(i).cElements = lengths(i)
        bounds(i).lLbound = 0
    Next i
    
    VariantType(ret) = arraytype Or vbArray
    MemLong(VarPtr(ret) + VARIANTDATA_OFFSET) = SafeArrayCreate(arraytype, cDims, bounds(0))
    CreateInstance = ret
End Function

''
' Creates an array of values.
'
' @param arraytype The target datatype of the resultant array and the values to be converted to.
' @param arrayelements The values to be put into the new array.
' @return An array of the specified datatype and containing the array elements with conversion.
Public Function NewArray(ByVal arraytype As ciArrayTypes, ParamArray arrayelements() As Variant) As Variant
    Dim ub As Long
    Dim ret As Variant
    Dim args() As Variant
    
    ub = UBound(arrayelements)
    If ub >= 0 Then
        VariantType(ret) = arraytype Or vbArray
        MemLong(VarPtr(ret) + VARIANTDATA_OFFSET) = SafeArrayCreateVector(arraytype, 0, ub + 1)
        Helper.Swap4 ByVal ArrPtr(args), ByVal Helper.DerefEBP(16)
        CopyEx args, 0, ret, 0, ub + 1
    End If
    NewArray = ret
End Function

''
' Returns the number of elements in the specified dimension. If 0 is
' specified, this returns the total number of elements in the array.
'
' @param arr The array to retrieve the number of elements from.
' @param dimension (Optional) the specific dimension to get the number of elements.
' @return The number of elements in the specified dimension or the entire array.
Public Function GetLength(ByRef arr As Variant, Optional ByVal dimension As Long = 0) As Long
    Dim pSA As Long
    Dim rank As Long
    Dim i As Long
    
    pSA = GetValidArrayPointer(arr)
    rank = MemWord(pSA)
    If dimension > rank Then _
        Throw Cor.NewIndexOutOfRangeException(Environment.GetResourceString(IndexOutOfRange_Dimension))
    
    pSA = pSA + SIZEOF_SAFEARRAY
    If dimension < 1 Then
        ' optimize for 99% of calls
        If rank = 1 Then
            GetLength = MemLong(pSA)
            Exit Function
        End If
        GetLength = 1
        For i = pSA To pSA + (SIZEOF_SAFEARRAYBOUND * (rank - 1)) Step SIZEOF_SAFEARRAYBOUND
            GetLength = GetLength * MemLong(i)
        Next i
    Else
        pSA = pSA + (rank - dimension) * SIZEOF_SAFEARRAYBOUND
        GetLength = MemLong(pSA)
    End If
End Function

''
' Returns the number of dimensions in the array.
'
' @param arr The array to retrieve the number of dimensions from.
' @return The number of dimension in the array.
Public Function GetRank(ByRef arr As Variant) As Long
    GetRank = MemWord(GetValidArrayPointer(arr))
End Function

''
' Returns the index of the first occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param arr The array to search.
' @param value The value to search for in the array.
' @param startindex The index to the element in the array to begin search from.
' @param count The number of elements to search, starting from startindex.
' @param comparer An optional comparer, primarily used for User-Defined types.
' @return Value indicating the index the value was found. If a value of less that
' the lower-bounds is returned, then the value was not found in the array.
Public Function IndexOf(ByRef arr As Variant, ByRef value As Variant, Optional ByVal vStartIndex As Variant, Optional ByVal vCount As Variant, Optional ByVal comparer As IComparer) As Long
    Dim pSA As Long
    Dim StartIndex As Long
    Dim Count As Long
    
    pSA = GetValidArrayPointer(arr)
    If IsMissing(vStartIndex) Then StartIndex = LBound(arr) Else StartIndex = vStartIndex
    If IsMissing(vCount) Then Count = cArray.GetLength(arr, 1) Else Count = vCount

    If StartIndex < LBound(arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "startindex", StartIndex)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Count"), "Count", Count)
    If StartIndex + Count - 1 > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    If SafeArrayGetDim(pSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))

    On Error GoTo errTrap
    If comparer Is Nothing Then
        ' if no comparer was supplied, then try some optimized routines.
        If TrySZIndexOf(pSA, value, StartIndex, Count, IndexOf) Then Exit Function
        Set comparer = VBCorLib.comparer.Default
    End If
    ' perform a search using comparer.
    Do While Count > 0
        If comparer.Compare(arr(StartIndex), value) = 0 Then
            IndexOf = StartIndex
            Exit Function
        End If
        StartIndex = StartIndex + 1
        Count = Count - 1
    Loop
    IndexOf = LBound(arr) - 1
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Returns if an array variable is uninitialized.
'
' @param arr The array to test for initialization.
' @return Value indicating if the array is initialized.
Public Function IsNull(ByRef arr As Variant) As Boolean
    IsNull = (GetArrayPointer(arr) = 0)
End Function

''
' Returns the index of the last occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param arr The array to search.
' @param value The value to search for in the array.
' @param startindex The index to the element in the array to begin search from.
' @param count The number of elements to search, starting from startindex.
' @param comparer An optional comparer, primarily used for User-Defined types.
' @return Value indicating the index the value was found. If a value of less that
' the lower-bounds is returned, then the value was not found in the array.
Public Function LastIndexOf(ByRef arr As Variant, ByRef value As Variant, Optional ByVal vStartIndex As Variant, Optional ByVal vCount As Variant, Optional ByVal comparer As IComparer) As Long
    Dim pSA As Long
    Dim StartIndex As Long
    Dim Count As Long
    
    pSA = GetValidArrayPointer(arr)
    If IsMissing(vStartIndex) Then StartIndex = UBound(arr) Else StartIndex = vStartIndex
    If IsMissing(vCount) Then Count = cArray.GetLength(arr) Else Count = vCount
    
    If StartIndex > UBound(arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "startindex", StartIndex)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Count"), "Count", Count)
    If StartIndex - Count + 1 < LBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    If SafeArrayGetDim(pSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))

    On Error GoTo errTrap
    If comparer Is Nothing Then
        ' if no comparer was supplied, then try some optimized routines.
        If TrySZLastIndexOf(pSA, value, StartIndex, Count, LastIndexOf) Then Exit Function
        Set comparer = VBCorLib.comparer.Default
    End If
    ' perform a search using comparer.
    Do While Count > 0
        If comparer.Compare(arr(StartIndex), value) = 0 Then
            LastIndexOf = StartIndex
            Exit Function
        End If
        StartIndex = StartIndex - 1
        Count = Count - 1
    Loop
    LastIndexOf = LBound(arr) - 1
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Reverses the elements in a subportion of an array.
'
' @param arr The array to reverse.
' @param index The starting element to begin reversing elements.
' @param length The number of elements to reverse in the array.
'
Public Sub Reverse(ByRef arr As Variant, Optional ByVal vStartIndex As Variant, Optional ByVal vLength As Variant)
    Dim pSA As Long
    Dim items As SortItems
    Dim i As Long
    Dim j As Long
    Dim swapper As ISwap
    Dim StartIndex As Long
    Dim Length As Long
    
    pSA = GetValidArrayPointer(arr)
    
    If IsMissing(vStartIndex) Then StartIndex = LBound(arr) Else StartIndex = vStartIndex
    If IsMissing(vLength) Then Length = cArray.GetLength(arr) Else Length = vLength
    
    If StartIndex < LBound(arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "startindex", StartIndex)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Length"), "Length", Length)
    If StartIndex + Length - 1 > UBound(arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "length")
    If SafeArrayGetDim(pSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    
    CopyMemory items.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
    With items.SA
        Select Case .cbElements
            Case 1, 2, 4, 8, 16
                Set swapper = GetSwapper(.cbElements)
            Case Else
                items.Buffer = CoTaskMemAlloc(.cbElements)
        End Select
        
        i = .pvData + StartIndex * .cbElements
        j = i + Length * .cbElements - .cbElements
        Do While i < j
            If Not swapper Is Nothing Then
                swapper.Call 0, ByVal i, ByVal j    ' the first argument is just a filler.
            Else
                CopyMemory ByVal items.Buffer, ByVal i, .cbElements
                CopyMemory ByVal i, ByVal j, .cbElements
                CopyMemory ByVal j, ByVal items.Buffer, .cbElements
            End If
            i = i + .cbElements
            j = j - .cbElements
        Loop
    End With
    If items.Buffer Then CoTaskMemFree items.Buffer
End Sub

''
' Sorts an entire array. An optionally supplied comparer object can be
' used to compare special elements, such as userdefined values.
'
' @param arr The array to search for a value.
' @param comparer (Optional) An IComparer object to perform custom comparisons.
'
Public Sub Sort(ByRef arr As Variant, Optional ByVal comparer As IComparer)
    InternalSort arr, Empty, LBound(arr), UBound(arr) - LBound(arr) + 1, comparer
End Sub

''
' Sorts an array, or subportion, given a startindex and length. An optionally
' supplied comparer object can be used to compare special elements, such as
' userdefined values.
'
' @param arr The array to search for a value.
' @param startindex The starting index in the array to start sorting.
' @param length The number of elements to be sorted.
' @param comparer (Optional) An IComparer object to perform custom comparisons.
'
Public Sub SortEx(ByRef arr As Variant, ByVal StartIndex As Long, ByVal Length As Long, Optional ByVal comparer As IComparer)
    InternalSort arr, Empty, StartIndex, Length, comparer
End Sub

''
' Sorts an entire array based on an array of keys. An optionally supplied
' comparer object can be used to compare special elements, such as userdefined values.
'
' @param keys An array the sorting is based on.
' @param items An array that is sorted based on the sorting of keys.
' @param comparer (Optional) An IComparer object to perform custom comparisons.
'
Public Sub SortKey(ByRef Keys As Variant, ByRef items As Variant, Optional ByVal comparer As IComparer)
    InternalSort Keys, items, LBound(Keys), UBound(Keys) - LBound(Keys) + 1, comparer
End Sub

''
' Sorts an array based on an array of keys. A subportion of the array can
' be sorted given a startindex and length. An optionally, supplied comparer
' object can be used to compare special elements, such as userdefined values.
'
' @param keys An array the sorting is based on.
' @param items An array that is sorted based on the sorting of keys.
' @param startindex The starting index in the array to start sorting.
' @param length The number of elements to be sorted.
' @param comparer (Optional) An IComparer object to perform custom comparisons.
'
Public Sub SortKeyEx(ByRef Keys As Variant, ByRef items As Variant, ByVal StartIndex As Long, ByVal Length As Long, Optional ByVal comparer As IComparer)
    InternalSort Keys, items, StartIndex, Length, comparer
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBinarySearchComparer = InitDelegator(mBinarySearchDelegator)
    Set mIndexOfComparer = InitDelegator(mIndexOfDelegator)
    Set mAssigner = InitDelegator(mAssignDelegator)
    Set mSorter = InitDelegator(mSortDelegator)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CopyValueTypes(ByVal srcpSA As Long, ByVal sourceindex As Long, ByVal dstpSA As Long, ByVal destinationindex As Long, ByVal Length As Long)
    Dim srcSA As SafeArray1d
    Dim dstSA As SafeArray1d
    
    CopyMemory srcSA, ByVal srcpSA, SIZEOF_SAFEARRAY1D
    SetCopyArrayDescriptor srcSA, sourceindex, Length
    CopyMemory dstSA, ByVal dstpSA, SIZEOF_SAFEARRAY1D
    SetCopyArrayDescriptor dstSA, destinationindex, Length
    
    SafeArrayCopyData VarPtr(srcSA), VarPtr(dstSA)
End Sub

' This differs from 'CopyValueTypes' in that this includes the 16 byte
' guid that precedes the array descriptor. This allows for user-defined
' types to be copied correctly.
Private Sub CopyReferenceTypes(ByVal srcpSA As Long, ByVal sourceindex As Long, ByVal dstpSA As Long, ByVal destinationindex As Long, ByVal Length As Long)
    Dim srcSA As GuidSafeArray1d
    Dim dstSA As GuidSafeArray1d
    
    CopyMemory srcSA, ByVal srcpSA - 16, LenB(srcSA)
    SetCopyArrayDescriptor srcSA.SA, sourceindex, Length
    CopyMemory dstSA, ByVal dstpSA - 16, LenB(dstSA)
    SetCopyArrayDescriptor dstSA.SA, destinationindex, Length
    
    SafeArrayCopyData VarPtr(srcSA.SA), VarPtr(dstSA.SA)
End Sub

Private Sub CopyWithConversion(ByVal srcpSA As Long, ByVal sourceindex As Long, ByVal dstpSA As Long, ByVal destinationindex As Long, ByVal Length As Long)
    Dim src As Long
    Dim dst As Long
    Dim srcLb As Long
    Dim dstLb As Long
    Dim srcSize As Long
    Dim dstSize As Long
    
    srcSize = SafeArrayGetElemsize(srcpSA)
    dstSize = SafeArrayGetElemsize(dstpSA)
    srcLb = SafeArrayGetLBound(srcpSA, 1)
    dstLb = SafeArrayGetLBound(dstpSA, 1)
    
    src = MemLong(srcpSA + SAFEARRAYDATAPOINTER_OFFSET) + (sourceindex - srcLb) * srcSize
    dst = MemLong(dstpSA + SAFEARRAYDATAPOINTER_OFFSET) + (destinationindex - dstLb) * dstSize
    
    Do While Length > 0
        mAssigner.Call dst, src
        src = src + srcSize
        dst = dst + dstSize
        Length = Length - 1
    Loop
End Sub

Private Function GetAssigningFunction(ByVal srcType As VbVarType, ByVal dstType As VbVarType) As Long
    Dim ret As Long
    Select Case srcType
        Case vbLong
            Select Case dstType
                Case vbDouble:      ret = FuncAddr(AddressOf WidenLongToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenLongToString)
                Case vbCurrency:    ret = FuncAddr(AddressOf WidenLongToCurrency)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenLongToVariant)
            End Select
        Case vbInteger, vbBoolean
            Select Case dstType
                Case vbLong:        ret = FuncAddr(AddressOf WidenIntegerToLong)
                Case vbDouble:      ret = FuncAddr(AddressOf WidenIntegerToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenIntegerToString)
                Case vbSingle:      ret = FuncAddr(AddressOf WidenIntegerToSingle)
                Case vbCurrency:    ret = FuncAddr(AddressOf WidenIntegerToCurrency)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenIntegerToVariant)
            End Select
        Case vbByte
            Select Case dstType
                Case vbLong:        ret = FuncAddr(AddressOf WidenByteToLong)
                Case vbInteger:     ret = FuncAddr(AddressOf WidenByteToInteger)
                Case vbDouble:      ret = FuncAddr(AddressOf WidenByteToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenByteToString)
                Case vbSingle:      ret = FuncAddr(AddressOf WidenByteToSingle)
                Case vbCurrency:    ret = FuncAddr(AddressOf WidenByteToCurrency)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenByteToVariant)
            End Select
        Case vbSingle
            Select Case dstType
                Case vbDouble:      ret = FuncAddr(AddressOf WidenSingleToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenSingleToString)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenSingleToVariant)
            End Select
        Case vbDate
            Select Case dstType
                Case vbDouble:      ret = FuncAddr(AddressOf WidenDateToDouble)
                Case vbString:      ret = FuncAddr(AddressOf WidenDateToString)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenDateToVariant)
            End Select
        Case vbObject, vbDataObject
            Select Case dstType
                Case vbVariant:     ret = FuncAddr(AddressOf WidenObjectToVariant)
            End Select
        Case vbCurrency
            Select Case dstType
                Case vbString:      ret = FuncAddr(AddressOf WidenCurrencyToString)
                Case vbVariant:     ret = FuncAddr(AddressOf WidenCurrencyToVariant)
            End Select
        Case vbString
            Select Case dstType
                Case vbVariant:     ret = FuncAddr(AddressOf WidenStringToVariant)
            End Select
        Case vbVariant
            Select Case dstType
                Case vbLong:                    ret = FuncAddr(AddressOf NarrowVariantToLong)
                Case vbInteger:                 ret = FuncAddr(AddressOf NarrowVariantToInteger)
                Case vbString:                  ret = FuncAddr(AddressOf NarrowVariantToString)
                Case vbDouble:                  ret = FuncAddr(AddressOf NarrowVariantToDouble)
                Case vbObject, vbDataObject:    ret = FuncAddr(AddressOf NarrowVariantToObject)
                Case vbSingle:                  ret = FuncAddr(AddressOf NarrowVariantToSingle)
                Case vbByte:                    ret = FuncAddr(AddressOf NarrowVariantToByte)
                Case vbDate:                    ret = FuncAddr(AddressOf NarrowVariantToDate)
                Case vbBoolean:                 ret = FuncAddr(AddressOf NarrowVariantToBoolean)
                Case vbCurrency:                ret = FuncAddr(AddressOf NarrowVariantToCurrency)
            End Select
    End Select
    GetAssigningFunction = ret
End Function

''
' Perform the sorting. If no comparer is supplied, then optimized sorting routines
' will be attempted, otherwise, the general routine will be used with the comparer.
Private Sub InternalSort(ByRef Keys As Variant, ByRef items As Variant, ByVal Index As Long, ByVal Length As Long, ByVal comparer As IComparer)
    Dim pKeysSA As Long
    Dim pItemsSA As Long
    
    pKeysSA = GetValidArrayPointer(Keys)
    If SafeArrayGetDim(pKeysSA) > 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
    If Index < LBound(Keys) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "index", Index)
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum, "Length"), "Length", Length)
    If Index + Length - 1 > UBound(Keys) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "length")
    
    If Not IsEmpty(items) Then
        pItemsSA = GetValidArrayPointer(items)
        If SafeArrayGetDim(pItemsSA) > 1 Then _
            Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimension))
        If (LBound(Keys) <> LBound(items)) Or _
           (UBound(Keys) <> UBound(items)) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_MatchingBounds))
        
        SetSortItems pItemsSA
    End If
    
    On Error GoTo errTrap
    If comparer Is Nothing Then
        If TrySZSort(pKeysSA, Index, Index + Length - 1) Then
            ClearSortKeys
            ClearSortItems
            Exit Sub
        End If
        Set comparer = VBCorLib.comparer.Default
    End If
    SetSortKeys pKeysSA
    Set SortComparer = comparer
    QuickSortGeneral Keys, Index, Index + Length - 1
    Set SortComparer = Nothing
    ClearSortKeys
    ClearSortItems
    Exit Sub
    
errTrap:
    ClearSortKeys
    ClearSortItems
    Set SortComparer = Nothing
    
    Dim ex As Exception
    Catch ex, Err
    If ex.HResult = 13 Then
        Throw Cor.NewInvalidCastException(Environment.GetResourceString(ArrayTypeMismatch_Compare))
    Else
        Throw ex
    End If
End Sub

Private Sub SetCopyArrayDescriptor(ByRef desc As SafeArray1d, ByVal Index As Long, ByVal Length As Long)
    With desc
        .pvData = .pvData + (Index - .lLbound) * .cbElements
        .cDims = 1
        .cElements = Length
    End With
End Sub

''
' This is an optimized search routine that uses a function pointer
' to call a specific comparison routine.
Private Function SZBinarySearch(ByVal pSA As Long, ByVal pValue As Long, ByVal StartIndex As Long, ByVal Length As Long, ByVal pfn As Long) As Long
    Dim size As Long
    Dim pvData As Long
    Dim lo As Long
    Dim hi As Long
    Dim md As Long
    Dim lb As Long
    
    mBinarySearchDelegator.pfn = pfn
    size = SafeArrayGetElemsize(pSA)
    pvData = MemLong(pSA + 12)
    
    lb = SafeArrayGetLBound(pSA, 1)
    lo = StartIndex - lb
    hi = lo + Length - 1
    
    Do While lo <= hi
        md = (lo + hi) \ 2
        Select Case mBinarySearchComparer.Call(pvData + md * size, pValue)
            Case 0
                SZBinarySearch = md + lb
                Exit Function
            Case Is > 0
                hi = md - 1
            Case Else
                lo = md + 1
        End Select
    Loop
    
    SZBinarySearch = (Not lo) + lb
End Function

Private Function SZIndexOf(ByVal pSA As Long, ByVal pValue As Long, ByVal StartIndex As Long, ByVal Count As Long, ByVal pfn As Long) As Long
    Dim size As Long
    Dim pvData As Long
    Dim lb As Long
    
    mIndexOfDelegator.pfn = pfn
    size = SafeArrayGetElemsize(pSA)
    pvData = MemLong(pSA + 12)
    lb = SafeArrayGetLBound(pSA, 1)
    
    StartIndex = (StartIndex - lb)
    Do While Count > 0
        If mIndexOfComparer.Call(pvData + StartIndex * size, pValue) Then
            SZIndexOf = StartIndex + lb
            Exit Function
        End If
        Count = Count - 1
        StartIndex = StartIndex + 1
    Loop

    SZIndexOf = lb - 1
End Function

Private Function SZLastIndexOf(ByVal pSA As Long, ByVal pValue As Long, ByVal StartIndex As Long, ByVal Count As Long, ByVal pfn As Long) As Long
    Dim size As Long
    Dim pvData As Long
    Dim lb As Long
    
    mIndexOfDelegator.pfn = pfn
    size = SafeArrayGetElemsize(pSA)
    pvData = MemLong(pSA + 12)
    lb = SafeArrayGetLBound(pSA, 1)
    
    StartIndex = (StartIndex - lb)
    Do While Count > 0
        If mIndexOfComparer.Call(pvData + StartIndex * size, pValue) Then
            SZLastIndexOf = StartIndex + lb
            Exit Function
        End If
        Count = Count - 1
        StartIndex = StartIndex - 1
    Loop

    SZLastIndexOf = lb - 1
End Function

''
' Attempt to use a specialized search for a specific data type.
Private Function TrySZBinarySearch(ByVal pSA As Long, ByRef value As Variant, ByVal StartIndex As Long, ByVal Length As Long, ByRef retval As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA)
        Case vbLong:                    retval = SZBinarySearch(pSA, VarPtr(CLng(value)), StartIndex, Length, AddressOf CompareLongs)
        Case vbString:                  retval = SZBinarySearch(pSA, VarPtr(StrPtr(value)), StartIndex, Length, AddressOf CompareStrings)
        Case vbDouble:                  retval = SZBinarySearch(pSA, VarPtr(CDbl(value)), StartIndex, Length, AddressOf CompareDoubles)
        Case vbObject, vbDataObject:    retval = SZBinarySearch(pSA, VarPtr(value), StartIndex, Length, AddressOf CompareIComparable)
        Case vbVariant:                 retval = SZBinarySearch(pSA, VarPtr(value), StartIndex, Length, AddressOf CompareVariants)
        Case vbInteger:                 retval = SZBinarySearch(pSA, VarPtr(CInt(value)), StartIndex, Length, AddressOf CompareIntegers)
        Case vbSingle:                  retval = SZBinarySearch(pSA, VarPtr(CSng(value)), StartIndex, Length, AddressOf CompareSingles)
        Case vbCurrency:                retval = SZBinarySearch(pSA, VarPtr(CCur(value)), StartIndex, Length, AddressOf CompareCurrencies)
        Case vbDate:                    retval = SZBinarySearch(pSA, VarPtr(CDate(value)), StartIndex, Length, AddressOf CompareDates)
        Case vbBoolean:                 retval = SZBinarySearch(pSA, VarPtr(CBool(value)), StartIndex, Length, AddressOf CompareBooleans)
        Case vbByte:                    retval = SZBinarySearch(pSA, VarPtr(CByte(value)), StartIndex, Length, AddressOf CompareBytes)
        Case Else: Exit Function
    End Select
    TrySZBinarySearch = True
End Function

Private Function TrySZIndexOf(ByVal pSA As Long, ByRef value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef retval As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    retval = SZIndexOf(pSA, VarPtr(CLng(value)), StartIndex, Count, AddressOf EqualsLong)
        Case vbString:                  retval = SZIndexOf(pSA, VarPtr(StrPtr(value)), StartIndex, Count, AddressOf EqualsString)
        Case vbDouble:                  retval = SZIndexOf(pSA, VarPtr(CDbl(value)), StartIndex, Count, AddressOf EqualsDouble)
        Case vbDate:                    retval = SZIndexOf(pSA, VarPtr(CDate(value)), StartIndex, Count, AddressOf EqualsDate)
        Case vbObject, vbDataObject:    retval = SZIndexOf(pSA, VarPtr(ObjPtr(value)), StartIndex, Count, AddressOf EqualsObject)
        Case vbVariant:                 retval = SZIndexOf(pSA, VarPtr(value), StartIndex, Count, AddressOf EqualsVariants)
        Case vbInteger:                 retval = SZIndexOf(pSA, VarPtr(CInt(value)), StartIndex, Count, AddressOf EqualsInteger)
        Case vbSingle:                  retval = SZIndexOf(pSA, VarPtr(CSng(value)), StartIndex, Count, AddressOf EqualsSingle)
        Case vbByte:                    retval = SZIndexOf(pSA, VarPtr(CSng(value)), StartIndex, Count, AddressOf EqualsByte)
        Case vbBoolean:                 retval = SZIndexOf(pSA, VarPtr(CBool(value)), StartIndex, Count, AddressOf EqualsBoolean)
        Case vbCurrency:                retval = SZIndexOf(pSA, VarPtr(CCur(value)), StartIndex, Count, AddressOf EqualsCurrency)
        Case Else: Exit Function
    End Select
    TrySZIndexOf = True
End Function

Private Function TrySZLastIndexOf(ByVal pSA As Long, ByRef value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef retval As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    retval = SZLastIndexOf(pSA, VarPtr(CLng(value)), StartIndex, Count, AddressOf EqualsLong)
        Case vbString:                  retval = SZLastIndexOf(pSA, VarPtr(StrPtr(value)), StartIndex, Count, AddressOf EqualsString)
        Case vbDouble:                  retval = SZLastIndexOf(pSA, VarPtr(CDbl(value)), StartIndex, Count, AddressOf EqualsDouble)
        Case vbDate:                    retval = SZLastIndexOf(pSA, VarPtr(CDate(value)), StartIndex, Count, AddressOf EqualsDate)
        Case vbObject, vbDataObject:    retval = SZLastIndexOf(pSA, VarPtr(ObjPtr(value)), StartIndex, Count, AddressOf EqualsObject)
        Case vbVariant:                 retval = SZLastIndexOf(pSA, VarPtr(value), StartIndex, Count, AddressOf EqualsVariants)
        Case vbInteger:                 retval = SZLastIndexOf(pSA, VarPtr(CInt(value)), StartIndex, Count, AddressOf EqualsInteger)
        Case vbSingle:                  retval = SZLastIndexOf(pSA, VarPtr(CSng(value)), StartIndex, Count, AddressOf EqualsSingle)
        Case vbByte:                    retval = SZLastIndexOf(pSA, VarPtr(CSng(value)), StartIndex, Count, AddressOf EqualsByte)
        Case vbBoolean:                 retval = SZLastIndexOf(pSA, VarPtr(CBool(value)), StartIndex, Count, AddressOf EqualsBoolean)
        Case vbCurrency:                retval = SZLastIndexOf(pSA, VarPtr(CCur(value)), StartIndex, Count, AddressOf EqualsCurrency)
        Case Else: Exit Function
    End Select
    TrySZLastIndexOf = True
End Function

''
'   Attempt to locate an optimized sort routine for a specific data type.
Private Function TrySZSort(ByVal pSA As Long, ByVal left As Long, ByVal right As Long) As Boolean
    Dim pfn As Long
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    pfn = FuncAddr(AddressOf QuickSortLong)
        Case vbString:                  pfn = FuncAddr(AddressOf QuickSortString)
        Case vbDouble, vbDate:          pfn = FuncAddr(AddressOf QuickSortDouble)
        Case vbObject, vbDataObject:    pfn = FuncAddr(AddressOf QuickSortObject)
        Case vbVariant:                 pfn = FuncAddr(AddressOf QuickSortVariant)
        Case vbInteger:                 pfn = FuncAddr(AddressOf QuickSortInteger)
        Case vbSingle:                  pfn = FuncAddr(AddressOf QuickSortSingle)
        Case vbByte:                    pfn = FuncAddr(AddressOf QuickSortByte)
        Case vbCurrency:                pfn = FuncAddr(AddressOf QuickSortCurrency)
        Case vbBoolean:                 pfn = FuncAddr(AddressOf QuickSortBoolean)
        Case Else: Exit Function
    End Select
    mSortDelegator.pfn = pfn
    mSorter.Call pSA, left, right
    TrySZSort = True
End Function

''
' This will verify that the array pointer is not Null. If it is null
' then an ArgumentNullException will be thrown.
Friend Function GetValidArrayPointer(ByRef arr As Variant) As Long
    GetValidArrayPointer = GetArrayPointer(arr)
    If GetValidArrayPointer = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "arr")
End Function

Private Function GetArrayPointer(ByRef arr As Variant) As Long
    Const BYREF_ARRAY As Long = VT_BYREF Or vbArray
    
    Dim vt As Long
    
    vt = VariantType(arr)
    Select Case vt And BYREF_ARRAY
        ' we have to double deref the original array pointer because
        ' the variant held a pointer to the original array variable.
        Case BYREF_ARRAY:   GetArrayPointer = MemLong(MemLong(VarPtr(arr) + VARIANTDATA_OFFSET))
        ' we won't need to deref again if the original array was dimensioned
        ' as a variant ie:
        '    Dim arr As Variant
        '    ReDim arr(1 To 10) As Long
        '
        ' The passed in variant will be the array variable, not a ByRef
        ' pointer to the array variable.
        Case vbArray:       GetArrayPointer = MemLong(VarPtr(arr) + VARIANTDATA_OFFSET)
        ' you bad person
        Case Else:          Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ArrayRequired), "arr")
    End Select
    
    ' HACK HACK HACK
    '
    ' When an uninitialized array of objects or UDTs is passed into a
    ' function as a ByRef Variant, the array is initialized with just the
    ' SafeArrayDescriptor, at which point, it is a valid array and can
    ' be used by UBound and LBound after the call. So, now we're just
    ' going to assume that any object or UDT array that has just the descriptor
    ' allocated was Null to begin with. That means whenever an Object or UDT
    ' array is passed to any cArray method, it will technically never
    ' be uninitialized, just zero-length.
    Select Case vt And &HFF
        Case vbObject, vbUserDefinedType: If MemLong(GetArrayPointer + PVDATA_OFFSET) = 0 Then GetArrayPointer = 0
    End Select
End Function
