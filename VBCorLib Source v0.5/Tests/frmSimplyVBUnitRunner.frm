VERSION 5.00
Object = "{AE240A3C-E73F-4E72-AD0B-CC377D3BDCFC}#1.1#0"; "SimplyVBUnit.ocx"
Begin VB.Form frmSimplyVBUnitRunner 
   Caption         =   "Simply VB Unit"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   15450
   Icon            =   "frmSimplyVBUnitRunner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   15450
   StartUpPosition =   2  'CenterScreen
   Begin SimplyVBUnit.SimplyVBUnitTestView SimplyVBUnitTestView1 
      Height          =   6495
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11456
   End
End
Attribute VB_Name = "frmSimplyVBUnitRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   frmSimplyVBUnitRunner
'
Option Explicit


' Namespaces available:
'       Assert.*            ie. Assert.IsTrue Value
'       Console.*           ie. Console.WriteLine "Hello, World"

' Adding a testcase:
'   Use AddTest <object>

' Steps to create a TestCase:
'
'   1. Add a new class
'   2. Name it as desired
'   3. (Optionally) Add a public sub named Setup if you want Setup run before each test in the class.
'   4. (Optionally) Add a public sub named Teardown if you want Teardown run after each test in the class.
'   5. Add public subs of the tests you want run. No parameters.

Private Sub Form_Load()
#Const ALLTESTS = 1
    
    ' Add test cases here.
    '
    'AddTest New MyTestCase
    
    
    
' StreamReader Determine Encoding tests
#If False Or ALLTESTS Then
    AddTest New TestStreamReaderDetermineEncoding
#End If
' BinaryReader tests
#If False Or ALLTESTS Then
    AddTest New TestBinaryReader
#End If
' BinaryWriter tests
#If False Or ALLTESTS Then
    AddTest New TestBinaryWriter
#End If
' FileInfo tests
#If False Or ALLTESTS Then
    AddTest New TestFileInfo
#End If
' File tests
#If False Or ALLTESTS Then
    AddTest New TestFile
#End If
' StreamReader tests
#If False Or ALLTESTS Then
    AddTest New TestStreamReader
#End If
' StreamWriter tests
#If False Or ALLTESTS Then
    AddTest New TestStreamWriter
    AddTest New TestStreamWriterWithMem
    AddTest New TestStreamWriterWithMemAutoFlush
#End If
' Directory tests
#If True Or ALLTESTS Then
    AddTest New TestDirectory
#End If
' DirectoryInfo tests
#If True Or ALLTESTS Then
    AddTest New TestDirectoryInfo
#End If
' StringReader tests
#If False Or ALLTESTS Then
    AddTest New TestStringReader
#End If
' StringWriter tests
#If False Or ALLTESTS Then
    AddTest New TestStringWriter
#End If
' UnicodeEncoding with big-endian order tests
#If False Or ALLTESTS Then
    AddTest New TestUnicodeEncodingBigEndian
#End If
' UnicodeEncoding tests
#If False Or ALLTESTS Then
    AddTest New TestUnicodeEncoding
#End If
' FileStream tests
#If False Or ALLTESTS Then
    AddTest New TestFileStreamWrite
    AddTest New TestFileStreamSmallBuffer
    AddTest New TestFileStream
#End If
' MemoryStream tests
#If False Or ALLTESTS Then
    AddTest New TestUserMemoryStream
    AddTest New TestMemoryStream
#End If
' UTF7Encoding tests
#If False Or ALLTESTS Then
    AddTest New TestUTF7GetChars
    AddTest New TestUTF7GetCharCount
    AddTest New TestUTF7GetBytes
    AddTest New TestUTF7GetByteCount
#End If
' UTF8Encoding tests
#If False Or ALLTESTS Then
    AddTest New TestUTF8GetChars
    AddTest New TestUTF8GetCharCount
    AddTest New TestUTF8Encoding
    AddTest New TestUTF8GetByteCount
#End If
' Path tests
#If True Or ALLTESTS Then
    AddTest New TestPath
#End If
' This only displays outputs from the Environment class.
#If False Or ALLTESTS Then
    AddTest New TestEnvironment
#End If
' TimeZone tests
#If False Or ALLTESTS Then
    AddTest New TestTimeZone
#End If
' DateTimeFormatInfo Invariant tests
#If False Or ALLTESTS Then
    AddTest New TestDateTimeFormatInfoInvariant
#End If
' CultureInfo tests
#If False Or ALLTESTS Then
    AddTest New TestCultureInfo
#End If
' MappedFile tests
#If False Or ALLTESTS Then
    AddTest New TestMappedFile
#End If
' FileNotFoundException tests
#If False Or ALLTESTS Then
    AddTest New TestFileNotFoundException
#End If
' cDateTime tests
#If False Or ALLTESTS Then
    AddTest New TestcDateTime
#End If
' TimeSpan tests
#If False Or ALLTESTS Then
    AddTest New TestTimeSpan
    AddTest New TestTimeSpan994394150ms
    AddTest New TestTimeSpanCreation
#End If
' Version tests
#If False Or ALLTESTS Then
    AddTest New TestVersion
#End If
' Random tests
#If False Or ALLTESTS Then
    AddTest New TestRandom
#End If
' BitConverter tests
#If False Or ALLTESTS Then
    AddTest New TestBitConverter
#End If
' NumberFormatInfo tests
#If False Or ALLTESTS Then
    AddTest New TestNumberFormatInfoInt
    AddTest New TestNumberFormatInfoFlt
    AddTest New TestNumberFormatInfoSng
#End If
' WeakReference tests
#If False Or ALLTESTS Then
    AddTest New TestWeakReference
#End If
' Hashtable tests
#If False Or ALLTESTS Then
    AddTest New TestHashTable
#End If
' Buffer tests
#If False Or ALLTESTS Then
    AddTest New TestBuffer
#End If
' BitArray tests
#If False Or ALLTESTS Then
    AddTest New TestBitArray
#End If
' SortedList tests
#If False Or ALLTESTS Then
    AddTest New TestSortedList
#End If
' DictionaryEntry
#If False Or ALLTESTS Then
    AddTest New TestDictionaryEntry
#End If
' Queue tests
#If False Or ALLTESTS Then
    AddTest New TestQueue
#End If
' Stack tests
#If False Or ALLTESTS Then
    AddTest New TestStack
#End If
' ArrayList tests
#If False Or ALLTESTS Then
    AddTest New TestArrayListExceptions
    AddTest New TestArrayListRange
    AddTest New TestArrayList10Items
    AddTest New TestArrayList
#End If
' cString tests
#If False Or ALLTESTS Then
    AddTest New TestcString
#End If
' StringBuilder tests
#If False Or ALLTESTS Then
    AddTest New TestStringBuilder
#End If
' Exception tests
#If False Or ALLTESTS Then
    AddTest New TestException
    AddTest New TestDefaultException
    AddTest New TestSystemException
    AddTest New TestDefaultSystemEx
    AddTest New TestArgumentException
    AddTest New TestDefaultArgumentEx
    AddTest New TestDefaultArgumentNull
    AddTest New TestArgumentNullException
    AddTest New TestArgumentOutOfRange
    AddTest New TestDefArgumentOutOfRange
#End If

' cArray tests
#If False Or ALLTESTS Then
    AddTest New TestcArray
    AddTest New TestInvalidCastException
    AddTest New TestDefInvalidCast
    AddTest New TestExceptionMethods
    AddTest New TestDefaultComparer
    AddTest New TestPosNumBinarySearch
    AddTest New TestMixNumBinarySearch
    AddTest New TestArraySort
    AddTest New TestArrayBinarySearch
    AddTest New TestArrayReverse
    AddTest New TestArrayIndexOf
    AddTest New TestArrayLastIndexOf
    AddTest New TestArrayCopy
    AddTest New TestArrayCreation
#End If

End Sub

Private Sub Form_Resize()
    SimplyVBUnitTestView1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Private Sub Form_Initialize()
    Caption = "Simply VB Unit - " & App.Title
    Me.SimplyVBUnitTestView1.Init App.EXEName
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub


