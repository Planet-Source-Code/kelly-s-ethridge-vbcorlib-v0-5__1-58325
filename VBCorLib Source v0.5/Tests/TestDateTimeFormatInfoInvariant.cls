VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDateTimeFormatInfoInvariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDateTimeFormatInfoInvariant
'
Option Explicit

Private info As DateTimeFormatInfo
Private i As Long
Private s() As String

Public Sub Setup()
    Set info = New DateTimeFormatInfo
End Sub

Public Sub TestAbbreviatedDayNames()
    s = info.AbbreviatedDayNames
    Assert.Equals 7, carray.GetLength(s), "Should return 7 days."
    For i = 0 To 6
        Assert.Equals Choose(i + 1, "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"), s(i)
    Next i
End Sub

Public Sub TestAbbreviatedMonthsNames()
    s = info.AbbreviatedMonthNames
    Assert.Equals 13, carray.GetLength(s), "Should return 13 months."
    For i = 0 To 12
        Assert.Equals Choose(i + 1, "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""), s(i)
    Next i
End Sub

Public Sub TestAMDesignator()
    Assert.Equals "AM", info.AMDesignator
End Sub

Public Sub TestDateSeparator()
    Assert.Equals "/", info.DateSeparator
End Sub

Public Sub TestDayNames()
    s = info.DayNames
    Assert.Equals 7, carray.GetLength(s), "Should return 7 days."
    For i = 0 To 6
        Assert.Equals Choose(i + 1, "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"), s(i)
    Next i
End Sub

Public Sub TestFirstDayOfWeek()
    Assert.Equals DayOfWeek.Sunday, info.FirstDayOfWeek
End Sub

Public Sub TestLongDatePattern()
    Assert.Equals "dddd, dd MMMM yyyy", info.LongDatePattern
End Sub

Public Sub TestLongTimePattern()
    Assert.Equals "HH:mm:ss", info.LongTimePattern
End Sub

Public Sub TestFullDateTimePattern()
    Assert.Equals "dddd, dd MMMM yyyy HH:mm:ss", info.FullDateTimePattern
End Sub

Public Sub TestGetAbbreviatedDayName()
    For i = 0 To 6
        Assert.Equals Choose(i + 1, "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"), info.GetAbbreviatedDayName(i)
    Next i
End Sub

Public Sub TestGetAllDateTimePatterns_LowerD()
    s = info.GetAllDateTimePatterns("d")
    Assert.Equals 1, carray.GetLength(s), "Should return 1 element array."
    Assert.Equals "MM/dd/yyyy", s(0)
End Sub

Public Sub TestGetAllDateTimePatterns_UpperD()
    s = info.GetAllDateTimePatterns("D")
    Assert.Equals 1, carray.GetLength(s), "Should return 1 element array."
    Assert.Equals "dddd, dd MMMM yyyy", s(0)
End Sub

Public Sub TestGetAllDateTimePatterns_UpperF()
    s = info.GetAllDateTimePatterns("F")
    Assert.Equals 1, carray.GetLength(s), "Should return 1 element array."
    Assert.Equals "dddd, dd MMMM yyyy HH:mm:ss", s(0)
End Sub

Public Sub TestGetAllDateTimePatterns_LowerF()
    s = info.GetAllDateTimePatterns("f")
    Assert.Equals 4, carray.GetLength(s), "Should return 4 element array."
    
    Dim i As Long
    For i = 0 To UBound(s)
        Assert.IsIn s(i), "dddd, dd MMMM yyyy HH:mm", "dddd, dd MMMM yyyy hh:mm tt", "dddd, dd MMMM yyyy H:mm", "dddd, dd MMMM yyyy h:mm tt"
    Next i
End Sub

Public Sub TestGetAllPatterns_LowerG()
    Dim s() As String
    Dim i As Long
    
    s = info.GetAllDateTimePatterns("g")
    Assert.Equals 4, carray.GetLength(s), "s should contain 4 elements."
    
    For i = 0 To 3
        Assert.IsIn s(i), "MM/dd/yyyy HH:mm", "MM/dd/yyyy hh:mm tt", "MM/dd/yyyy H:mm", "MM/dd/yyyy h:mm tt"
    Next i
End Sub

Public Sub TestGetAllPatterns_UpperG()
    Dim s() As String
    
    s = info.GetAllDateTimePatterns("G")
    Assert.Equals 1, carray.GetLength(s), "s should contain 1 element."
    Assert.Equals "MM/dd/yyyy HH:mm:ss", s(0)
End Sub

Public Sub TestGetAllPatterns()
    Dim s() As String
    Dim i As Long
    
    s = info.GetAllDateTimePatterns
    Assert.Equals 26, carray.GetLength(s), "s should contain 26 elements."
End Sub

Public Sub TestFormatDayPart()
    Assert.Equals "Thursday Thu 01 1", info.Format(#1/1/2004#, "dddd ddd dd d")
End Sub

Public Sub TestFormatTime()
    Assert.Equals "05:30:15", info.Format(#5:30:15 AM#, "HH:mm:ss")
End Sub

Public Sub TestCustomFormat()
    Assert.Equals "Thursday, 01 January 02004 18:30", info.Format(#1/1/2004 6:30:00 PM#, "dddd, dd MMMM yyyyy HH:mm")
End Sub

Public Sub TestIsReadOnly()
    Assert.IsFalse info.IsReadOnly, "Locally created info should not be ReadOnly."
    Assert.IsTrue CultureInfo.CurrentCulture.DateTimeFormat.IsReadOnly, "CurrentCulture supplied should be ReadOnly."
    Assert.IsTrue CultureInfo.InvariantCulture.DateTimeFormat.IsReadOnly, "CurrentCulture supplied should be ReadOnly."
End Sub

Public Sub TestTimeZoneFormats()
    Dim s As String
    
    s = info.Format(Time, "%z")
    Assert.Equals "-8", s
    
    s = info.Format(Time, "%zz")
    Assert.Equals "-08", s
    
    s = info.Format(Time, "%zzz")
    Assert.Equals "-08:00", s
    
    s = info.Format(Time, "%zzzzzzz")
    Assert.Equals "-08:00", s
End Sub

Public Sub TestEscapeSequence()
    Assert.Equals "d", info.Format(Now, "\d")
    Assert.Equals "\", info.Format(Now, "\\")
End Sub
