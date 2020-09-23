VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestcDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDateTime
'
Option Explicit

Public Sub TestZeroToOADate()
    Dim dt As cDateTime
    Set dt = New cDateTime
    Assert.Equals #12:00:00 AM#, dt.ToOADate, "ToOADate should return a date of zero."
End Sub

Public Sub Test01012004ToOADate()
    Dim dt As cDateTime
    Set dt = newcdatetime(2004, 1, 1, 8, 30, 30)
    Assert.Equals #1/1/2004 8:30:30 AM#, dt.ToOADate, "ToOADate should return #1/1/2004 8:30:30am#."
End Sub

Public Sub Test01011503ToOADate()
    Dim dt As cDateTime
    Set dt = newcdatetime(1503, 1, 1, 8, 30, 30)
    Assert.Equals #1/1/1503 8:30:30 AM#, dt.ToOADate, "ToOADate should return #1/1/1503 8:30:30am#."
End Sub

Public Sub TestYear1to2010()
    Dim dt As cDateTime
    Dim i As Long
    
    For i = 1 To 2010
        Set dt = newcdatetime(i, 12, 31)
        Assert.Equals i, dt.Year, "Year should return " & i & "."
    Next i
End Sub

Public Sub TestIsLeapYear()
    Assert.IsTrue cDateTime.IsLeapYear(2000), "IsLeapYear should return true."
    Assert.IsTrue cDateTime.IsLeapYear(2004), "IsLeapYear should return true."
    Assert.IsFalse cDateTime.IsLeapYear(2100), "IsLeapYear should return false."
End Sub

Public Sub TestTimeOnly()
    Dim dt As cDateTime
    
    Set dt = newcdatetime(, , , 18, 39, 52)
    Assert.Equals #6:39:52 PM#, dt.ToOADate
End Sub

Public Sub TestDay1to31()
    Dim dt As cDateTime
    Dim i As Long
    
    For i = 1 To 31
        Set dt = newcdatetime(2004, 12, i, 13, 31, 31, 501)
        Assert.Equals i, dt.Day
        Set dt = newcdatetime(2004, 1, i)
        Assert.Equals i, dt.Day
    Next i
End Sub

Public Sub TestMonth()
    Dim dt As cDateTime
    Dim i As Long
    Dim m As Long
    
    For i = 2000 To 2009
        For m = 1 To 12
            Set dt = newcdatetime(i, m, 1, 13, 31, 31, 501)
            Assert.Equals m, dt.Month
        Next m
    Next i
End Sub

Public Sub TestDayOfYear()
    Dim dt As cDateTime
    Dim i As Long
    Dim d As Date
    
    d = #1/1/2004#
    For i = 1 To 366
        Set dt = newcdatetime(2004, Month(d), Day(d), 13, 31, 31, 501)
        Assert.Equals i, dt.DayOfYear
        d = DateAdd("d", 1, d)
    Next i
End Sub

Public Sub TestDayOfWeek()
    Dim dt As cDateTime
    Dim i As Long
    Dim d As Date
    Dim dow As Long
    
    d = #1/1/2004#
    dow = Weekday(d) - 1
    For i = 1 To 366
        Set dt = newcdatetime(2004, Month(d), Day(d), 13, 31, 31, 501)
        Assert.Equals dow, dt.DayOfWeek
        d = DateAdd("d", 1, d)
        dow = (dow + 1) Mod 7
    Next i
End Sub

Public Sub TestHour()
    Dim dt As cDateTime
    Dim i As Long
    
    For i = 0 To 23
        Set dt = newcdatetime(2004, 1, 1, i, 0, 0, 501)
        Assert.Equals i, dt.Hour
    Next i
End Sub

Public Sub TestMinute()
    Dim dt As cDateTime
    Dim i As Long
    
    For i = 0 To 59
        Set dt = newcdatetime(2004, 1, 1, 1, i, 31, 501)
        Assert.Equals i, dt.Minute
    Next i
End Sub

Public Sub TestSecond()
    Dim dt As cDateTime
    Dim i As Long
    
    For i = 0 To 59
        Set dt = newcdatetime(2004, 1, 1, 1, 1, i, 501)
        Assert.Equals i, dt.Second
    Next i
End Sub

Public Sub TestMillisecond()
    Dim dt As cDateTime
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 0)
    Assert.Equals 0, dt.Millisecond
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 999)
    Assert.Equals 999, dt.Millisecond
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 501)
    Assert.Equals 501, dt.Millisecond
End Sub

Public Sub TestTotalMilliseconds()
    Dim dt As cDateTime
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    Assert.Equals 63208515661001@, dt.TotalMilliseconds
End Sub

Public Sub TestTicks()
    Dim dt As cDateTime
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    Assert.Equals CDec("632085156610010000"), dt.Ticks
End Sub

Public Sub TestDateOnly()
    Dim dt As cDateTime
    Dim dt2 As cDateTime
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    Set dt2 = dt.DateOnly
    Assert.Equals 2004, dt2.Year
    Assert.Equals 1, dt2.Month
    Assert.Equals 1, dt2.Day
    Assert.Equals 0, dt2.Hour
    Assert.Equals 0, dt2.Minute
    Assert.Equals 0, dt2.Second
    Assert.Equals 0, dt2.Millisecond
End Sub

Public Sub TestTimeOfDay()
    Dim dt As cDateTime
    Dim tod As TimeSpan
    
    Set dt = newcdatetime(2004, 1, 1, 6, 10, 15, 100)
    Set tod = dt.TimeOfDay
    Assert.Equals 6, tod.Hours
    Assert.Equals 10, tod.Minutes
    Assert.Equals 15, tod.Seconds
    Assert.Equals 100, tod.Milliseconds
End Sub

Public Sub TestAdd()
    Dim dt As cDateTime
    Dim ts As TimeSpan
    Dim i As Long
    
    Set ts = newtimespan(1, 1, 1, 1, 1)
    Set dt = New cDateTime
    For i = 1 To 100
        Set dt = dt.Add(ts)
    Next i
    Assert.Equals 1, dt.Year
    Assert.Equals 4, dt.Month
    Assert.Equals 15, dt.Day
    Assert.Equals 5, dt.Hour
    Assert.Equals 41, dt.Minute
    Assert.Equals 40, dt.Second
    Assert.Equals 100, dt.Millisecond
End Sub

Public Sub TestAddMilliseconds()
    Dim dt As cDateTime
    Dim i As Long
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    For i = 1 To 100
        Set dt = dt.AddMilliseconds(i)
    Next i
    Assert.Equals 1, dt.Month
    Assert.Equals 1, dt.Day
    Assert.Equals 2004, dt.Year
    Assert.Equals 1, dt.Hour
    Assert.Equals 1, dt.Minute
    Assert.Equals 6, dt.Second
    Assert.Equals 51, dt.Millisecond
End Sub

Public Sub TestAddDays()
    Dim dt As cDateTime
    Dim i As Long
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    For i = 1 To 100
        Set dt = dt.AddDays(i)
    Next i
    Assert.Equals 10, dt.Month
    Assert.Equals 29, dt.Day
    Assert.Equals 2017, dt.Year
    Assert.Equals 1, dt.Hour
    Assert.Equals 1, dt.Minute
    Assert.Equals 1, dt.Second
    Assert.Equals 1, dt.Millisecond
End Sub

Public Sub TestAddHours()
    Dim dt As cDateTime
    Dim i As Long
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    For i = 1 To 100
        Set dt = dt.AddHours(i)
    Next i
    Assert.Equals 7, dt.Month
    Assert.Equals 29, dt.Day
    Assert.Equals 2004, dt.Year
    Assert.Equals 11, dt.Hour
    Assert.Equals 1, dt.Minute
    Assert.Equals 1, dt.Second
    Assert.Equals 1, dt.Millisecond
End Sub

Public Sub TestAddMinutes()
    Dim dt As cDateTime
    Dim i As Long
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    For i = 1 To 100
        Set dt = dt.AddMinutes(i)
    Next i
    Assert.Equals 1, dt.Month
    Assert.Equals 4, dt.Day
    Assert.Equals 2004, dt.Year
    Assert.Equals 13, dt.Hour
    Assert.Equals 11, dt.Minute
    Assert.Equals 1, dt.Second
    Assert.Equals 1, dt.Millisecond
End Sub

Public Sub TestAddMonths()
    Dim dt As cDateTime
    Dim i As Long
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    For i = 1 To 100
        Set dt = dt.AddMonths(i)
    Next i
    Assert.Equals 11, dt.Month
    Assert.Equals 1, dt.Day
    Assert.Equals 2424, dt.Year
    Assert.Equals 1, dt.Hour
    Assert.Equals 1, dt.Minute
    Assert.Equals 1, dt.Second
    Assert.Equals 1, dt.Millisecond
End Sub

Public Sub TestDaysInMonth()
    Dim i As Long
    
    For i = 1 To 12
        Assert.Equals Day(DateAdd("d", -1, DateSerial(1999, i + 1, 1))), cDateTime.DaysInMonth(1999, i)
    Next i
    
    ' leap year
    For i = 1 To 12
        Assert.Equals Day(DateAdd("d", -1, DateSerial(2000, i + 1, 1))), cDateTime.DaysInMonth(2000, i)
    Next i
End Sub

Public Sub TestAddSeconds()
    Dim dt As cDateTime
    Dim i As Long
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1, 1)
    For i = 1 To 100
        Set dt = dt.AddSeconds(i)
    Next i
    Assert.Equals 1, dt.Month
    Assert.Equals 1, dt.Day
    Assert.Equals 2004, dt.Year
    Assert.Equals 2, dt.Hour
    Assert.Equals 25, dt.Minute
    Assert.Equals 11, dt.Second
    Assert.Equals 1, dt.Millisecond
End Sub

Public Sub TestAddYear()
    Dim dt As cDateTime
    
    Set dt = newcdatetime(2004, 5, 16)
    Set dt = dt.AddYears(5)
    Assert.Equals 2009, dt.Year
    Assert.Equals 5, dt.Month
    Assert.Equals 16, dt.Day
End Sub

Public Sub TestCompareTo()
    Dim dt As cDateTime
    Dim lessThan As cDateTime
    Dim greaterThan As cDateTime
    Dim equalTo As cDateTime
    
    Set dt = newcdatetime(2004, 5, 5, 5, 5, 5)
    Set lessThan = newcdatetime(2004, 5, 5, 5, 5, 4)
    Set greaterThan = newcdatetime(2004, 5, 5, 5, 5, 6)
    Set equalTo = newcdatetime(2004, 5, 5, 5, 5, 5)
    
    Assert.IsTrue dt.CompareTo(lessThan) > 0
    Assert.IsTrue dt.CompareTo(greaterThan) < 0
    Assert.IsTrue dt.CompareTo(equalTo) = 0
    Assert.IsTrue dt.CompareTo(Nothing) > 0
    Assert.IsTrue dt.CompareTo(Null) > 0
    Assert.IsTrue dt.CompareTo(#1/1/2004#) > 0
    Assert.IsTrue dt.CompareTo(#1/1/2005#) < 0
    Assert.IsTrue dt.CompareTo(#5/5/2004 5:05:05 AM#) = 0
End Sub

Public Sub TestFromOADate()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromOADate(#1/1/2004 8:30:30 AM#)
    Assert.Equals 1, dt.Month
    Assert.Equals 1, dt.Day
    Assert.Equals 2004, dt.Year
    Assert.Equals 8, dt.Hour
    Assert.Equals 30, dt.Minute
    Assert.Equals 30, dt.Second
    
    Set dt = cDateTime.FromOADate(#1/1/150 11:59:59 PM#)
    Assert.Equals 1, dt.Month
    Assert.Equals 1, dt.Day
    Assert.Equals 150, dt.Year
    Assert.Equals 23, dt.Hour
    Assert.Equals 59, dt.Minute
    Assert.Equals 59, dt.Second
End Sub

Public Sub TestEquals()
    Dim dt As cDateTime
    
    Set dt = newcdatetime(2004, 1, 1, 1, 1, 1)
    Assert.IsTrue dt.Equals(newcdatetime(2004, 1, 1, 1, 1, 1))
    Assert.IsTrue dt.Equals(#1/1/2004 1:01:01 AM#)
    Assert.IsFalse dt.Equals(Null)
    Assert.IsFalse dt.Equals(Nothing)
    Assert.IsFalse dt.Equals(newcdatetime(2005, 1, 1, 1, 1, 1))
End Sub

Public Sub TestSubtractTimeSpan()
    Dim dt As cDateTime
     
    Set dt = newcdatetime(2004, 6, 6, 6, 6, 6, 6)
    Set dt = dt.Subtract(newtimespan(6, 6, 6, 6, 6))
    Assert.Equals 2004, dt.Year
    Assert.Equals 5, dt.Month
    Assert.Equals 31, dt.Day
    Assert.Equals 0, dt.Hour
    Assert.Equals 0, dt.Minute
    Assert.Equals 0, dt.Second
    Assert.Equals 0, dt.Millisecond
End Sub

Public Sub TestSubtractcDateTime()
    Dim dt As cDateTime
    Dim ts As TimeSpan
    
    Set dt = newcdatetime(2004, 6, 6, 6, 6, 6, 6)
    Set ts = dt.Subtract(newcdatetime(2004, 6, 6))
    Assert.Equals 0, ts.Days
    Assert.Equals 6, ts.Hours
    Assert.Equals 6, ts.Minutes
    Assert.Equals 6, ts.Seconds
    Assert.Equals 6, ts.Milliseconds
End Sub

Public Sub TestSubtractDate()
    Dim dt As cDateTime
    Dim ts As TimeSpan
    
    Set dt = newcdatetime(2004, 6, 6, 6, 6, 6, 6)
    Set ts = dt.Subtract(#6/6/2004#)
    Assert.Equals 0, ts.Days
    Assert.Equals 6, ts.Hours
    Assert.Equals 6, ts.Minutes
    Assert.Equals 6, ts.Seconds
    Assert.Equals 6, ts.Milliseconds
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestToFileTime()
    Dim dt As cDateTime

    Set dt = newcdatetime(2004, 6, 6, 6, 6, 6, 6)
    Assert.Equals 12731000766006@, dt.ToFileTime
    
    Set dt = newcdatetime(2004, 1, 6, 6, 6, 6, 6)
    Assert.Equals 12717871566006@, dt.ToFileTime
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestToFileTimeUtc()
    Dim dt As cDateTime
    
    Set dt = newcdatetime(2004, 6, 6, 6, 6, 6, 6)
    Assert.Equals 12730975566006@, dt.ToFileTimeUtc
    
    Set dt = newcdatetime(2004, 1, 6, 6, 6, 6, 6)
    Assert.Equals 12717842766006@, dt.ToFileTimeUtc
End Sub

Public Sub TestNow()
    Dim dt As cDateTime
    Dim d As Date
    Dim i As Long
    
    Set dt = cDateTime.Now
    Set dt = dt.AddMilliseconds(-dt.Millisecond)
    Do While dt.CompareTo(d) <> 0 And i = 0
        d = Now
        i = i + 1
    Loop
    Assert.Equals 0, dt.CompareTo(d)
End Sub

Public Sub TestNowUtc()
    Dim dt As cDateTime
    Dim d As cDateTime
    Dim i As Long
    
    Set dt = cDateTime.NowUtc
    Set dt = dt.AddMilliseconds(-dt.Millisecond)
    Do While dt.CompareTo(d) <> 0 And i = 0
        Set d = cDateTime.FromOADate(Now)
        Set d = d.ToUniversalTime
        i = i + 1
    Loop
    Assert.Equals 0, dt.CompareTo(d)
End Sub

Public Sub TestToday()
    Dim dt As cDateTime
    Dim d As Date
    
    d = Date
    Set dt = cDateTime.Today
    Assert.Equals 0, dt.CompareTo(d)
End Sub

Public Sub TestParse()
    Dim dt As cDateTime
    
    Set dt = cDateTime.Parse("1/1/2004 8:30:30am")
    Assert.Equals 0, dt.CompareTo(#1/1/2004 8:30:30 AM#)
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestFromFileTime()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromFileTime(1000000000@)
    Assert.IsTrue dt.Equals(#1/12/1601 5:46:40 AM#), "FromFileTime should return 1/12/1601 5:46:40am."
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestFromFileTimeUtc()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromFileTimeUtc(1000000000@)
    Assert.IsTrue dt.Equals(#1/12/1601 1:46:40 PM#), "FromFileTime should return 1/12/1601 1:46:40 PM."
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestToLocalTime()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromOADate(#1/1/2004 12:00:00 PM#)
    Set dt = dt.ToLocalTime
    Assert.IsTrue dt.Equals(#1/1/2004 4:00:00 AM#), "ToLocalTime should return 1/1/2004 4AM."
    
    Set dt = cDateTime.FromOADate(#6/1/2004 12:00:00 PM#)
    Set dt = dt.ToLocalTime
    Assert.IsTrue dt.Equals(#6/1/2004 5:00:00 AM#), "ToLocalTime should return 6/1/2004 5AM."
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestToUniversalTime()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromOADate(#1/1/2004 12:00:00 PM#)
    Set dt = dt.ToUniversalTime
    Assert.IsTrue dt.Equals(#1/1/2004 8:00:00 PM#), "ToUniversalTime should return 1/1/2004 8PM."
    
    Set dt = cDateTime.FromOADate(#6/1/2004 12:00:00 PM#)
    Set dt = dt.ToUniversalTime
    Assert.IsTrue dt.Equals(#6/1/2004 7:00:00 PM#), "ToUniversalTime should return 6/1/2004 7PM."
End Sub
