VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTimeSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTimeSpan
'
Option Explicit
Private ts As timespan

Public Sub TestZero()
    Set ts = timespan.Zero
    Assert.Equals 0@, ts.TotalMilliseconds, "TotalMilliseconds should return 0."
End Sub

Public Sub TestEquals()
    Dim ts1 As timespan
    Dim ts2 As timespan
    
    Set ts1 = timespan.FromHours(2.5)
    Set ts2 = timespan.FromMinutes(150)
    Assert.IsTrue ts1.Equals(ts2), "ts1 should equal ts2."
End Sub

Public Sub TestAdd()
    Dim ts1 As timespan
    Dim ts2 As timespan
    Dim sum As timespan
    
    Set ts1 = timespan.FromHours(1)
    Set ts2 = timespan.FromMinutes(30)
    Set sum = ts1.Add(ts2)
    Assert.Equals 90#, sum.TotalMinutes, "TotalMinutes should return 90."
End Sub

Public Sub TestDuration()
    Dim negTS As timespan
    Dim posts As timespan
    
    Set negTS = timespan.FromHours(-1.5)
    Set posts = negTS.Duration
    Assert.Equals 90#, posts.TotalMinutes, "TotalMinutes should return 90."
End Sub

Public Sub TestCompareTo()
    Set ts = timespan.FromHours(1)
    
    Assert.Equals -1, ts.CompareTo(timespan.FromMinutes(61)), "CompareTo should return -1."
    Assert.Equals 0, ts.CompareTo(timespan.FromMinutes(60)), "CompareTo should return 0."
    Assert.Equals 1, ts.CompareTo(timespan.FromMilliseconds(59)), "CompareTo should return 1."
    Assert.Equals 1, ts.CompareTo(Nothing), "CompareTo should return 1."
    Assert.Equals 1, ts.CompareTo(Null), "CompareTo should return 1."
End Sub

Public Sub TestNegate()
    Dim negTS As timespan
    
    Set ts = timespan.FromHours(1)
    Set negTS = ts.Negate
    Assert.Equals -60#, negTS.TotalMinutes, "TotalMinutes should return -60."
End Sub

Public Sub TestSubtrace()
    Dim sum As timespan
    
    Set ts = timespan.FromHours(1)
    Set sum = ts.Subtract(timespan.FromMinutes(30))
    Assert.Equals 30#, sum.TotalMinutes, "TimeMinutes should return 30."
End Sub

Public Sub TestToString()
    Set ts = newtimespan(2, 3, 4, 5, 10)
    Assert.Equals "2.03:04:05.0100000", ts.ToString
End Sub

Public Sub TestGetHashCode()
    Dim ts2 As timespan
    Dim ts3 As timespan
    
    Set ts = newtimespan(1, 2, 3, 4, 5)
    Set ts2 = newtimespan(1, 2, 3, 4, 5)
    Set ts3 = newtimespan(1, 2, 3, 4, 6)
    
    Assert.IsTrue ts.GetHashCode = ts2.GetHashCode, "Hashcodes should be same."
    Assert.IsFalse ts.GetHashCode = ts3.GetHashCode, "Hashcode should be different."
End Sub

Public Sub TestNegToString()
    Set ts = newtimespan(-2, -3, -4, -5, -10)
    Assert.Equals "-2.03:04:05.0100000", ts.ToString
    
End Sub

