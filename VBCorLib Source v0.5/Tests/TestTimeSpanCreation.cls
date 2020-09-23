VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTimeSpanCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTimeSpanCreation
'
Option Explicit

Private ts As timespan


Public Sub TestFromMilliseconds()
    Set ts = timespan.FromMilliseconds(1234567)
    Assert.Equals 1234567@, ts.TotalMilliseconds
End Sub

Public Sub TestFromSeconds()
    Set ts = timespan.FromSeconds(12345789)
    Assert.Equals 12345789#, ts.TotalSeconds
End Sub

Public Sub TestFromMinutes()
    Set ts = timespan.FromMinutes(123456789)
    Assert.Equals 123456789#, ts.TotalMinutes
End Sub

Public Sub TestFromHours()
    Set ts = timespan.FromHours(123456789)
    Assert.Equals 123456789#, ts.TotalHours
End Sub

Public Sub TestFromDays()
    Set ts = timespan.FromDays(1234567)
    Assert.Equals 1234567#, ts.TotalDays
End Sub

Public Sub TestFromTicks()
    Set ts = timespan.FromTicks(CDec("123456789012345"))
    Assert.Equals CDec("123456789012345"), ts.Ticks
End Sub

Public Sub TestParseHMMSS()
    Dim s As String
    Dim ts As timespan
    
    s = "   1:06:07  "
    Set ts = timespan.Parse(s)
    Assert.Equals 1, ts.Hours
    Assert.Equals 6, ts.Minutes
    Assert.Equals 7, ts.Seconds
End Sub

Public Sub TestParseHHMMSS()
    Dim s As String
    Dim ts As timespan
    
    s = "22:14:14"
    Set ts = timespan.Parse(s)
    Assert.Equals 22, ts.Hours
    Assert.Equals 14, ts.Minutes
    Assert.Equals 14, ts.Seconds
End Sub

Public Sub TestNegativeHHMMSS()
    Dim s As String
    Dim ts As timespan
    
    s = "-22:14:14"
    Set ts = timespan.Parse(s)
    Assert.Equals -22, ts.Hours
    Assert.Equals -14, ts.Minutes
    Assert.Equals -14, ts.Seconds
End Sub

Public Sub TestParseDHHMMSS()
    Dim s As String
    Dim ts As timespan
    
    s = "45.5:05:05"
    Set ts = timespan.Parse(s)
    Assert.Equals 45, ts.Days
    Assert.Equals 5, ts.Hours
    Assert.Equals 5, ts.Minutes
    Assert.Equals 5, ts.Seconds
End Sub

Public Sub TestParseHHMMSSFF()
    Dim s As String
    Dim ts As timespan
    
    s = "5:06:07.1"
    Set ts = timespan.Parse(s)
    Assert.Equals 5, ts.Hours
    Assert.Equals 6, ts.Minutes
    Assert.Equals 7, ts.Seconds
    Assert.Equals 100, ts.Milliseconds
End Sub

Public Sub TestParseDHHMMSSFF()
    Dim s As String
    Dim ts As timespan
    
    s = "-7.15:35:41.7777777"
    Set ts = timespan.Parse(s)
    Assert.Equals CDec(-6609417777777#), ts.Ticks
End Sub

Public Sub TestParseFormatError()
    Dim s As String
    Dim ts As timespan
    
    s = "15"
    On Error GoTo errTrap
    Set ts = timespan.Parse(s)
    On Error GoTo 0
    Assert.Fail "An error should be thrown."

errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is FormatException, "Wrong exception type caught."
End Sub
