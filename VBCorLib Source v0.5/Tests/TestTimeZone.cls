VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTimeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' change these tests to match your specific time zone.
'

Private zone As timezone

Public Sub Setup()
    Set zone = timezone.CurrentTimeZone
End Sub

Public Sub TestDayLightName()
    Assert.Equals "Pacific Daylight Time", zone.DayLightName
End Sub

Public Sub TestStandardName()
    Assert.Equals "Pacific Standard Time", zone.StandardName
End Sub

Public Sub TestGetUtcOffsetOutsideOfDaylightSaving()
    Dim ts As TimeSpan
    
    Set ts = zone.GetUtcOffset(#1/1/2004#)
    Assert.Equals -8#, ts.TotalHours, "Should have an offset of -8 hours."
    
    Set ts = zone.GetUtcOffset(#6/6/2004#)
    Assert.Equals -7#, ts.TotalHours, "Should have an offset of -7 hours."
End Sub

Public Sub TestIsDayLightSavings()
    Assert.IsFalse zone.IsDayLightSavingTime(#4/4/2004 1:59:00 AM#), "4/4/2004 1:59am Should not be DLS."
    Assert.IsTrue zone.IsDayLightSavingTime(#4/4/2004 2:00:00 AM#), "4/4/2004 2am Should be DLS."
    Assert.IsTrue zone.IsDayLightSavingTime(#10/31/2004 1:59:00 AM#), "10/31/2004 1:59am should be DLS."
    Assert.IsFalse zone.IsDayLightSavingTime(#10/31/2004 2:00:00 AM#), "10/31/2004 2am should not be DLS."
End Sub

Public Sub TestDayLightChanges2003()
    Dim dlt As DayLightTime
    
    Set dlt = zone.GetDayLightChanges(2003)
    Assert.IsTrue dlt.StartTime.Equals(#4/6/2003 2:00:00 AM#)
    Assert.IsTrue dlt.EndTime.Equals(#10/26/2003 2:00:00 AM#)
End Sub

Public Sub TestToLocalTime()
    Dim dt As cDateTime
    
    Set dt = zone.ToLocalTime(#1/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#1/1/2004 4:00:00 AM#), "ToLocalTime should return 1/1/2004 4am."
    
    Set dt = zone.ToLocalTime(#6/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#6/1/2004 5:00:00 AM#), "ToLocalTime should return 6/1/2004 5am."
End Sub

Public Sub TestToUniversalTime()
    Dim dt As cDateTime
    
    Set dt = zone.ToUniversalTime(#1/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#1/1/2004 8:00:00 PM#), "ToUniversalTime should return 1/1/2004 8pm."
    
    Set dt = zone.ToUniversalTime(#6/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#6/1/2004 7:00:00 PM#), "ToUniversalTime should return 6/1/2004 7pm."
End Sub
