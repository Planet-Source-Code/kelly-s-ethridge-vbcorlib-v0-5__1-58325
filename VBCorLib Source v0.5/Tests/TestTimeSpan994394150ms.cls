VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTimeSpan994394150ms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTimeSpan90061010
'
Option Explicit

Private ts As timespan

Public Sub Setup()
    Set ts = timespan.FromTicks(9943941500123#)
End Sub

Public Sub TestTotalMilliseconds()
    Assert.Equals 994394150.0123@, ts.TotalMilliseconds, "TotalMilliseconds should return 994394150.0123."
End Sub

Public Sub TestTotalSeconds()
    Assert.Equals 994394.1500123, ts.TotalSeconds, "Total seconds should be 994394.1500123"
End Sub

Public Sub TestTotalMinutes()
    Assert.IsCloseTo 16573.2358333333, ts.TotalMinutes, 0.000001, "Total minutes should be approx 16573.2358333333"
End Sub

Public Sub TestTotalHours()
    Assert.IsCloseTo 276.220597222, ts.TotalHours, 0.0000001, "Total hours should be approx 276.220597222"
End Sub

Public Sub TestTotalDays()
    Assert.IsCloseTo 11.5091915510683, ts.TotalDays, 0.000000000001, "Total days should be approx 11.5091915510683"
End Sub

Public Sub TestMilliseconds()
    Assert.Equals 150, ts.Milliseconds, "Milliseconds should return 150."
End Sub

Public Sub TestSeconds()
    Assert.Equals 14, ts.Seconds, "Seconds should return 14."
End Sub

Public Sub TestMinutes()
    Assert.Equals 13, ts.Minutes, "Minutes should return 13."
End Sub

Public Sub TestHours()
    Assert.Equals 12, ts.Hours, "Hours should return 12."
End Sub

Public Sub TestDays()
    Assert.Equals 11, ts.Days, "Days should return 11."
End Sub

Public Sub TestTicks()
    Assert.Equals CDec("9943941500123"), ts.Ticks, "Ticks should return 9943941500000."
End Sub
