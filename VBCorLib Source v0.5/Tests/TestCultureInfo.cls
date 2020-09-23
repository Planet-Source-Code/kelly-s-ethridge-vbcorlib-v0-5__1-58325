VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCultureInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCultureInfo
'
Option Explicit

Public Sub TestInvariant()
    Dim c As CultureInfo
    
    Set c = cor.newcultureinfo(127)
    Assert.Equals &H7F, c.LCID
    Assert.Equals "", c.Name
    Assert.Equals "Invariant Language (Invariant Country)", c.EnglishName
    Assert.Equals "Invariant Language (Invariant Country)", c.DisplayName
    Assert.IsFalse c.IsNeutralCulture, "Culture should not be neutral."
    Assert.Equals "Invariant Language (Invariant Country)", c.NativeName
    Assert.Equals "IVL", c.ThreeLetterISOLanguageName
    Assert.Equals "IVL", c.ThreeLetterWindowsLanguageName
    Assert.Equals "iv", c.TwoLetterISOLanguageName
    Assert.IsFalse c.UseUserOverride, "UseUserOverride should return false."
    
    Dim dt As datetimeformatinfo
    Dim i As Long
    Dim s() As String
    
    Set dt = c.DateTimeFormat
    s = dt.AbbreviatedDayNames
    Assert.Equals 7, carray.GetLength(s), "AbbreviatedDayNames should return 7 elements."
    For i = 0 To 6
        Assert.Equals Choose(i + 1, "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"), s(i)
    Next i
    
    Assert.Equals "dddd, dd MMMM yyyy", dt.LongDatePattern, "Incorrect LongDatePattern."
    Assert.Equals "HH:mm:ss", dt.LongTimePattern, "Incorrect LongTimePattern."
    
    Dim nm As NumberFormatInfo
    Dim n() As Long
    
    Set nm = c.NumberFormat
    Assert.Equals ChrW$(&H2030), nm.PerMilleSymbol
    
    n = nm.NumberGroupSizes
    Assert.Equals 1, carray.GetLength(n)
    Assert.Equals 3, n(0)
End Sub

Public Sub TestEnglishUS()
    Dim c As CultureInfo
    
    Set c = cor.newcultureinfo("en-US")
    Assert.Equals 1033, c.LCID
    Assert.Equals "en-US", c.Name
    Assert.Equals "English (United States)", c.EnglishName
    Assert.Equals "English (United States)", c.DisplayName
    Assert.IsFalse c.IsNeutralCulture, "Culture should not be neutral."
    Assert.Equals "English (United States)", c.NativeName
    Assert.Equals "eng", c.ThreeLetterISOLanguageName
    Assert.Equals "ENU", c.ThreeLetterWindowsLanguageName
    Assert.Equals "en", c.TwoLetterISOLanguageName
    
    Dim dt As datetimeformatinfo
    Dim i As Long
    Dim s() As String
    
    Set dt = c.DateTimeFormat
    s = dt.DayNames
    Assert.Equals 7, carray.GetLength(s), "AbbreviatedDayNames should return 7 elements."
    For i = 0 To 6
        Assert.Equals Choose(i + 1, "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"), s(i)
    Next i
    
    Assert.Equals "dddd, MMMM dd, yyyy", dt.LongDatePattern, "Incorrect LongDatePattern."
    Assert.Equals "h:mm:ss tt", dt.LongTimePattern, "Incorrect LongTimePattern."
    
End Sub
