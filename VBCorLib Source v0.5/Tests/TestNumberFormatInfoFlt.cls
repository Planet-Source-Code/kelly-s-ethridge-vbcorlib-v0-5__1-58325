VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestNumberFormatInfoFlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestNumberFormatInfoFlt
'
Option Explicit

Private info As NumberFormatInfo

Public Sub Setup()
    Set info = New NumberFormatInfo
End Sub

Public Sub TestZero()
    Assert.Equals "0", info.Format(0#)
End Sub

Public Sub TestGeneral()
    Assert.Equals "12345", info.Format(12345#)
    Assert.Equals "-12345", info.Format(-12345#)
    Assert.Equals "1.2345", info.Format(1.2345)
    Assert.Equals "0.012345", info.Format(0.012345)
    Assert.Equals "1234.5678", info.Format(1234.5678)
    Assert.Equals "1234500000000000000000000000", info.Format(1.2345E+27, "g28")
    Assert.Equals "1.2345E-05", info.Format(0.000012345)
    Assert.Equals "1.2345E+15", info.Format(1.2345E+15)
    Assert.Equals "1.235e+04", info.Format(12345#, "g4")
End Sub

Public Sub TestFixed()
    Assert.Equals "12345.00", info.Format(12345#, "f")
    Assert.Equals "12345", info.Format(12345#, "f0")
    Assert.Equals "123450.0000000000", info.Format(123450#, "f10")
    Assert.Equals "123.45", info.Format(123.45, "f")
    Assert.Equals "1.23", info.Format(1.2345, "f")
    Assert.Equals "1.235", info.Format(1.2345, "f3")
    Assert.Equals "0.12", info.Format(0.12345, "f")
    Assert.Equals "0.000", info.Format(0.00012345, "f3")
    Assert.Equals "0.0001235", info.Format(0.00012345, "f7")
    Assert.Equals "0.00012345000000000000", info.Format(0.00012345, "f20")
End Sub

Public Sub TestScientific()
    Assert.Equals "1.234500e+004", info.Format(12345#, "e")
    Assert.Equals "-1.2345e+004", info.Format(-12345#, "e4")
    Assert.Equals "1.235e+004", info.Format(12345#, "e3")
    Assert.Equals "1e+004", info.Format(12345#, "e0")
    Assert.Equals "1.2E+004", info.Format(12345#, "E1")
    Assert.Equals "1.234500e+006", info.Format(1234500#, "e")
    Assert.Equals "1.235e+006", info.Format(1234500#, "e3")
    Assert.Equals "1.234500e-003", info.Format(0.0012345, "e")
    Assert.Equals "1.235e-003", info.Format(0.0012345, "e3")
    Assert.Equals "1.2345000e-200", info.Format(1.2345E-200, "e7")
End Sub

Public Sub TestDecimal()
    On Error GoTo errTrap
    info.Format 1#, "d"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is FormatException, "Wrong exception type caught."
End Sub

Public Sub TestNumber()
    Assert.Equals "1.23", info.Format(1.23, "n")
    Assert.Equals "1,234.56", info.Format(1234.56, "n")
    Assert.Equals "1,234,567.89", info.Format(1234567.89, "n")
    Assert.Equals "12,345,678.91", info.Format(12345678.91, "n")
    Assert.Equals "123", info.Format(123.45, "n0")
    Assert.Equals "123.45678000", info.Format(123.45678, "n8")
    Assert.Equals "-123", info.Format(-123.45, "n0")
    Assert.Equals "0.00", info.Format(0.00123, "n")
    Assert.Equals "0.00124", info.Format(0.001235, "n5")
    Assert.Equals "123.00", info.Format(123#, "n")
    Assert.Equals "123,123,123,000,000,000,000.00000", info.Format(1.23123123E+20, "n5")
    
    info.NumberNegativePattern = [- n]
    Assert.Equals "- 123.00", info.Format(-123#, "n")

    info.NumberNegativePattern = [n-]
    Assert.Equals "123.00-", info.Format(-123#, "n")

    info.NumberNegativePattern = [n -]
    Assert.Equals "123.00 -", info.Format(-123#, "n")

    info.NumberNegativePattern = [(n)]
    Assert.Equals "(123.00)", info.Format(-123#, "n")

    Dim sizes() As Long
    sizes = carray.NewArray(ciLong, 1, 2, 3)
    info.NumberGroupSizes = sizes
    Assert.Equals "12,345,67,8.91", info.Format(12345678.91, "n")

    sizes = carray.NewArray(ciLong, 1, 2, 1, 0)
    info.NumberGroupSizes = sizes
    Assert.Equals "12345,6,78,9.12", info.Format(123456789.12, "n")
    
    info.NumberGroupSizes = carray.CreateInstance(ciLong)
    Assert.Equals "1234567890", info.Format(1234567890#, "n0")
End Sub

Public Sub TestCurrency()
    Assert.Equals "¤1.23", info.Format(1.23, "c")
    Assert.Equals "¤1,234.56", info.Format(1234.56, "c")
    Assert.Equals "¤1,234,567.89", info.Format(1234567.89, "c")
    Assert.Equals "¤12,345,678.91", info.Format(12345678.91, "c")
    Assert.Equals "¤123", info.Format(123.45, "c0")
    Assert.Equals "¤123.45678000", info.Format(123.45678, "c8")
    Assert.Equals "(¤123)", info.Format(-123.45, "c0")
    Assert.Equals "¤0.00", info.Format(0.00123, "c")
    Assert.Equals "¤0.00124", info.Format(0.001235, "c5")
    Assert.Equals "¤123.00", info.Format(123#, "c")
    Assert.Equals "¤123,123,123,000,000,000,000.00000", info.Format(1.23123123E+20, "c5")
    
    info.CurrencyNegativePattern = [-$ n]
    Assert.Equals "-¤ 123.00", info.Format(-123#, "c")

    info.CurrencyNegativePattern = [$ n-]
    Assert.Equals "¤ 123.00-", info.Format(-123#, "c")

    info.CurrencyNegativePattern = [(n $)]
    Assert.Equals "(123.00 ¤)", info.Format(-123#, "c")

    info.CurrencyNegativePattern = [($n)]
    Assert.Equals "(¤123.00)", info.Format(-123#, "c")

    info.CurrencyGroupSizes = carray.NewArray(ciLong, 1, 2, 3)
    Assert.Equals "¤12,345,67,8.91", info.Format(12345678.91, "c")

    info.CurrencyGroupSizes = carray.NewArray(ciLong, 1, 2, 1, 0)
    Assert.Equals "¤12345,6,78,9.12", info.Format(123456789.12, "c")
    
    info.CurrencyGroupSizes = carray.CreateInstance(ciLong)
    Assert.Equals "¤1234567890", info.Format(1234567890#, "c0")
End Sub

Public Sub TestPercent()
    Assert.Equals "0.00 %", info.Format(0#, "p")
    Assert.Equals "123.00 %", info.Format(1.23, "p")
    Assert.Equals "123,456.00 %", info.Format(1234.56, "p")
    Assert.Equals "123,456,789.00 %", info.Format(1234567.89, "p")
    Assert.Equals "1,234,567,891.00 %", info.Format(12345678.91, "p")
    Assert.Equals "12,345 %", info.Format(123.45, "p0")
    Assert.Equals "12,345.67800000 %", info.Format(123.45678, "p8")
    Assert.Equals "-12,345 %", info.Format(-123.45, "p0")
    Assert.Equals "0.12 %", info.Format(0.00123, "p")
    Assert.Equals "0.124 %", info.Format(0.001235, "p3")
    Assert.Equals "12,300.00 %", info.Format(123#, "p")
    Assert.Equals "12,312,312,300,000,000,000,000.00000 %", info.Format(1.23123123E+20, "p5")
    
    info.PercentNegativePattern = [-# %]
    Assert.Equals "-12,300.00 %", info.Format(-123#, "p")

    info.PercentNegativePattern = [-#%]
    Assert.Equals "-12,300.00%", info.Format(-123#, "p")

    info.PercentNegativePattern = [-%#]
    Assert.Equals "-%12,300.00", info.Format(-123#, "p")

    info.PercentGroupSizes = carray.NewArray(ciLong, 1, 2, 3)
    Assert.Equals "1,234,567,89,1.00 %", info.Format(12345678.91, "p")

    info.PercentGroupSizes = carray.NewArray(ciLong, 1, 2, 1, 0)
    Assert.Equals "1234567,8,91,2.00 %", info.Format(123456789.12, "p")
    
    info.PercentGroupSizes = carray.CreateInstance(ciLong)
    Assert.Equals "123456789000 %", info.Format(1234567890#, "p0")
End Sub

Public Sub TestNAN()
    Dim d As Double
    On Error Resume Next
    d = 0# / 0#
    On Error GoTo 0
    Assert.Equals "NaN", info.Format(d)
End Sub

Public Sub TestPosInfinity()
    Dim d As Double
    On Error Resume Next
    d = 1# / 0#
    On Error GoTo 0
    Assert.Equals "Infinity", info.Format(d)
End Sub

Public Sub TestNegInfinity()
    Dim d As Double
    On Error Resume Next
    d = -1# / 0#
    On Error GoTo 0
    Assert.Equals "-Infinity", info.Format(d)
End Sub
