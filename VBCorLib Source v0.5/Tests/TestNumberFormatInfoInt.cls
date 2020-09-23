VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestNumberFormatInfoInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private info As NumberFormatInfo

Public Sub Setup()
    Set info = New NumberFormatInfo
End Sub


Public Sub TestZero()
    Assert.Equals "0", info.Format(0)
End Sub

Public Sub TestGeneral()
    Assert.Equals "12300", info.Format(12300, "")
    Assert.Equals "12345", info.Format(12345)
    Assert.Equals "1.23e+04", info.Format(12345, "g3")
    Assert.Equals "1.235e+04", info.Format(12345, "g4")
    Assert.Equals "1.23E+04", info.Format(12345, "G3")
    Assert.Equals "-12345", info.Format(-12345)
    Assert.Equals "2e+01", info.Format(15, "g1")
    Assert.Equals "12345", info.Format(12345, "g26")
    Assert.Equals "ggg", info.Format(12345, "ggg")
End Sub

Public Sub TestFixed()
    Assert.Equals "12345.00", info.Format(12345, "f")
    Assert.Equals "-12345", info.Format(-12345, "f0")
    Assert.Equals "12345.0000000000", info.Format(12345, "f10")
    Assert.Equals "1234500.000", info.Format(1234500, "F3")
    Assert.Equals "fff", info.Format(12345, "fff")
End Sub

Public Sub TestScientific()
    Assert.Equals "1.234500e+004", info.Format(12345, "e")
    Assert.Equals "-1.2345e+004", info.Format(-12345, "e4")
    Assert.Equals "1.235e+004", info.Format(12345, "e3")
    Assert.Equals "1e+004", info.Format(12345, "e0")
    Assert.Equals "1.2E+004", info.Format(12345, "E1")
End Sub

Public Sub TestDecimal()
    Assert.Equals "12345", info.Format(12345, "d")
    Assert.Equals "12345", info.Format(12345, "d1")
    Assert.Equals "0012345", info.Format(12345, "d7")
End Sub

Public Sub TestHexidecimal()
    Assert.Equals "A0", info.Format(160, "X")
    Assert.Equals "ab", info.Format(171, "x")
    Assert.Equals "ffffffff", info.Format(-1&, "x")
    Assert.Equals "ffff", info.Format(-1 And &HFFFF&, "x")
    Assert.Equals "00123", info.Format(&H123, "x5")
    Assert.Equals "ABCDE", info.Format(&HABCDE, "X1")
    Assert.Equals "000", info.Format(0, "x3")
    Assert.Equals "ff", info.Format(CByte(255), "x")
End Sub

Public Sub TestNumber()
    Assert.Equals "0.00", info.Format(0, "n")
    Assert.Equals "1.00", info.Format(1, "n")
    Assert.Equals "1,234.00", info.Format(1234, "n")
    Assert.Equals "1,234,567.00", info.Format(1234567, "n")
    Assert.Equals "12,345,678.00", info.Format(12345678, "n")
    Assert.Equals "123", info.Format(123, "n0")
    Assert.Equals "123.00000", info.Format(123, "n5")
    Assert.Equals "-123", info.Format(-123, "n0")

    info.NumberNegativePattern = [- n]
    Assert.Equals "- 123.00", info.Format(-123, "n")

    info.NumberNegativePattern = [n-]
    Assert.Equals "123.00-", info.Format(-123, "n")

    info.NumberNegativePattern = [n -]
    Assert.Equals "123.00 -", info.Format(-123, "n")

    info.NumberNegativePattern = [(n)]
    Assert.Equals "(123.00)", info.Format(-123, "n")

    Dim sizes() As Long
    sizes = carray.NewArray(ciLong, 1, 2, 3)
    info.NumberGroupSizes = sizes
    Assert.Equals "12,345,67,8.00", info.Format(12345678, "n")

    sizes = carray.NewArray(ciLong, 1, 2, 1, 0)
    info.NumberGroupSizes = sizes
    Assert.Equals "12345,6,78,9.00", info.Format(123456789, "n")
End Sub

Public Sub TestCurrency()
    Assert.Equals "¤0.00", info.Format(0, "c")
    Assert.Equals "¤1.00", info.Format(1, "c")
    Assert.Equals "¤1,234.00", info.Format(1234, "c")
    Assert.Equals "¤1,234,567.00", info.Format(1234567, "c")
    Assert.Equals "¤12,345,678.00", info.Format(12345678, "c")
    Assert.Equals "¤123", info.Format(123, "c0")
    Assert.Equals "¤123.00000", info.Format(123, "c5")
    Assert.Equals "(¤123)", info.Format(-123, "c0")
    
    info.CurrencyNegativePattern = [-$ n]
    Assert.Equals "-¤ 123.00", info.Format(-123, "c")
    
    info.CurrencyNegativePattern = [$ n-]
    Assert.Equals "¤ 123.00-", info.Format(-123, "c")
    
    info.CurrencyGroupSizes = carray.NewArray(ciLong, 1, 2, 3)
    Assert.Equals "¤12,345,67,8.00", info.Format(12345678, "c")
    
    info.CurrencyGroupSizes = carray.NewArray(ciLong, 1, 2, 1, 0)
    Assert.Equals "¤12345,6,78,9.00", info.Format(123456789, "c")
End Sub

Public Sub TestPercent()
    Assert.Equals "0.00 %", info.Format(0, "p")
    Assert.Equals "100.00 %", info.Format(1, "p")
    Assert.Equals "123,400.00 %", info.Format(1234, "p")
    Assert.Equals "123,456,700.00 %", info.Format(1234567, "p")
    Assert.Equals "1,234,567,800.00 %", info.Format(12345678, "p")
    Assert.Equals "12,300 %", info.Format(123, "p0")
    Assert.Equals "12,300.00000 %", info.Format(123, "p5")
    Assert.Equals "-12,300 %", info.Format(-123, "p0")
    
    info.PercentNegativePattern = [-# %]
    Assert.Equals "-12,300.00 %", info.Format(-123, "p")
    
    info.PercentNegativePattern = [-#%]
    Assert.Equals "-12,300.00%", info.Format(-123, "p")
    
    info.PercentNegativePattern = [-%#]
    Assert.Equals "-%12,300.00", info.Format(-123, "p")
    
    info.PercentGroupSizes = carray.NewArray(ciLong, 1, 2, 3)
    Assert.Equals "1,234,567,80,0.00 %", info.Format(12345678, "p")
    
    info.PercentGroupSizes = carray.NewArray(ciLong, 1, 2, 1, 0)
    Assert.Equals "1234567,8,90,0.00 %", info.Format(123456789, "p")
End Sub

