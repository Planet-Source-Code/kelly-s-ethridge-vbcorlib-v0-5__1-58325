VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestVersion
'
Option Explicit

Public Sub TestZeroVersion()
    Dim v As New version
    Assert.Equals "0.0", v.ToString, "ToString should return 0.0"
End Sub

Public Sub TestOnePtTwo()
    Dim v As version
    Set v = newversion(1, 2)
    Assert.Equals "1.2", v.ToString, "ToString should return 1.2"
End Sub

Public Sub TestOnePtTwoPtThree()
    Dim v As version
    Set v = newversion(1, 2, 3)
    Assert.Equals "1.2.3", v.ToString
End Sub

Public Sub TestOnePtTwoPtThreeptFour()
    Dim v As version
    Set v = newversion(1, 2, 3, 4)
    Assert.Equals "1.2.3.4", v.ToString
End Sub

Public Sub TestToStringFieldCount()
    Dim v As version
    Set v = newversion(1, 2, 3, 4)
    Assert.Equals "", v.ToString(0)
    Assert.Equals "1", v.ToString(1)
    Assert.Equals "1.2", v.ToString(2)
    Assert.Equals "1.2.3", v.ToString(3)
    Assert.Equals "1.2.3.4", v.ToString(4)
End Sub

Public Sub TestClone()
    Dim v As version
    Set v = newversion(1, 2, 3, 4)
    
    Dim v2 As version
    Set v2 = v.Clone
    Assert.Equals "1.2.3.4", v2.ToString
End Sub

Public Sub TestCompareTo()
    Dim v As version
    Set v = newversion(1, 2, 3, 4)
    Assert.Equals 0, v.CompareTo(newversion(1, 2, 3, 4)), "CompareTo should return zero."
    Assert.Equals -1, v.CompareTo(newversion(1, 2, 4)), "CompareTo should return negative value."
    Assert.Equals 1, v.CompareTo(newversion(1, 2, 2)), "CompareTo should return positive value."
    Assert.Equals 1, v.CompareTo(Nothing), "CompareTo should return positive value."
End Sub

Public Sub TestEquals()
    Dim v As version
    Set v = newversion(1, 2)
    Assert.IsTrue v.Equals(newversion(1, 2)), "Equals should return true."
End Sub

Public Sub TestGetHashCode()
    Dim v As version
    Set v = newversion(1, 2, 3, 4)
    Assert.Equals 270544900, v.GetHashCode, "Wrong hash code."
End Sub

Public Sub TestPersistence()
    Dim v1 As version
    Set v1 = newversion(1, 2, 3, 4)
    
    Dim pp1 As New PropertyBag
    Dim pp2 As New PropertyBag
    Dim v2 As version
    
    pp1.WriteProperty "Version", v1
    pp2.Contents = pp1.Contents
    Set v2 = pp2.ReadProperty("Version")
    
    Assert.IsTrue v2.Equals(v1), "V2 should equal V1."
End Sub

Public Sub TestParse()
    Dim v As version
    
    Set v = version.Parse("1.2.3.4")
    Assert.IsTrue v.Equals(newversion(1, 2, 3, 4)), "Versions should be equal."
End Sub
