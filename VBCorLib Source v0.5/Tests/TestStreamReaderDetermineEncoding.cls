VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStreamReaderDetermineEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamReaderDetermineEncoding
'
Option Explicit



Public Sub TestUTF8Encoding()
    TestEncoding cor.NewUTF8Encoding(True, False)
End Sub

Public Sub TestUnicodeEnocding()
    TestEncoding cor.NewUnicodeEncoding(False, True)
End Sub

Public Sub TestUnicodeEncodingBigEndian()
    TestEncoding cor.NewUnicodeEncoding(True, True)
End Sub



Private Sub TestEncoding(ByVal en As Encoding)
    Dim ms As MemoryStream
    Dim sr As StreamReader
    Dim b() As Byte
    Dim currEn As Encoding
    
    Set ms = New MemoryStream
    ms.WriteBlock en.GetPreamble, 0, carray.GetLength(en.GetPreamble)
    b = en.GetBytes(ChrW$(55))
    ms.WriteBlock b, 0, carray.GetLength(b)
    ms.Position = 0
    
    Set sr = cor.NewStreamReader(ms, , True)
    Set currEn = sr.CurrentEncoding
    
    Assert.Equals 55, sr.Read
    Assert.IsFalse sr.CurrentEncoding Is currEn
    Assert.Equals TypeName(en), TypeName(sr.CurrentEncoding)
End Sub
