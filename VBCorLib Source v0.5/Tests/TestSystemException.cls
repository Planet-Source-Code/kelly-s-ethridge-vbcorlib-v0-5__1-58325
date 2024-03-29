VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSystemException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EXPECTED_MESSAGE As String = "User Error."
Private Const MESSAGE_PREFIX As String = "VBCorLib.SystemException: "

Private ex As SystemException
Private Inner As Exception

Public Sub Setup()
    Set Inner = New Exception
    Set ex = cor.NewSystemException(EXPECTED_MESSAGE, Inner)
End Sub

Public Sub TestUserDefinedMessage()
    Assert.Equals EXPECTED_MESSAGE, ex.Message, "Message should return 'User Error.'"
End Sub

Public Sub TestUserDefinedInnerException()
    Assert.Equals Inner, ex.InnerException, "InnerException should return UserException instance."
End Sub

Public Sub TestGetBaseException()
    Dim firstLevelEx As SystemException
    Set firstLevelEx = cor.NewSystemException(EXPECTED_MESSAGE, ex)
    Assert.Equals Inner, firstLevelEx.GetBaseException, "GetBaseException should return Inner instance."
End Sub

Public Sub TestUserDefinedToString()
    Assert.Equals MESSAGE_PREFIX & EXPECTED_MESSAGE, ex.ToString, "ToString return wrong value."
End Sub

Public Sub TestPersitence()
    Dim copy As SystemException
    Dim Bag As New PropertyBag
    
    Bag.WriteProperty "ex", ex
    Set copy = Bag.ReadProperty("ex")
    
    Assert.Equals EXPECTED_MESSAGE, copy.Message, "Message should return 'User Error.'"
    Assert.IsNotNothing copy.InnerException, "InnerException should not be Nothing."
    
End Sub

