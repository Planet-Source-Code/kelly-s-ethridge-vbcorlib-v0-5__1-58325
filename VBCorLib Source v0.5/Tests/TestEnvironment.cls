VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestEnvironment
'
Option Explicit

' These don't really test the Environment class since the results
' will vary from machine to machine. Instead they simply display
' the results for the current machine for visual inspection.
Public Sub TestFixtureSetup()
    Console.WriteLine ""
    Console.WriteLine "***** ENVIRONMENT CLASS *****"
End Sub

Public Sub TestCommandLine()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "Command Line: " & environment.CommandLine
End Sub

Public Sub TestMachineName()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "Machine Name: " & environment.MachineName
End Sub

Public Sub TestOSVersion()
    Dim os As OperatingSystem
    
    Set os = environment.OSVersion
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "OS: " & os.ToString
End Sub

Public Sub TestSystemDirectory()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "System Directory: " & environment.SystemDirectory
End Sub

Public Sub TestTickCount()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "TickCount: " & environment.TickCount
End Sub

Public Sub TestUserName()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "UserName: " & environment.userName
End Sub

Public Sub TestUserDomainName()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "UserDomainName: " & environment.UserDomainName
End Sub

Public Sub TestUserInteractive()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "UserInteractive: " & environment.UserInteractive
End Sub

Public Sub TestWorkingSet()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "Returns 0 for 95, 98 and ME installations."
    Console.WriteLine "WorkingSet (Memory Usage):" & environment.WorkingSet
End Sub

Public Sub TestExpandEnvironmentVariables()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "--- ExpandingEnvironmentVariables ---"
    Console.WriteLine "'My OS is %OS%' becomes: " & environment.ExpandEnvironmentVariables("My OS is %OS%")
End Sub

Public Sub TestGetCommandLineArgs()
    Dim s() As String
    Dim i As Long
    
    s = environment.GetCommandLineArgs
    
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "--- GetCommandLingArgs ---"
    Console.WriteLine "Should have no empty args unless quoted."
    Console.WriteLine "Count: " & carray.GetLength(s)
    Console.WriteLine "Program: " & s(0)
    For i = 1 To UBound(s)
        Console.WriteLine "Arg(" & i & ") = " & s(i)
    Next i
End Sub

Public Sub TestGetEnvironmentVariable()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "GetEnvironmentVariable (os): " & environment.GetEnvironmentVariable("os")
End Sub

Public Sub TestGetEnvironmentVariables()
    Dim ht As Hashtable
    Dim i As Long
    
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "--- TestGetEnvironmentVariables ---"
    
    Set ht = environment.GetEnvironmentVariables
    Console.WriteLine "Count: " & ht.Count
    
    Dim de As DictionaryEntry
    For Each de In ht
        i = i + 1
        Console.WriteLine "Variable(" & i & "): " & de.Key & " = " & de.value
    Next de
End Sub

Public Sub TestGetFolderPath()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "--- TestGetFolderPath ---"
    Console.WriteLine "My Documents: " & environment.GetFolderPath(Personal)
    Console.WriteLine "Desktop: " & environment.GetFolderPath(DesktopDirectory)
    Console.WriteLine "Windows: " & environment.GetFolderPath(Windows)
    Console.WriteLine "System: " & environment.GetFolderPath(SystemDir)
End Sub

Public Sub TestGetLogicalDrives()
    Dim s() As String
    Dim i As Long
    
    s = environment.GetLogicalDrives
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "--- GetLogicalDrives ---"
    For i = 0 To UBound(s)
        Console.WriteLine "Drive - " & s(i)
    Next i
End Sub
