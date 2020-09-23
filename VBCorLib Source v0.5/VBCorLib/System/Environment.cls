VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: Environment
'

''
' The Environment class provides functions for retrieving information about the local machine
'
Option Explicit

' APIs defined here are to allow for easy access to the ANSI version of the call with strings.
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetEnvironmentVariableA Lib "kernel32.dll" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const ERR_PATHNOTFOUND              As Long = 76
Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31
Private Const MAX_PATH                      As Long = 260
Private Const UNLEN                         As Long = 256
Private Const UOI_FLAGS                     As Long = 1
Private Const WSF_VISIBLE                   As Long = 1
Private Const CHAR_EQUAL                    As Long = 61
Private Const CHAR_QUOTE                    As Long = 34
Private Const CHAR_SPACE                    As Long = 32
Private Const CHAR_UPPER_A                  As Long = 65

Public Enum SpecialFolder
   ApplicationData = &H1A
   CommonApplicationData = &H23
   CommonProgramFiles = &H2B
   Cookies = &H21
   Desktop = &H0
   DesktopDirectory = &H10
   Favorites = &H6
   History = &H22
   InternetCache = &H20
   LocalApplicationData = &H1C
   MyComputer = &H11
   MyMusic = &HD
   MyPictures = &H27
   Personal = &H5
   ProgramFiles = &H26
   Programs = &H2
   Recent = &H8
   SendTo = &H9
   StartMenu = &HB
   Startup = &H7
   SystemDir = &H25
   Templates = &H15
   Windows = &H24
End Enum


Private mBuffer(1023) As Byte
Private mPtrBuffer As Long
Private mWordBuffer As WordBuffer
Private mOSVersion As OperatingSystem


''
' Returns the command specified when executing the current application,
' including any arguments typed after the application name.
'
' @return The command line typed in at a command prompt.
'
Public Property Get CommandLine() As String
    CommandLine = SysAllocString(GetCommandLine)
End Property

''
' Returns the current directy of the executing application.
'
' @return The current directory.
'
Public Property Get CurrentDirectory() As String
    CurrentDirectory = CurDir
End Property

''
' Sets the current directory for the executing application.
'
' @param RHS The new directory to change to.
'
Public Property Let CurrentDirectory(ByVal RHS As String)
    If Len(RHS) = 0 Then _
        Throw Cor.NewArgumentException("Path cannot be an empty string.", "CurrentDirectory")
    
    On Error GoTo errTrap
    ChDir RHS
    Exit Property
    
errTrap:
    If Err.Number = ERR_PATHNOTFOUND Then
        Throw Cor.NewFileNotFoundException("The specified path was not found.", RHS)
    Else
        Throw Cor.NewArgumentException(Err.Description, "CurrentDirectory")
    End If
End Property

''
' Returns the name of the local computer.
'
' @return The name of the local computer.
'
Public Property Get MachineName() As String
    Dim size As Long
    
    size = MAX_COMPUTERNAME_LENGTH + 1
    GetComputerName mPtrBuffer, size
    MachineName = BufferToString(size)
End Property

''
' Returns string of a newline character for the local computer.
'
' @return The newline characters.
' @remarks Being this always going to be in VB, the return is just vbCrLf.
'
Public Property Get NewLine() As String
    NewLine = vbCrLf
End Property

''
' Returns an object representing the current operating system.
'
' @return The OperatingSystem object that represents the current OS.
'
Public Property Get OSVersion() As OperatingSystem
    If mOSVersion Is Nothing Then LoadOSVersion
    Set OSVersion = mOSVersion
End Property

''
' Returns the System directory for the current OS.
'
' @return The system directory.
' @remarks Generally this will be C:\Windows\System32 or C:\Winnt\System32 for NT4 system.
'
Public Property Get SystemDirectory() As String
    Dim size As Long
    
    size = GetSystemDirectory(mPtrBuffer, MAX_PATH + 1)
    If size > 0 Then SystemDirectory = BufferToString(size)
End Property

''
' Returns the number of milliseconds the system has been running since it started.
'
' @return Number of milliseconds since the system booted up.
' @remarks Once the maximum value is reached, it will wrap around to be negative,
' at which point 0 is returned. Wrapping will take place in approximatively 24.85 days.
'
Public Property Get TickCount() As Long
    TickCount = timeGetTime
    If TickCount < 0 Then TickCount = 0
End Property

''
' Returns the name of the local machine within a domain
'
' @return The computer name within a domain.
' @remarks Typically the domain name is the name of the machine, but
' can change depending on the application being deployed.
'
Public Property Get UserDomainName() As String
    Dim os As OperatingSystem
    Dim peUse As Long
    Dim size As Long
    Dim sid As String
    
    Set os = OSVersion
    If os.Platform <> Win32NT Then _
        Throw Cor.NewPlatformNotSupportedException("Must have Microsoft Windows NT 4.0 or later.")
    
    sid = String$(128, 0)
    size = 1024
    If LookUpAccountName(vbNullString, UserName, StrPtr(sid), 128, mPtrBuffer, size, peUse) Then
        UserDomainName = SysAllocString(mPtrBuffer)
    End If
End Property

''
' Determines if the current application is allowed to have user interaction.
'
' @return Indication if user interaction is allowed.
' @remarks This is used primarily for when the application is a service. It can
' be used to supress message dialogs and user entry when not allowed.
'
Public Property Get UserInteractive() As Boolean
    Dim hObj As Handle
    Dim info As UserObjectFlags
    
    If OSVersion.Platform = Win32NT Then
        hObj = GetProcessWindowStation
        If GetUserObjectInformation(hObj, UOI_FLAGS, info, LenB(info), 0&) Then
            UserInteractive = (info.dwFlags And WSF_VISIBLE)
        End If
    Else
        UserInteractive = True
    End If
End Property

''
' Returns the name of the currently logged in user of the local computer.
'
' @return The currently logged in user's name.
'
Public Property Get UserName() As String
    Dim size As Long
    
    size = UNLEN + 1
    GetUserName mPtrBuffer, size
    UserName = BufferToString(size - 1)
End Property

''
' Returns the amount of physical memory the current application is allocated.
'
' @return The amount of memory allocated to the application.
' @remarks This method only works for NT installations. 95, 98, and ME will
' receive a 0 when this method is called.
'
Public Property Get WorkingSet() As Long
    Dim h As Handle
    Dim info As Process_Memory_Counters
    
    If OSVersion.Platform <> Win32NT Then Exit Property
    
    h = GetCurrentProcess
    If GetProcessMemoryInfo(h, info, LenB(info)) Then
        WorkingSet = info.WorkingSetSize
    End If
End Property

''
' Replaces environment tags within a string with the actual values.
'
' @param Name A string containing environment variable tags to be replaced.
' @return A string containing actual environment variable values in place of tags.
' @remarks The environment variables are tagged using %.
' A string such as 'My OS is %OS%.' would become 'My OS is Windows_NT.' for NT machines.
'
Public Function ExpandEnvironmentVariables(ByVal Name As String) As String
    Const DEF_SIZE As Long = 128
    
    Dim buf As String
    Dim size As Long
    
    If Len(Name) = 0 Then Exit Function
    
    buf = SysAllocStringLen(0, DEF_SIZE)
    size = ExpandEnvironmentStrings(Name, buf, DEF_SIZE)
    
    If size > DEF_SIZE Then
        buf = SysAllocStringLen(0, size)
        size = ExpandEnvironmentStrings(Name, buf, size)
    End If
    
    ExpandEnvironmentVariables = left$(buf, size - 1)
End Function

''
' Returns an array of parsed arguments from the command line.
'
' @return An array of command line arguments, including the application name.
' @remarks The application is the first argument in the array. Arguments are
' separated by spaces. If spaces are embedded in an argument and should not
' be used to separate the argument, then the argument must be surrounded in quotes.
'
Public Function GetCommandLineArgs() As String()
    Dim s As String
    Dim buf As WordBuffer
    Dim pos As Long
    Dim tok As ArrayList
    Dim inQ As Boolean
    Dim sLen As Long
    Dim prevPos As Long
    
    Set tok = New ArrayList
    s = SysAllocString(GetCommandLine)
    sLen = Len(s)
    InitWordBuffer buf, StrPtr(s), Len(s)
    
    Do While pos < sLen
        Select Case buf.data(pos)
            Case CHAR_QUOTE
                If inQ Then
                    tok.Add Mid$(s, prevPos + 1, pos - prevPos)
                    inQ = False
                Else
                    inQ = True
                End If
                prevPos = pos + 1
            Case CHAR_SPACE
                If Not inQ Then
                    If buf.data(pos - 1) <> CHAR_SPACE Then
                        If pos - prevPos > 0 Then tok.Add Mid$(s, prevPos + 1, pos - prevPos)
                    End If
                    Do While buf.data(pos) = CHAR_SPACE
                        pos = pos + 1
                        If pos = sLen Then Exit Do
                    Loop
                    prevPos = pos
                    pos = pos - 1
                End If
        End Select
        pos = pos + 1
    Loop
    If buf.data(pos - 1) <> CHAR_SPACE And buf.data(pos - 1) <> CHAR_QUOTE Then
        tok.Add Mid$(s, prevPos + 1, pos - prevPos)
    End If
    Dim ret() As String
    ret = tok.ToArray(ciString)
    GetCommandLineArgs = ret
End Function

''
' Returns an environment variable value.
'
' @param Name The name of the environment variable to return.
' @return The value of the environment variable.
' @remarks If the variable does not exist, then an empty string is returned.
'
Public Function GetEnvironmentVariable(ByVal Name As String) As String
    Dim size As Long
    Dim buf As String
    
    size = 128
    buf = String$(size, 0)
    size = GetEnvironmentVariableA(Name, buf, size)
    If size > 128 Then
        buf = String$(size, 0)
        size = GetEnvironmentVariableA(Name, buf, size)
    End If
    If size > 0 Then
        GetEnvironmentVariable = left$(buf, size)
    End If
End Function

''
' Returns an IDictionary object containing all environment variables as name/value pairs.
'
' @return A dictionary containing the name/value pairs.
' @remarks The hidden environment variables are not included in the dictionary.
'
Public Function GetEnvironmentVariables() As IDictionary
    Dim Chars() As Byte
    Dim SA As SafeArray1d
    Dim ret As Hashtable
    Dim isHidden As Boolean
    Dim pos As Long
    Dim startPos As Long
    Dim Name As String
    Dim value As String
    
    With SA
        .cbElements = 1
        .cDims = 1
        .pvData = GetEnvironmentStrings
        .cElements = &H7FFFFFFF
    End With
    SAPtr(Chars) = VarPtr(SA)
    
    Set ret = New Hashtable
    
    Do While Chars(pos) <> 0
        startPos = pos
        Do
            If Chars(pos) = CHAR_EQUAL Then
                If startPos < pos Then Exit Do
                isHidden = True
            End If
            pos = pos + 1
        Loop
        
        If Not isHidden Then Name = BytesToString(Chars, startPos, pos - startPos)
        
        startPos = pos + 1
        Do While Chars(pos) <> 0: pos = pos + 1: Loop
        
        If Not isHidden Then
            value = BytesToString(Chars, startPos, pos - startPos)
            ret.Add Name, value
        End If
        pos = pos + 1
        isHidden = False
    Loop
    
    FreeEnvironmentStrings SA.pvData
    SAPtr(Chars) = 0
    Set GetEnvironmentVariables = ret
End Function

''
' Returns the path for a specified folder type.
'
' @param Folder The folder to return the path of.
' @return The path for the specified folder.
'
Public Function GetFolderPath(ByVal Folder As SpecialFolder) As String
    SHGetFolderPath 0, Folder, 0, 0, mPtrBuffer
    GetFolderPath = BufferToString(lstrlen(mPtrBuffer))
End Function

''
' Returns the logical drives installed on the local machine.
'
' @return An array containing all of the logical drives installed.
'
Public Function GetLogicalDrives() As String()
    GetLogicalDrives = Directory.GetLogicalDrives
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetResourceString(ByVal resId As Long, ParamArray args() As Variant) As String
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16)
    GetResourceString = cString.FormatArray(LoadResString(resId), vArgs)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BufferToString(ByVal size As Long) As String
    BufferToString = BytesToString(mBuffer, 0, size)
End Function

Private Function BytesToString(ByRef Bytes() As Byte, ByVal Index As Long, ByVal size As Long) As String
    Dim ret As String
    Dim i As Long
    
    If size = 0 Then Exit Function
    
    ret = SysAllocStringLen(0, size)
    mWordBuffer.SA.pvData = StrPtr(ret)
    For i = 0 To size - 1
        mWordBuffer.data(i) = Bytes(Index + i)
    Next i
    BytesToString = ret
End Function

Private Sub LoadOSVersion()
    Dim info As OsVersionInfo
    
    info.dwOSVersionInfoSize = LenB(info)
    GetVersionEx info
    Set mOSVersion = Cor.NewOperatingSystem(info.dwPlatformId, Cor.NewVersion(info.dwMajorVersion, info.dwMinorVersion))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mPtrBuffer = VarPtr(mBuffer(0))
    InitWordBuffer mWordBuffer, 0, &H7FFFFFFF
End Sub
