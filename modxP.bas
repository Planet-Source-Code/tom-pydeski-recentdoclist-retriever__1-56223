Attribute VB_Name = "modXP"
Option Explicit
'usage
'in a new module place the following
'Sub Main()
'XPMain
'Load frmMain
'frmMain.Show
'End Sub
'
'Hello, The above code will make sure you declare everything.
'This module will allow you to run an ordinary VB6 application (on WinXP) using the cool Windows XP Style Controls.
'I used bits from a few various projects downloaded from PSC. I just put them all together to
'create a 100% working XP style Control Module.
'You can however give me credit for resolving the timing issue with running only one
'instance of the application. (I made sure the program will always run - whether it's run for the
'first time and not...)
'NOTES:
'1. This only works when running the compiled .EXE file.
'2. This only works with WINDOWS XP - Don't try with an older version: IT WON'T WORK!
'3. A "App.EXEname.exe.MANIFEST" file must be created in the same directory as the EXE.
'        EG. 'WinXP.exe.MANIFEST' in the same directory as 'WinXP.exe'
'        (This module will do it automatically)
'4. A MANIFEST file is a .NET FrameWork file - In other words: We are 'instructing' WinXP into
'        changing the old VB6 generated controls into the new look XP style
'5. If there wan't a manifest file to begin with - and you only just created one when a user runs
'        your program for the first time - YOU NEED TO RESTART your application. (This module
'        will do this automatically)
'6. If you only want one instance of your program to run - There WAS a timing issue which i sorted
'        out using the registry.
'7. You will notice i have code which detects the OS - This isn't necessary as creating a MANIFEST
'        file in older versions and then trying to apply it - WON'T effect/apply anything.
'8. All this is done automatically if you use VB.NET.
'9. It's a Sunday evening - hence please tolerate any spelling mistakes. :)
'Enjoy!
'David Sykes
'dsykes@ mighty.co.za
'#####################################################################################
'#
'#      This code will apply WinXP style controls into your program
'#
'#      NB: Only works in the compiled .EXE file
'#
'#      By: David Sykes
'#
'#####################################################################################
'Function to get OS
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'Function to Start a external program/file
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
'Function to apply WinXP style controls
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEX_TYPE) As Long
'For XP style controls
Public Const ICC_INTERNET_CLASSES = &H800
'For OS detection
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
'For 'Start a external program/file'
Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum
'For OS detection
Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long           'NT: Build Number, 9x: High-Order has Major/Minor ver, Low-Order has build
    PlatformID As Long
    szCSDVersion As String * 128    'NT: ie- "Service Pack 3", 9x: 'arbitrary additional information'
End Type
'For OS detection
Public Enum cnWin32Ver
    UnknownOS = 0
    Windows95 = 1
    Windows98 = 2
    WindowsME = 3
    WindowsNT4 = 4
    Windows2k = 5
    WindowsXP = 6
End Enum
'For XP Style controls
Public Type INITCOMMONCONTROLSEX_TYPE
    dwSize As Long
    dwICC As Long
End Type

Public Sub XPMain()
'THE PROGRAM STARTS HERE: (If you don't want to start your program in the Sub Main
'module then copy the below code in Sub Main into your Form_Load of the first form
'that loads.
'/OR/
'Call Main from the Form_Load of the first form that loads ("Call Main")
'NB: It's good coding practice to start a program in Sub Main
'HINT: To start a program in Sub Main then goto:
' Click "Project"             in "File, Edit, View menu"
' Click "Project" Properties
' Change "StartUp object" to "Sub Main"
Dim comctls As INITCOMMONCONTROLSEX_TYPE
Dim retval As Long
Dim CanProceed As Boolean
CanProceed = IsManifestFile 'CanProceed = True if MANIFEST FILE EXISTS
'Hence - No need to auto restart the application
If Val(Win32Ver) > 5 Then 'IF WINDOWS XP
    If MakeMANIFESTfile Then 'IF MANIFEST FILE WAS SUCCESSFULLY CREATED
        'Apply XP Style Controls
        With comctls
            .dwSize = Len(comctls)
            .dwICC = ICC_INTERNET_CLASSES
        End With
        retval = InitCommonControlsEx(comctls)
        '#######################
    Else
        'Manifest file -->wasn't<--- successfully created
        'Hence - can't apply xp style controls
        'Hence - No need to auto restart the application
        CanProceed = True
    End If
Else
    'Not WindowsXP
    'Hence - can't apply xp style controls
    'Hence - No need to auto restart the application
    CanProceed = True
End If
If CanProceed Then   'Can continue with this exe session.
    'Load FIRST FORM
    'Load Form1
    'Show FIRST FORM
    'Form1.Show
Else
    'The application needs to be auto restarted
    'USE THIS CODE IF YOU WANT ONE INSTANCE OF YOUR PROGRAM TO RUN:
    SaveSetting App.EXEName, "Settings", "CanRun", "YES"
    '###################################################
    'START THE EXE FILE AGAIN: (Shelldocument = True if program was successfully
    'restarted // False if it wasn't
    If ShellDocument(App.Path & "\" & App.EXEName & ".exe", , , , START_NORMAL) Then
        End ' End this program --> New one auto started
    Else 'Wasn't able to start exe file (Proceed in current exe session)
        'USE THIS CODE IF YOU WANT ONE INSTANCE OF YOUR PROGRAM TO RUN:
        SaveSetting App.EXEName, "Settings", "CanRun", "NO"
        '##################################################
        'Load FIRST FORM
    '    Load Form1
        'SHOW FIRST FORM
    '    Form1.Show
    End If
End If
End Sub

Public Property Get MakeMANIFESTfile() As Boolean
'THE BELOW CODE CREATES A MANIFEST FILE:
'MakeMANIFESTfile returns True if it was able to create the file
MakeMANIFESTfile = False
On Local Error GoTo MakeMANIFESTfile_Err
Dim ManifestFileName As String
Dim NewFreeFile As Integer
ManifestFileName = App.Path & "\" & App.EXEName & ".exe.MANIFEST"
NewFreeFile = FreeFile
'Note:  CHR(34)   =   "
Open ManifestFileName For Output As NewFreeFile
Print #NewFreeFile, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
Print #NewFreeFile, "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">"
Print #NewFreeFile, "<assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "x86" & Chr(34) & " name=" & Chr(34) & "prjThemed" & Chr(34) & " type=" & Chr(34) & "Win32" & Chr(34) & " />"
Print #NewFreeFile, "<dependency>"
Print #NewFreeFile, "<dependentAssembly>"
Print #NewFreeFile, "<assemblyIdentity type=" & Chr(34) & "Win32" & Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "x86" & Chr(34) & " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & " language=" & Chr(34) & "*" & Chr(34) & " />"
Print #NewFreeFile, "</dependentAssembly>"
Print #NewFreeFile, "</dependency>"
Print #NewFreeFile, "</assembly>"
Close NewFreeFile
MakeMANIFESTfile = True 'FILE CREATED OK
Exit Property
MakeMANIFESTfile_Err:
MakeMANIFESTfile = False 'ERROR CREATING FILE
End Property

Public Property Get IsManifestFile() As Boolean
'THE BELOW CODE CHECK THE EXISTANCE OF THE MANIFEST FILE:
'IF FILE EXISTS THEN IsManifestFile returns True
'THIS CODE SIMPLY TRIES TO OPEN THE FILE - IF THERE IS AN ERROR
'THEN THE FILE DOESN'T EXIST
IsManifestFile = False
On Local Error GoTo IsManifestFile_Err
Dim ManifestFileName As String
Dim NewFreeFile As Integer
ManifestFileName = App.Path & "\" & App.EXEName & ".exe.MANIFEST"
NewFreeFile = FreeFile
Open ManifestFileName For Input Access Read As NewFreeFile
Close NewFreeFile
IsManifestFile = True 'FILE DOES EXIST
Exit Property
IsManifestFile_Err:
IsManifestFile = False 'FILE DOESN'T EXIST
End Property

Public Function ShellDocument(sDocName As String, Optional ByVal Action As String = "Open", Optional ByVal Parameters As String = vbNullString, Optional ByVal Directory As String = vbNullString, Optional ByVal WindowState As StartWindowState) As Boolean
'THIS CODE IS TO RUN AN EXTERNAL APPLICATION / DOCUMENT:
'SHELLDOCUMENT will return a True if the file was run
'Else False if not
Dim Response
Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
Select Case Response
    Case Is < 33
        ShellDocument = False
    Case Else
        ShellDocument = True
End Select
End Function

Public Function Win32Ver() As cnWin32Ver
'THE BELOW CODE IS TO DETECT OPERATING SYSTEM
'WE USE THIS TO SEE IF THE USER IS RUNNING Windows XP
'# Public subs/functions
'# Returns the asso. cnWin32Ver eNum value of the current Win32 OS
Dim oOSV As OSVERSIONINFO
oOSV.OSVSize = Len(oOSV)
'If the API returned a valid value
If GetVersionEx(oOSV) = 1 Then
    'If we're running WindowsXP
    '   If VER_PLATFORM_WIN32_NT, dwVerMajor is 5 and dwVerMinor is 1, it's WindowsXP
    If (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1) Then
        Win32Ver = WindowsXP
        'If we're running WindowsNT2000 (NT5)
        '   If VER_PLATFORM_WIN32_NT, dwVerMajor is 5 and dwVerMinor is 0, it's Windows2k
    ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0) Then
        Win32Ver = Windows2k
        'If we're running WindowsNT4
        '   If VER_PLATFORM_WIN32_NT and dwVerMajor is 4
    ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 4) Then
        Win32Ver = WindowsNT4
        'If we're running Windows ME
        '   If VER_PLATFORM_WIN32_WINDOWS and
        '   dwVerMajor = 4,  and dwVerMinor > 0, return true
    ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 90) Then
        Win32Ver = WindowsME
        'If we're running Windows98
        '   If VER_PLATFORM_WIN32_WINDOWS and
        '   dwVerMajor => 4, or dwVerMajor = 4 and
        '   dwVerMinor > 0, return true
    ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (oOSV.dwVerMajor > 4) Or (oOSV.dwVerMajor = 4 And oOSV.dwVerMinor > 0) Then
        Win32Ver = Windows98
        'If we're running Windows95
        '   If VER_PLATFORM_WIN32_WINDOWS and
        '   dwVerMajor = 4, and dwVerMinor = 0,
    ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 0) Then
        Win32Ver = Windows95
        'Else the OS is not reconized by this function
    Else
        Win32Ver = UnknownOS
    End If
    'Else the OS is not reconized by this function
Else
    Win32Ver = UnknownOS
End If
End Function

Public Function isNT() As Boolean
'Returns true if the OS is WindowsNT4, Windows2k or WindowsXP
'Determine the return value of Win32Ver() and set the return value accordingly
Select Case Win32Ver()
    Case WindowsNT4, Windows2k, WindowsXP
        isNT = True
    Case Else
        isNT = False
End Select
End Function

Public Function is9x() As Boolean
'Returns true if the OS is Windows95, Windows98 or WindowsME
'Determine the return value of Win32Ver() and set the return value accordingly
Select Case Win32Ver()
    Case Windows95, Windows98, WindowsME
        is9x = True
    Case Else
        is9x = False
End Select
End Function

Public Function isWinXP() As Boolean
'Returns true if the OS is WindowsXP
Dim oOSV As OSVERSIONINFO
oOSV.OSVSize = Len(oOSV)
'If the API returned a valid value
If (GetVersionEx(oOSV) = 1) Then
    isWinXP = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1)
End If
End Function

Public Function isWin2k() As Boolean
'Returns true if the OS is Windows2k
Dim oOSV As OSVERSIONINFO
oOSV.OSVSize = Len(oOSV)
'If the API returned a valid value
If (GetVersionEx(oOSV) = 1) Then
    isWin2k = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0)
End If
End Function

Public Function isWinNT4() As Boolean
'Returns true if the OS is WindowsNT4
Dim oOSV As OSVERSIONINFO
oOSV.OSVSize = Len(oOSV)
'If the API returned a valid value
If (GetVersionEx(oOSV) = 1) Then
    isWinNT4 = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 4)
End If
End Function

Public Function isWinME() As Boolean
'Returns true if the OS is WindowsME
Dim oOSV As OSVERSIONINFO
oOSV.OSVSize = Len(oOSV)
'If the API returned a valid value
If (GetVersionEx(oOSV) = 1) Then
    isWinME = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 90)
End If
End Function

Public Function isWin98() As Boolean
'# Returns true if the OS is Windows98
Dim oOSV As OSVERSIONINFO
oOSV.OSVSize = Len(oOSV)
'If the API returned a valid value
If (GetVersionEx(oOSV) = 1) Then
    isWin98 = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (oOSV.dwVerMajor > 4) Or (oOSV.dwVerMajor = 4 And oOSV.dwVerMinor > 0)
End If
End Function

Public Function isWin95() As Boolean
'# Returns true if the OS is Windows95
Dim oOSV As OSVERSIONINFO
oOSV.OSVSize = Len(oOSV)
'If the API returned a valid value
If (GetVersionEx(oOSV) = 1) Then
    isWin95 = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 0)
End If
End Function
