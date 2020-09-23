VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RecentDocs 
   Caption         =   "Tom Pydeski's RecentDocList Retriever"
   ClientHeight    =   7560
   ClientLeft      =   1530
   ClientTop       =   1605
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RecentDocList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7560
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RecentDocList.frx":12D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RecentDocList.frx":15EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   3480
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   50
      Width           =   6255
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6690
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   7140
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Press Delete to Clear entire selected Node"
      Top             =   45
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   12594
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mFind 
         Caption         =   "&Find in Current Key"
         Shortcut        =   ^F
      End
      Begin VB.Menu mFindAll 
         Caption         =   "Find in &All Keys"
      End
   End
End
Attribute VB_Name = "RecentDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:Tom Pydeski
'BitWise Industrial Automation, Inc.
'
'this can also be added to this program
'Run Dialog Recent Menu
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU
'HKEY_USERS\S-1-5-21-127730482-1884467411-3661661970-1007\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU
'typed url list
'HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\TypedURLs
'HKEY_USERS\S-1-5-21-127730482-1884467411-3661661970-1007\Software\Microsoft\Internet Explorer\TypedURLs
'then a listing...url1...url2...etc
'
'This Program will read the Recent Doc List located in the registry at:
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs
'This data is stored in a binary format, so we have to read the binary data into
'a byte array and build it into strings.
'As with all of my submissions, I have utilized code found on PSC and elsewhere
'for various functions, but the rest was written by me.
'Special thanks to Kegham, whose Winstartup 2004 project had some valuable code
'for enumerating and walking through registry keys and for some treeview pointers
'and to MrBoBo who also had some very useful code for the registry
'also to David Sykes for his XP style module that i have implemented in all
'of my projects for that XP Look
'once the key values are loaded into the list, pressing delete will delete the selected
'entry from the registry.
'
'Disclaimer:
'THIS PROGRAM ACCESSES AND MODIFIES ENTRIES IN THE REGISTRY!
'I tested it only on my machine, which is windows XP service pack 2
'I am not responsible for any bad things that may happen due to the
'use of this program
'
'As with all software using the registry
'BACKUP your registry before using
'This ran fine on my machine and the only thing it deletes are the
'binary entries for the recent doc list
'
Option Explicit
' This API function allows us to change the parent of any component that has a hWnd
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'usage SetParent Check1.hwnd, Command1.hwnd
Dim MaxFiles
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Declare the API function call.
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam$) As Long
' Add API constant
Const LB_ITEMFROMPOINT = &H1A9
Const LB_SETTOPINDEX = &H197
Const LB_FINDSTRING = &H18F
Const LB_SELITEMRANGEEX = &H183
Dim Findstr As String
Dim FoundPos As Long
Dim FoundLine As Long
Dim fStart As Integer
Dim Replstr  As String
Dim Inits As Byte
Dim LogFile$
Dim LogFileOut$
Dim i As Integer
Dim OldIndex As Integer
Dim hKey As Long
Dim lRetVal As Long
Dim Indy As Integer
Dim SelectedKeyNum As Integer
Dim Confirm As Long

Private Sub Form_Resize()
If WindowState = vbMinimized Then Exit Sub
'maximize the height and width to fit the screen
TV1.Height = (Me.Height - TV1.Top) - 850
List1.Height = (Me.Height - List1.Top) - 850
List1.Width = (Me.Width - List1.Left) - 200
List2.Top = List1.Top
List2.Height = List1.Height
List2.Width = List1.Width
Text1.Width = List1.Width
End Sub

Private Sub mExit_Click()
Unload Me
Set RecentDocs = Nothing
End
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    DeleteFromList True
End If
End Sub

Private Sub list1_DblClick()
OldIndex = List1.ListIndex
If OldIndex = -1 Then
End If
Text1.Text = List1.List(OldIndex)
cont:
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lParam&, Result&
Indy = List1.ListIndex
Beeep
Text1.Text = List1.List(Indy)
Exit Sub
'Result& = SendMessage(List1.hwnd, LB_SETTOPINDEX, INDY, lParam&)
List1.Visible = False
List1.TopIndex = Indy
List1.Visible = True
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' present related tip message
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long
'
If Button = 0 Then ' if no button was pressed
    lXPoint = CLng(X / Screen.TwipsPerPixelX)
    lYPoint = CLng(Y / Screen.TwipsPerPixelY)
    '
    With List1
        ' get selected item from list
        lIndex = SendMessage(.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
        ' show tip or clear last one
        If (lIndex >= 0) And (lIndex <= .ListCount) Then
            .ToolTipText = .List(lIndex) & " " & .ItemData(lIndex)
        Else
            .ToolTipText = ""
        End If
    End With '(List1)
End If '(button=0)
End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
OldIndex = List2.ListIndex
If KeyCode = vbKeyDelete Then
    'open registry key
    List1_KeyDown vbKeyDelete, 0
    List2.RemoveItem (List2.ListIndex)
    Refresh
    DoEvents
    List2.ListIndex = OldIndex
    Refresh
    DoEvents
    List2.SetFocus
End If
End Sub

Private Sub List2_DblClick()
OldIndex = List2.ListIndex
If OldIndex = -1 Then
End If
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List2_DblClick
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lParam&, Result&
Indy = List2.ListIndex
Beeep
Text1.Text = List2.List(Indy)
Exit Sub
'Result& = SendMessage(List2.hwnd, LB_SETTOPINDEX, INDY, lParam&)
List2.Visible = False
List2.TopIndex = Indy
List2.Visible = True
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' present related tip message
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long
'
If Button = 0 Then ' if no button was pressed
    lXPoint = CLng(X / Screen.TwipsPerPixelX)
    lYPoint = CLng(Y / Screen.TwipsPerPixelY)
    '
    With List2
        ' get selected item from list
        lIndex = SendMessage(.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
        ' show tip or clear last one
        If (lIndex >= 0) And (lIndex <= .ListCount) Then
            .ToolTipText = .List(lIndex)
        Else
            .ToolTipText = ""
        End If
    End With '(List2)
End If '(button=0)
End Sub

Private Sub Exit_Click()
Timer1.Enabled = False
Unload Me
'CloseMutEx
End
End Sub

Private Sub Form_Load()
On Error GoTo Oops
AppDir = App.Path
RecMax = 0
AppName = App.EXEName
ChDir App.Path
TV1.Nodes.Clear 'clear tv1's of any previous nodes
'add root node
TV1.Nodes.Add , , "Root", "RecentDocs", 1, 2
'
'remove items from select
List1.Clear
List2.Height = List1.Height
'
getfile:
GetRootBinary
PopList (0)
'populate tree view with the subkeys
For i = 1 To UBound(SubKeyName)
    TV1.Nodes.Add "Root", tvwChild, "Key" & i, SubKeyName(i), 1, 2
Next i
'now add for the following
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU
'HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\TypedURLs
TV1.Nodes.Add "Root", tvwNext, "RunMRU", "RunMRU", 1, 2
TV1.Nodes.Add "Root", tvwNext, "TypedURLs", "TypedURLs", 1, 2
'
RecentDocs.Refresh
GoTo Exit_Form_Load
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Form_Load "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Form_Load"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Close
Exit_Form_Load:
Beeep
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> 1 Then
    Exit_Click
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    List2.Visible = False
    List1.Visible = True
    Text1.Text = ""
End If
End Sub

Private Sub mfind_Click()
Dim fMess$
Screen.MousePointer = 11
Dim FindIn As String
fStart = 0
FindIn = InputBox("Enter the string to find", "Find in THIS key...", Findstr)
Findstr = FindIn
If Findstr = "" Then GoTo nofind
' Find the text specified in the listbox control.
'i could use the api call, but it works on a matchcase basis
'so i would rather do it the old fashioned way
FoundPos = -1
Screen.MousePointer = 11
For i = 0 To List1.ListCount - 1
    If InStr(1, List1.List(i), Findstr, vbTextCompare) > 0 Then
        List1.TopIndex = i
        Text1.Text = List1.List(i)
        FoundPos = i
        'Exit For
    End If
Next i
If FoundPos <> -1 Then
    ' Returns number of line containing found text.
    Beeep
    fMess$ = "Found in Value " & CStr(FoundPos)
    GoTo nofind
End If
Alarm
MsgBox Findstr & " not found!"
nofind:
Screen.MousePointer = 0
End Sub

Private Sub mFindAll_Click()
'finds a string in all of the subkeys within the recentdocs key
Dim FindIn As String
Dim fMess$
Dim j As Integer
Screen.MousePointer = 11
'first lets read all of the trees
For j = 1 To RecentSubKeys
    sKeyName = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs\" & SubKeyName(j)
    GetBinary sKeyName, j
Next j
fStart = 0
FindIn = InputBox("Enter the string to find", "Find in ALL keys...", Findstr)
Findstr = FindIn
If Findstr = "" Then GoTo nofind
'Find the text specified in the listbox control.
FoundPos = -1
Screen.MousePointer = 11
For i = 0 To RecentSubKeys
    For j = 1 To RecentMax(0)
        'check the value array
        If InStr(1, RegValue(i, j), Findstr, vbTextCompare) > 0 Then
            'lets force our tree to that node
            TV1.Nodes(i + 1).Selected = True
            TV1_NodeClick TV1.Nodes(i + 1)
            DoEvents
            Refresh
            'lists start at 0 so let's subtract 1
            List1.TopIndex = j - 1
            Text1.Text = List1.List(j - 1)
            FoundPos = i
            Exit For
        End If
    Next j
    If FoundPos >= 0 Then Exit For
Next i
If FoundPos <> -1 Then
    ' Returns number of line containing found text.
    Beeep
    fMess$ = "Found in Value " & CStr(FoundPos)
    GoTo nofind
End If
Alarm
MsgBox Findstr & " not found!"
nofind:
Screen.MousePointer = 0
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim Result&, wParam&, s$
'text1 is the filter for displaying all values that match into list 2 from list1
Dim sStr As String
Dim chkFile$
Dim chks As Integer
If KeyCode = vbKeyDown Then
    If List2.ListIndex < List2.ListCount - 1 Then
        List2.ListIndex = List2.ListIndex + 1
    End If
    On Error Resume Next
    List2.SetFocus
    Exit Sub
End If
DoEvents
Refresh
sStr = UCase(Text1.Text)
If Len(sStr) = 0 Then Exit Sub
List2.Clear
List1.Visible = False
List2.Visible = False
For i = List1.ListCount - 1 To 0 Step -1
    chkFile$ = UCase(List1.List(i))
    chks = InStr(chkFile$, sStr)
    If sStr = "" Then chks = 1
    If chks > 0 Then
        'List1.Visible = False
        'Result& = SendMessage(List1.hwnd, LB_SETTOPINDEX, i, lParam&)
        List1.ListIndex = i
        List1.TopIndex = i
        'List1.Refresh
        List2.AddItem List1.List(i), 0
        List2.ItemData(0) = List1.ItemData(i)
        'Exit For
    End If
Next i
If List2.ListCount > 0 Then List2.ListIndex = 0
'List1.Visible = True
List2.Visible = True
If KeyCode = 13 Then
    sStr = UCase(Text1.Text)
    chkFile$ = UCase(List1.List(List1.ListIndex))
    chks = InStr(chkFile$, sStr)
    If chks > 0 Then
        list1_DblClick
    Else
        
    End If
Else
End If
'
'wParam& = -1
's$ = Text1.Text
'Result& = SendMessageByString(List1.hwnd, LB_FINDSTRING, wParam&, s$)
'List1.ListIndex = Result&
End Sub

Private Sub TV1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Confirm = MsgBox("Are you sure you want to delete all entries for " & TV1.SelectedItem.Text & "?", vbOKCancel + vbQuestion + vbMsgBoxSetForeground, "Confirm Deletion of Registry Values")
    If Confirm = vbCancel Then Exit Sub
    '
    'delete all listings from the bottom up
    For i = List1.ListCount - 1 To 0 Step -1
        List1.ListIndex = i
        If List1.Text <> AddSpace(Str$(i), 3) & vbTab Then
            DeleteFromList False
        
        End If
    Next i
End If
End Sub

Private Sub TV1_NodeClick(ByVal Node As MSComctlLib.Node)
Debug.Print Node.Key
List1.Visible = True
List2.Visible = False
SelectedKeyNum = Node.Index - 1
Debug.Print Node.Key
If Node.Key = "Root" Then
    'Now lets populate the list with the new stuff
    sKeyName = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs\"
    GetBinary sKeyName, 0
    PopList 0
ElseIf Node.Key = "RunMRU" Then
    'Now lets populate the list with the new stuff
    sKeyName = "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    GetRegText sKeyName, "", 1
    PopList RecentSubKeys + 1
ElseIf Node.Key = "TypedURLs" Then
    'Now lets populate the list with the new stuff
    sKeyName = "Software\Microsoft\Internet Explorer\TypedURLs"
    GetRegText sKeyName, "url", 2
    PopList RecentSubKeys + 2
Else
    'Now lets populate the list with the new stuff for the selected key.
    'we could alternatively do this at startup, but why get it before you need it.
    sKeyName = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs\" & TV1.Nodes(TV1.SelectedItem.Index).Text
    GetBinary sKeyName, SelectedKeyNum
    PopList SelectedKeyNum
    Debug.Print "Opening "; SelectedKeyNum; " = "; TV1.Nodes(TV1.SelectedItem.Index).Text
End If
End Sub

Sub PopList(KeyNum As Integer)
'populate list with the string values for the selected subkey
With List1
    .Visible = False
    .Clear
    For i = 0 To RecentMax(KeyNum)
       .AddItem AddSpace(Str$(i), 3) & vbTab & RegValue(KeyNum, i)
       .ItemData(.ListCount - 1) = i
    Next i
    .Visible = True
End With
End Sub

Sub DeleteFromList(Optional Confirmation As Boolean)
Dim KeyRoot$
OldIndex = List1.ListIndex
'confirm the delete
If Confirmation = True Then
    Confirm = MsgBox("Are you sure you want to delete" & vbCrLf & RegValue(SelectedKeyNum, OldIndex) & "?", vbOKCancel + vbQuestion + vbMsgBoxSetForeground, "Confirm Deletion of Registry Value")
    If Confirm = vbCancel Then Exit Sub
End If
'set the registry keyname based on which key is open
KeyRoot$ = ""
If TV1.SelectedItem.Text = "RunMRU" Then
    sKeyName = "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    'open registry key
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    If lRetVal <> 0 Then
        MsgBox "Failed to Open " & sKeyName
    End If
    'delete the binary value from the registry
    lRetVal = DeleteValue(HKEY_CURRENT_USER, sKeyName, Chr$(65 + List1.ItemData(List1.ListIndex)))
    If lRetVal <> 0 Then
        MsgBox "Failed to Delete " & sKeyName
    End If
    'now let's re-load the selected branch
    GetRegText sKeyName, "", 1
    PopList RecentSubKeys + 1
ElseIf TV1.SelectedItem.Text = "TypedURLs" Then
    sKeyName = "Software\Microsoft\Internet Explorer\TypedURLs"
    KeyRoot$ = "url"
    'open registry key
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    If lRetVal <> 0 Then
        MsgBox "Failed to Open " & sKeyName
    End If
    'delete the binary value from the registry
    lRetVal = DeleteValue(HKEY_CURRENT_USER, sKeyName, KeyRoot$ & List1.ItemData(List1.ListIndex))
    If lRetVal <> 0 Then
        MsgBox "Failed to Delete " & sKeyName
    End If
    'now let's re-load the selected branch
    GetRegText sKeyName, "url", 2
    PopList RecentSubKeys + 2
Else
    sKeyName = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs\" & SubKeyName(SelectedKeyNum)
    'open registry key
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    If lRetVal <> 0 Then
        MsgBox "Failed to Open " & sKeyName
    End If
    'delete the binary value from the registry
    lRetVal = DeleteValue(HKEY_CURRENT_USER, sKeyName, KeyRoot$ & List1.ItemData(List1.ListIndex))
    If lRetVal <> 0 Then
        MsgBox "Failed to Delete " & sKeyName
    End If
    'now let's re-load the selected branch
    GetBinary sKeyName, SelectedKeyNum
    PopList (SelectedKeyNum)
End If
Refresh
DoEvents
On Error Resume Next
List1.ListIndex = OldIndex
Refresh
DoEvents
On Error Resume Next
List1.SetFocus
End Sub

