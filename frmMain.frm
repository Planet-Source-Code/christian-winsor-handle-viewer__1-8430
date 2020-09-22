VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Process Viewer"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   1635
   End
   Begin MSComctlLib.TreeView tvProc 
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5424
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "Close Process"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'getwindow is used to get the relative window specified by the flag
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'getwindowtext gets the window title
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'getwindowtextlength will get the length of the window title, this is used to create a null string to store the window title in
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
'post message is used to to send windows messages, used to send the wm_close message to close a process
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'the windows message to tell a window to close
Const WM_CLOSE = &H10

'store the currently clicked node
Public CurrentNode As MSComctlLib.Node
'store the currently clicked node's index
Public CurrentIndex

'get window constants, gets windows relative to the hwnd specified
Const GW_CHILD = 5
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_OWNER = 4


Private Sub cmdRefresh_Click()
'clear the treeview
tvProc.Nodes.Clear
'get the processes and populate the treeview
GetProcess
End Sub

Private Sub Form_Resize()
'resize the treeview and position the button
tvProc.Width = frmMain.ScaleWidth
tvProc.Height = frmMain.ScaleHeight - cmdRefresh.Height - 60
cmdRefresh.Top = tvProc.Top + tvProc.Height + 30
End Sub

Private Sub GetProcess()
Dim sTitle As String
Dim myHwnd
'get the owner of my window hwnd
myHwnd = GetWindow(Me.hwnd, GW_OWNER)
'go into a loop to get all the processes until there are none left
Do While myHwnd <> 0
    DoEvents
    'get the window title
    sTitle = GetWinCaption(myHwnd)
    'if the title isn't empty then add it to the treeview _
    if this if statement wasn't here you would get many more processes but _
    most without titles
    If sTitle <> "" Then
        'add the item to the treeview
        tvProc.Nodes.Add , , "a" & Str(myHwnd), Str(myHwnd) & " : " & sTitle
            'get all the child processes of the current window
            GetChildren myHwnd, tvProc.Nodes.Count
    End If
    'get the next window
    myHwnd = GetWindow(myHwnd, GW_HWNDNEXT)
Loop
End Sub


Public Function GetWinCaption(hwnd) As String
Dim iTextLength As Integer
Dim sTitle As String
'get the window title length so that we can make an empty string
iTextLength = GetWindowTextLength(hwnd)
'make an empty string the length of the window title
sTitle = String(iTextLength, 0)
'get the window title and store in sTitle
GetWindowText hwnd, sTitle, iTextLength + 1
'return the value
GetWinCaption = sTitle
End Function

Private Sub GetChildren(hwnd, Index)
Dim myHwnd
Dim sTitle As String
'working the same as the getprocess sub, we get the child process
myHwnd = GetWindow(hwnd, GW_CHILD)
'loop until there are no more processes
Do While myHwnd <> 0
    'get the title of the process
    sTitle = GetWinCaption(myHwnd)
    'if the title isn't empty then add as a child node in the treeview _
    if this if statement wasn't here you would get many more processes but _
    most without titles
    If sTitle <> "" Then
        DoEvents
        'add the child process as a child node in the treeview
        tvProc.Nodes.Add "a" & Str(hwnd), tvwChild, "a" & Str(myHwnd), Str(myHwnd) & " : " & sTitle
    End If
    'get the next child window
    myHwnd = GetWindow(myHwnd, GW_HWNDNEXT)
Loop
End Sub

Private Sub mnuClose_Click()
Dim lReturn As Long
'send the wm_close message to close the window
lReturn = PostMessage(Val(Right(CurrentNode.Key, Len(CurrentNode.Key) - 1)), WM_CLOSE, 0, 0)
If lReturn <> 0 Then
    'remove the item closed from the treeview
    tvProc.Nodes.Remove CurrentIndex
End If
End Sub

Private Sub tvProc_NodeClick(ByVal Node As MSComctlLib.Node)
'store the current node clicked on so we can use it in mnuClose_Click
Set CurrentNode = Node
'store the current node index clicked on so we can use it in mnuClose_Click
CurrentIndex = Node.Index
'popup the menu, menu currently only has a close item
PopupMenu mnuFile
End Sub

'CW 2000
