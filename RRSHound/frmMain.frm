VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D28F8786-0BB9-402B-92DC-F32DE23A324E}#3.0#0"; "OutlookBar.ocx"
Begin VB.Form frmMain 
   Caption         =   "RSSHound"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRight 
      Height          =   5175
      Left            =   2100
      ScaleHeight     =   5115
      ScaleWidth      =   6915
      TabIndex        =   3
      Top             =   420
      Width           =   6975
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5040
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0634
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":094E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox PicSplitH 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         Height          =   4800
         Left            =   4200
         ScaleHeight     =   2090.126
         ScaleMode       =   0  'User
         ScaleWidth      =   780
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   72
      End
      Begin VB.PictureBox picBrowser 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   60
         ScaleHeight     =   2055
         ScaleWidth      =   9585
         TabIndex        =   4
         Top             =   2280
         Width           =   9585
         Begin VB.PictureBox wbHeader 
            BackColor       =   &H00FFFFFF&
            Height          =   1035
            Left            =   180
            ScaleHeight     =   975
            ScaleWidth      =   5835
            TabIndex        =   9
            Top             =   0
            Width           =   5895
            Begin VB.TextBox txtDescription 
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   1920
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   60
               Width           =   2895
            End
            Begin VB.Label lblLink 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "View Feed"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   225
               Left            =   1920
               TabIndex        =   11
               Top             =   780
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.Image Image1 
               Height          =   615
               Left            =   0
               Top             =   60
               Width           =   1875
            End
         End
         Begin SHDocVwCtl.WebBrowser wb 
            Height          =   975
            Left            =   240
            TabIndex        =   5
            Top             =   1020
            Width           =   3615
            ExtentX         =   6376
            ExtentY         =   1720
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin MSComctlLib.ListView lvFeeds 
         Height          =   1380
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   2434
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Headline"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgSplitH 
         Height          =   3105
         Left            =   360
         MousePointer    =   7  'Size N S
         Top             =   2580
         Width           =   150
      End
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3600
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            ImageKey        =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            ImageKey        =   "View Details"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5505
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11245
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "5/6/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:17 AM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C68
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D7A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E8C
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F9E
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10B0
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11C2
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D4
            Key             =   "View Details"
         EndProperty
      EndProperty
   End
   Begin OutlookBar.ctxOutlookBar OB1 
      Align           =   3  'Align Left
      Height          =   5085
      Left            =   0
      TabIndex        =   8
      Top             =   420
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   8969
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormatControl   =   "frmMain.frx":13E6
      FormatGroup     =   "frmMain.frx":156E
      FormatGroupHover=   "frmMain.frx":28CE
      FormatGroupPressed=   "frmMain.frx":2996
      FormatGroupSelected=   "frmMain.frx":3C7E
      FormatItem      =   "frmMain.frx":3D2A
      FormatItemLargeIcons=   "frmMain.frx":3E3A
      FormatItemHover =   "frmMain.frx":3F36
      FormatItemPressed=   "frmMain.frx":3FE2
      FormatItemSelected=   "frmMain.frx":408E
      FormatSmallIcon =   "frmMain.frx":413A
      FormatSmallIconHover=   "frmMain.frx":424A
      FormatSmallIconPressed=   "frmMain.frx":4A62
      FormatSmallIconSelected=   "frmMain.frx":52B6
      FormatLargeIcon =   "frmMain.frx":5ACE
      FormatLargeIconHover=   "frmMain.frx":5BDE
      FormatLargeIconPressed=   "frmMain.frx":63F6
      FormatLargeIconSelected=   "frmMain.frx":6C4A
      Groups          =   "frmMain.frx":7462
      OleDragMode     =   1
   End
   Begin VB.Image imgSplitter 
      Height          =   2145
      Left            =   1440
      MousePointer    =   9  'Size W E
      Top             =   2940
      Width           =   210
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilemnuFileNewFeed 
         Caption         =   "New Feed"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "Sen&d to"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Details"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Arrange &Icons"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHistory 
         Caption         =   "&History"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu mnuGroup 
      Caption         =   "GroupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuGroupUnsubscribe 
         Caption         =   "Unscubscribe"
      End
   End
   Begin VB.Menu mnuFeeds 
      Caption         =   "Feeds"
      Visible         =   0   'False
      Begin VB.Menu mnuFeedOpenInBrowser 
         Caption         =   "Open in browser"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Const TableText = "<table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='100%' id='AutoNumber1'>" & _
                   "<tr><td width='100%' align='center'>[Replace]</td></tr></table>"
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Const sglSplitLimit = 500
Dim bloading As Boolean
Dim mOpenURL As String

Private Const CREATE_NEW_CONSOLE As Long = &H10
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const INFINITE As Long = -1
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const SW_SHOWNORMAL As Long = 1

Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Private Const ERROR_BAD_FORMAT As Long = 11

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwflags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadID As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
  (ByVal lpAppName As String, _
   ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long
     
Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long

Private Declare Function FindExecutable Lib "shell32" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Private Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
  (ByVal nSize As Long, _
   ByVal lpBuffer As String) As Long
         

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    imgSplitter.Left = GetSetting(App.Title, "Settings", "vsplit", imgSplitter.Left)
    imgSplitH.Top = GetSetting(App.Title, "Settings", "hsplit", imgSplitH.Top)

    'wb.Navigate App.path & "\images\index.htm"
    wb.Navigate "about:blank"
    If cn.State = 1 Then cn.Close
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open App.path & "\rsshound.mdb"
    
    LoadSources
    If isConnectedtoNet() = True Then
        sbStatusBar.Panels(1).Text = "Online Mode" 'GetNetConnectString()
    Else
        sbStatusBar.Panels(1).Text = "Off-line Mode" ' GetNetConnectString()
    End If
    
End Sub


Private Sub Form_Paint()
    lvFeeds.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    Select Case lvFeeds.View
        Case lvwIcon
            tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
        Case lvwSmallIcon
            tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
        Case lvwList
            tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
        Case lvwReport
            tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    SaveSetting App.Title, "Settings", "ViewMode", lvFeeds.View
    SaveSetting App.Title, "Settings", "vsplit", imgSplitter.Left
    SaveSetting App.Title, "Settings", "hsplit", imgSplitH.Top
    
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
    
End Sub


Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    

    'set the width
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    OB1.Width = X
    imgSplitter.Left = X
    picRight.Left = X + 40
    picRight.Width = Me.Width - (OB1.Width + 140)



    'set the top
  

    If tbToolBar.Visible Then
        OB1.Top = tbToolBar.Height
    Else
        OB1.Top = 0
    End If

  picRight.Top = OB1.Top
    

    'set the height
    If sbStatusBar.Visible Then
        OB1.Height = Me.ScaleHeight - (sbStatusBar.Height)
    Else
        OB1.Height = Me.ScaleHeight
    End If
    

    picRight.Height = OB1.Height
    imgSplitter.Top = OB1.Top
    imgSplitter.Height = OB1.Height
End Sub



Private Sub lblLink_Click()

    wb.Navigate lblLink.Tag
    lblLink.Visible = False

End Sub

Private Sub lvFeeds_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim rs As New ADODB.Recordset
    
    'wb.Navigate Item.Tag
    
    Item.Bold = False
    Item.ListSubItems(1).Bold = False
    Item.SmallIcon = 1
    
    rs.Open "Select FeedURL, FeedName, FeedDescription from ViewedFeeds where FeedURL = '" & Item.Tag & "'", cn, adOpenDynamic, adLockOptimistic
    
    If rs.EOF And rs.BOF Then
        rs.AddNew
        rs("FeedURL") = Item.Tag
        rs("FeedName") = Item.Text
        rs("FeedDescription") = Item.ToolTipText
        rs.Update
    End If
    
    rs.Close
        
    txtDescription.Text = Item.ToolTipText 'rs("FeedDescription")
    lblLink.Tag = Item.Tag
    lblLink.Visible = True
    wb.Navigate "ABOUT:BLANK"
    
    rs.Open "Select FeedURL, feedimageurl, Feedname, FeedDescription from Feeds where FeedURL = '" & lvFeeds.Tag & "'", cn, adOpenDynamic, adLockOptimistic
    
    If Not (rs.EOF And rs.BOF) Then
        If FileExists(rs("FeedImageUrl")) Then
            Set Image1.Picture = LoadPicture(rs("FeedImageUrl"))
        Else
            Set Image1.Picture = Nothing
        End If
    End If
    
    rs.Close
    wbHeader_Resize
End Sub



Private Sub lvFeeds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        If Not lvFeeds.HitTest(X, Y) Is Nothing Then
            mOpenURL = lvFeeds.HitTest(X, Y).Tag
            PopupMenu mnuFeeds
        End If
    End If

End Sub

Private Sub mnuFeedOpenInBrowser_Click()

    If mOpenURL = "" Then Exit Sub
    
    StartNewBrowser (mOpenURL)
    
    mOpenURL = ""

End Sub

Private Sub mnuViewHistory_Click()

    Dim frm As frmhistory
    
    Set frm = New frmhistory
    
    Load frm
    
    frm.Show vbModal
    
    Unload frm

End Sub

Private Sub OB1_ButtonClick(ByVal oBtn As OutlookBar.cButton)

    Dim rs As New ADODB.Recordset
    Dim rsview As New ADODB.Recordset
    Dim oDom As New FreeThreadedDOMDocument30
    Dim xmlServer As New XMLHTTP40

    Dim oElement As IXMLDOMElement
    Dim oNodelist As IXMLDOMNodeList
    Dim xNode As IXMLDOMNode
    Dim sTemp As String
    Dim mli As MSComctlLib.ListItem
    Dim iFreeFile As Long
    Dim sPicture As String
    
    If oBtn.Key = "" Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    
    lvFeeds.ListItems.Clear
    
    rs.Open "SELECT * FROM Feeds " & _
            "WHERE Feeds.FeedID=" & Mid$(oBtn.Key, 2), cn, adOpenDynamic, adLockOptimistic

    If Not (rs.EOF And rs.BOF) Then
        xmlServer.Open "GET", rs("FeedURL")
        xmlServer.send
        
        sTemp = xmlServer.responseText
        If Left(sTemp, 5) = "<HTML" Then

            wb.Navigate rs("FEEDURL")
            If MsgBox("There appears to be an error with this link.  Would you like to unsubscribe from this feed?", vbYesNo, "Unscubscribe?") = vbYes Then
                rs("Subscribed") = 0
                rs.Update
                wb.Navigate "about:blank"
                LoadSources
                Exit Sub
            End If
        Else
            wb.Navigate "about:blank"
            Set Image1.Picture = Nothing
            txtDescription.Text = ""
            lblLink.Visible = False
            lvFeeds.Tag = rs("FEEDURL")
            sTemp = Replace(sTemp, "? 2005 Cable News Network LP, LLLP.", "")
            If InStr(1, sTemp, "<copyright") > 0 Then
                sTemp = Left(sTemp, InStr(1, sTemp, "<copyright") - 1) & Mid(sTemp, InStrRev(sTemp, "</copyright>") + 12)
            End If
            
            oDom.loadXML sTemp
            
            If Not oDom.selectSingleNode("//channel/description") Is Nothing Then
                rs("FeedDescription") = oDom.selectSingleNode("//channel/description").Text
                rs.Update
            End If
            
            If IsNull(rs("feedimageurl")) Then
ReloadImage:
                If Not oDom.selectSingleNode("//image") Is Nothing Then
                    If FileExists(App.path & "\images\" & "Feed-" & rs("feedid") & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))) Then
                        rs("feedImageUrl") = App.path & "\images\" & "Feed-" & rs("feedid") & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))
                    Else
                        DownloadFile oDom.selectSingleNode("//image/url").Text, App.path & "\images\" & "Feed-" & rs("feedid") & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))
                        rs("feedImageUrl") = App.path & "\images\" & "Feed-" & rs("feedid") & Mid$(oDom.selectSingleNode("//image/url").Text, InStrRev(oDom.selectSingleNode("//image/url").Text, "."))
                    End If
                    rs.Update
                End If
            ElseIf Not FileExists(rs("feedimageurl")) Then
                GoTo ReloadImage
            End If
'            If Not IsNull(rs("Feedimageurl")) Then
'                Set Image1.Picture = LoadPicture(rs("Feedimageurl"))
'            Else
'                Set Image1.Picture = Nothing
'            End If
            
            'txtDescription.Text = IIf(IsNull(rs("FeedDescription")), "", rs("FeedDescription"))
            
            Set oNodelist = oDom.selectNodes("//item")
            For Each xNode In oNodelist
                If Not xNode.selectSingleNode("title") Is Nothing Then
                    Set mli = lvFeeds.ListItems.Add()
                    mli.Tag = xNode.selectSingleNode("link").Text
                    If Not xNode.selectSingleNode("pubDate") Is Nothing Then
                        mli.Text = xNode.selectSingleNode("pubDate").Text
                    Else
                        mli.Text = "Unknown"
                    End If
                    
                    mli.SubItems(1) = xNode.selectSingleNode("title").Text
                    mli.ToolTipText = xNode.selectSingleNode("description").Text
                    rsview.Open "Select * from ViewedFeeds where FeedURL = '" & mli.Tag & "'", cn
                    If Not (rsview.EOF And rsview.BOF) Then
                        mli.Bold = False
                        mli.ListSubItems(1).Bold = False
                        mli.SmallIcon = 1
                    Else
                        mli.Bold = True
                        mli.ListSubItems(1).Bold = True
                        mli.SmallIcon = 2
                    End If
                    
                    rsview.Close
                End If
            Next
        End If
    End If
    
    rs.Close

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub OB1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuGroup
    End If

End Sub

Private Sub picBrowser_Resize()

    On Error Resume Next
    wbHeader.Move 0, 0, picBrowser.ScaleWidth
    wb.Move 0, wbHeader.Height, picBrowser.ScaleWidth, picBrowser.ScaleHeight - wbHeader.Height

End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    SizeSubControls imgSplitH.Top
    lvAutosizeMax lvFeeds
End Sub

Private Sub SizeSubControls(Y As Single)
    
    On Error Resume Next
    
  '  wbHeader.Top = 0
    imgSplitH.Height = 100
    imgSplitH.Width = picRight.ScaleWidth
    'set the width
    If Y < 1500 Then Y = 1500
    If Y > (picRight.ScaleHeight - (1500)) Then Y = picRight.ScaleHeight - (1500)
    lvFeeds.Top = 0
    lvFeeds.Height = Y
'    wbHeader.Height = Y
    
    imgSplitH.Top = Y
    PicSplitH.Top = Y
    PicSplitH.Height = 100
    PicSplitH.Width = picRight.ScaleWidth
    picBrowser.Top = Y + 40
    picBrowser.Height = picRight.ScaleHeight - ((lvFeeds.Height) + 140)
    picBrowser.Left = 0
    lvFeeds.Left = 0 'wbHeader.Width + 40
    lvFeeds.Width = picRight.ScaleWidth '- (wbHeader.Width + 40)
    picBrowser.Width = picRight.ScaleWidth
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            'ToDo: Add 'New' button code.
            MsgBox "Add 'New' button code."
        Case "Delete"
            mnuFileDelete_Click
        Case "Properties"
            mnuFileProperties_Click
        Case "View Large Icons"
            lvFeeds.View = lvwIcon
        Case "View Small Icons"
            lvFeeds.View = lvwSmallIcon
        Case "View List"
            lvFeeds.View = lvwList
        Case "View Details"
            lvFeeds.View = lvwReport
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
End Sub



Private Sub mnuVAIByDate_Click()
    'ToDo: Add 'mnuVAIByDate_Click' code.
'  lvListView.SortKey = DATE_COLUMN
End Sub


Private Sub mnuVAIByName_Click()
    'ToDo: Add 'mnuVAIByName_Click' code.
'  lvListView.SortKey = NAME_COLUMN
End Sub


Private Sub mnuVAIBySize_Click()
    'ToDo: Add 'mnuVAIBySize_Click' code.
'  lvListView.SortKey = SIZE_COLUMN
End Sub


Private Sub mnuVAIByType_Click()
    'ToDo: Add 'mnuVAIByType_Click' code.
'  lvListView.SortKey = TYPE_COLUMN
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileSendTo_Click()
    'ToDo: Add 'mnuFileSendTo_Click' code.
    MsgBox "Add 'mnuFileSendTo_Click' code."
End Sub

Private Sub mnuFileRename_Click()
    'ToDo: Add 'mnuFileRename_Click' code.
    MsgBox "Add 'mnuFileRename_Click' code."
End Sub

Private Sub mnuFileDelete_Click()
    'ToDo: Add 'mnuFileDelete_Click' code.
    MsgBox "Add 'mnuFileDelete_Click' code."
End Sub


Private Sub mnuFilemnuFileNewFeed_Click()
    'ToDo: Add 'mnuFilemnuFileNewFeed_Click' code.
    MsgBox "Add 'mnuFilemnuFileNewFeed_Click' code."
End Sub

Public Function LoadSources()

    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim xmlServer As New XMLHTTP40
    Dim iFreeFile As Long
    Dim sPicture As String

    OB1.Groups.Clear
    
    rs.Open "SELECT Groups.GroupId, Groups.GroupText FROM Groups WHERE CustomGroup = -1 ORDER BY Groups.GroupText", cn
    
    Do Until rs.EOF

        With OB1.Groups
            
            With .Add(rs("GroupText")).GroupItems
                
                rs2.Open "SELECT Feeds.FeedID, Feeds.FeedName, feeds.feedimageurl FROM Feeds " & _
                        "WHERE Feeds.CustomId=" & rs("Groupid") & " AND Feeds.Subscribed=-1", cn, adOpenDynamic, adLockOptimistic
                Do Until rs2.EOF
                    .Add(rs2("FeedName"), Key:="K" & rs2("FeedID"), SmallIcon:=LoadPicture(App.path & "\images\default.bmp")).Parent.IconsType = ucsIcsSmallIcons
                    rs2.MoveNext
                Loop
                
                rs2.Close
                '.Add "My Computer", LoadResPicture("la-MyComputer.bmp", vbResBitmap), LoadResPicture("sm-MyComputer.bmp", vbResBitmap)
                '.Add "Outlook Today", LoadResPicture("la-OutlookToday.bmp", vbResBitmap), LoadResPicture("sm-OutlookToday.bmp", vbResBitmap)
                '.Add "Inbox", LoadResPicture("la-Inbox.bmp", vbResBitmap), LoadResPicture("sm-Inbox.bmp", vbResBitmap)
                'With .Add("Calendar", LoadResPicture("la-Calendar.bmp", vbResBitmap), LoadResPicture("sm-Calendar.bmp", vbResBitmap))
                '    .Enabled = False
                'End With
                '.Add "Contacts", LoadResPicture("la-Contacts.bmp", vbResBitmap), LoadResPicture("sm-Contacts.bmp", vbResBitmap)
            End With
        End With
        rs.MoveNext
    Loop
    
    rs.Close
    
    rs.Open "SELECT Groups.GroupId, Groups.GroupText FROM Groups WHERE CustomGroup = 0 ORDER BY Groups.GroupText", cn
    
    Do Until rs.EOF
       ' OB1.FormatControl.BackGradient.Alpha = vbBlue
       ' OB1.FormatControl.BackGradient.SecondColor = vbBlack
        With OB1.Groups
            
            With .Add(rs("GroupText")).GroupItems
                
                rs2.Open "SELECT Feeds.FeedID, Feeds.FeedName, feeds.feedimageurl FROM Feeds " & _
                        "WHERE Feeds.GroupID=" & rs("Groupid") & " AND Feeds.Subscribed=-1", cn, adOpenDynamic, adLockOptimistic
                Do Until rs2.EOF
                    .Add(rs2("FeedName"), Key:="K" & rs2("FeedID"), SmallIcon:=LoadPicture(App.path & "\images\default.bmp")).Parent.IconsType = ucsIcsSmallIcons
                    rs2.MoveNext
                Loop
                
                rs2.Close
                '.Add "My Computer", LoadResPicture("la-MyComputer.bmp", vbResBitmap), LoadResPicture("sm-MyComputer.bmp", vbResBitmap)
                '.Add "Outlook Today", LoadResPicture("la-OutlookToday.bmp", vbResBitmap), LoadResPicture("sm-OutlookToday.bmp", vbResBitmap)
                '.Add "Inbox", LoadResPicture("la-Inbox.bmp", vbResBitmap), LoadResPicture("sm-Inbox.bmp", vbResBitmap)
                'With .Add("Calendar", LoadResPicture("la-Calendar.bmp", vbResBitmap), LoadResPicture("sm-Calendar.bmp", vbResBitmap))
                '    .Enabled = False
                'End With
                '.Add "Contacts", LoadResPicture("la-Contacts.bmp", vbResBitmap), LoadResPicture("sm-Contacts.bmp", vbResBitmap)
            End With
        End With
        rs.MoveNext
    Loop
    
    rs.Close
    
End Function

Private Sub imgSplith_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitH
        PicSplitH.Move .Left - 40, .Top, .Width, .Height / 2
    End With
    PicSplitH.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplith_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = Y + imgSplitH.Top
        If sglPos < sglSplitLimit Then
            PicSplitH.Top = sglSplitLimit
        ElseIf sglPos > picRight.ScaleHeight - sglSplitLimit Then
            PicSplitH.Top = picRight.ScaleHeight - sglSplitLimit
        Else
            PicSplitH.Top = sglPos
        End If
    End If
End Sub


Private Sub imgSplith_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeSubControls PicSplitH.Top
    PicSplitH.Visible = False
    mbMoving = False
End Sub



Private Function StartNewBrowser(sURL As String) As Boolean
       
  'start a new instance of the user's browser
  'at the page passed as sURL
   Dim success As Long
   Dim hProcess As Long
   Dim sBrowser As String
   Dim start As STARTUPINFO
   Dim proc As PROCESS_INFORMATION
   Dim sCmdLine As String
   
   sBrowser = GetBrowserName(success)
   
  'did sBrowser get correctly filled?
   If success >= ERROR_FILE_SUCCESS Then
   
      sCmdLine = BuildCommandLine(sBrowser)
      
     'prepare STARTUPINFO members
      With start
         .cb = Len(start)
         .dwflags = STARTF_USESHOWWINDOW
         .wShowWindow = SW_SHOWNORMAL
      End With
      
     'start a new instance of the default
     'browser at the specified URL. The
     'lpCommandLine member (second parameter)
     'requires a leading space or the call
     'will fail to open the specified page.
      success = CreateProcess(sBrowser, _
                              sCmdLine & sURL, _
                              0&, 0&, 0&, _
                              NORMAL_PRIORITY_CLASS, _
                              0&, 0&, start, proc)
                                  
     'if the process handle is valid, return success
      StartNewBrowser = proc.hProcess <> 0
     
     'don't need the process
     'handle anymore, so close it
      Call CloseHandle(proc.hProcess)

     'and close the handle to the thread created
      Call CloseHandle(proc.hThread)

   End If

End Function


Private Function GetBrowserName(dwFlagReturned As Long) As String

  'find the full path and name of the user's
  'associated browser
   Dim hFile As Long
   Dim sResult As String
   Dim sTempFolder As String
        
  'get the user's temp folder
   sTempFolder = GetTempDir()
   
  'create a dummy html file in the temp dir
   hFile = FreeFile
      Open sTempFolder & "dummy.html" For Output As #hFile
   Close #hFile

  'get the file path & name associated with the file
   sResult = Space$(MAX_PATH)
   dwFlagReturned = FindExecutable("dummy.html", sTempFolder, sResult)
  
  'clean up
   Kill sTempFolder & "dummy.html"
   
  'return result
   GetBrowserName = TrimNull(sResult)
   
End Function


Private Function BuildCommandLine(ByVal sBrowser As String) As String

  'just in case the returned string is mixed case
   sBrowser = LCase$(sBrowser)
   
  'try for internet explorer
   If InStr(sBrowser, "iexplore.exe") > 0 Then
      BuildCommandLine = " -nohome "
   
  'try for netscape 4.x
   ElseIf InStr(sBrowser, "netscape.exe") > 0 Then
      BuildCommandLine = " "
   
  'try for netscape 7.x
   ElseIf InStr(sBrowser, "netscp.exe") > 0 Then
      BuildCommandLine = " -url "
   
   Else
   
     'not one of the usual browsers, so
     'either determine the appropriate
     'command line required through testing
     'and adding to ElseIf conditions above,
     'or just return a default 'empty'
     'command line consisting of a space
     '(to separate the exe and command line
     'when CreateProcess assembles the string)
      BuildCommandLine = " "
      
   End If
   
End Function


Private Function TrimNull(Item As String)

  'remove string before the terminating null(s)
   Dim pos As Integer
   
   pos = InStr(Item, Chr$(0))
   
   If pos Then
      TrimNull = Left$(Item, pos - 1)
   Else
      TrimNull = Item
   End If
   
End Function


Public Function GetTempDir() As String

  'retrieve the user's system temp folder
   Dim tmp As String
   
   tmp = Space$(MAX_PATH)
   Call GetTempPath(Len(tmp), tmp)
   
   GetTempDir = TrimNull(tmp)
    
End Function

Private Sub wbHeader_Resize()
    On Error Resume Next
    Image1.Move 0, (wbHeader.ScaleHeight / 2) - (Image1.Height / 2)
    
    txtDescription.Move Image1.Width + 40, txtDescription.Top, wbHeader.ScaleWidth - (Image1.Width + 40)
    lblLink.Left = txtDescription.Left
End Sub
