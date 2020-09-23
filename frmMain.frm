VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "TabPad"
   ClientHeight    =   5145
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7290
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRandomTip 
      Interval        =   20000
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer tmrAutoRecoverySave 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   0
      Top             =   600
   End
   Begin VB.Frame fraSettings 
      Caption         =   "TabPad Settings"
      Height          =   1695
      Left            =   480
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtRecentDocuments 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "10"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkCrashRecovery 
         Caption         =   "Enable Crash Recovery"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdSettingsCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdSettingsSet 
         Caption         =   "Set"
         Default         =   -1  'True
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Recent Documents:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Timer tmrStatusBar 
      Interval        =   50
      Left            =   1200
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   4890
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   450
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   7225
            MinWidth        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Text            =   "Line 1, Col 0"
            TextSave        =   "Line 1, Col 0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraTabKeySettings 
      Caption         =   "Tab Key Settings"
      Height          =   1695
      Left            =   3000
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton cmdTabKeySet 
         Caption         =   "Set"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdTabKeyCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtTabKeyCustAmount 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "4"
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optTabKeyCustom 
         Caption         =   "Custom amount"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optTabKeyDefault 
         Caption         =   "System default (~ 6)"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CheckBox chkTabIS 
         Caption         =   "Tab inserts spaces"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCloseTab 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   1
      Top             =   35
      Width           =   255
   End
   Begin VB.Frame fraFontSettings 
      Caption         =   "Font Settings"
      Height          =   1695
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdSaveFonts 
         Caption         =   "Set"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox cmbFSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":2CFA
         Left            =   600
         List            =   "frmMain.frx":2D28
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbFonts 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblExample 
         Caption         =   "Example"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Font:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Timer tmrLoadPH 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   0
   End
   Begin VB.TextBox txtNoWrap 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin VB.TextBox txtWrap 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Timer tmrCheckFocus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.TabStrip Tabs 
      Height          =   3825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6747
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Untitled 1"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFNew 
         Caption         =   "&New Tab"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFCloseTab 
         Caption         =   "&Close Tab"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFBrk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFMore 
         Caption         =   "More"
         Begin VB.Menu mnuFSaveAll 
            Caption         =   "Save All"
            Shortcut        =   ^Q
         End
         Begin VB.Menu mnuFBrk2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFLoadTabset 
            Caption         =   "Load Tabset"
         End
         Begin VB.Menu mnuFSaveTabset 
            Caption         =   "Save Tabset"
         End
      End
      Begin VB.Menu mnuFBrk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFSettings 
         Caption         =   "Settings"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuFBrk4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFRecent 
         Caption         =   "Recent Documents"
         Begin VB.Menu mnuFRecentList 
            Caption         =   "Recent Doc"
            Index           =   0
            Shortcut        =   ^{INSERT}
         End
         Begin VB.Menu mnuFRDBrk 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFRecentClear 
            Caption         =   "Clear List"
         End
      End
      Begin VB.Menu mnuFNotepad 
         Caption         =   "Open in Notepad"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuESA 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEClear 
         Caption         =   "Clear"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEBrk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEFind 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEReplace 
         Caption         =   "Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEBrk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEInsertDT 
         Caption         =   "Insert Date + Time"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEAllowTabs 
         Caption         =   "Tab Key Settings"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Begin VB.Menu mnuFWordWrap 
         Caption         =   "Word Wrap"
      End
      Begin VB.Menu mnuFFont 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuVStatus 
         Caption         =   "Status Bar"
      End
      Begin VB.Menu mnuVSBRandTips 
         Caption         =   "Random Tips on Status Bar"
      End
      Begin VB.Menu mnuVBrk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVBlackBack 
         Caption         =   "Black Background/White Text"
      End
      Begin VB.Menu mnuVRandomize 
         Caption         =   "Randomize current colours"
      End
      Begin VB.Menu mnuVBrk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVDocInfo 
         Caption         =   "Document Info"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHRandomTip 
         Caption         =   "New Random Tip"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHBrk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuStatus 
      Caption         =   ""
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Michael Dickens ( LFI.net )
'Source code released under GPL
'Please leave credit to the author (and my website) in any modifications made
'All modifications must be opensource, with source code available with binary

'Project page: www.LFI.net/LFI/prjTP.htm
'MichaelDickens@gmail.com
Option Explicit
Public RetrieveMode As Boolean
Public SrchStr As String
Public Pos As Long, Place As Long
Public HadResult As Boolean

Private ShowRecoverySaved As Boolean

Public txtMain As TextBox

Private Sub chkBold_Click()
lblExample.FontBold = chkBold.Value 'On Fonts frame - for 'Example text'
End Sub

Private Sub chkTabIS_Click()
If chkTabIS.Value = 1 Then
    Me.optTabKeyCustom.Visible = True
    Me.optTabKeyDefault.Visible = True
    If Me.optTabKeyCustom.Value = True Then
        Me.txtTabKeyCustAmount.Visible = True
    Else
        Me.txtTabKeyCustAmount.Visible = False
    End If
Else
    Me.optTabKeyCustom.Visible = False
    Me.optTabKeyDefault.Visible = False
    Me.txtTabKeyCustAmount.Visible = False
End If
End Sub

Private Sub cmbFonts_Click()
'On Fonts frame - for 'Example text'
lblExample.FontName = cmbFonts.Text
lblExample.FontBold = Me.chkBold.Value
lblExample.FontItalic = False
lblExample.FontUnderline = False
End Sub

Private Sub cmbFSize_Click()
'On Fonts frame - for 'Example text'
lblExample.FontSize = cmbFSize.Text
End Sub

Private Sub cmdCancel_Click()
'On Fonts frame
Me.fraFontSettings.Visible = False
End Sub

Private Sub cmdCloseTab_Click()
Dim i As Integer
If Tabs.Tabs.Count = 1 Then
    'Won't actually occur, must be old code
    'Oh wait, probably happens on Form_Unload
    Exit Sub
End If

Dim SelIndex As Integer
If Right(Tabs.SelectedItem.Caption, 3) = "  *" Then 'If the asterisk is after the filename, it's not saved (or up-to-date saved)
        Select Case MsgBox("Do you want to save the changes to " & Left(Tabs.SelectedItem.Caption, Len(Tabs.SelectedItem.Caption) - 3) & "?", vbQuestion + vbYesNoCancel) 'Do you want to Save messagebox
            Case vbYes 'Yes, save the document
                mnuFSave_Click
            
            Case vbCancel 'I don't want to close the tab anymore
                Exit Sub
            
            'If no is selected then we continue on (like Yes) but without saving
        End Select
End If

If Tabs.SelectedItem.Index = Tabs.Tabs.Count Then
    'If it's the very last tab it is easy to remove info from TabInfo
    SelIndex = Tabs.SelectedItem.Index - 1 'Selected Index - TabInfo indexes are always one less than that of the control 'Tabs' (control Tabs starts at #1, TabInfo starts at 0)
    Tabs.Tabs.Remove Tabs.SelectedItem.Index 'Remove it from tabs column
    Tabs.Tabs(Tabs.Tabs.Count).Selected = True 'Selects the 2nd last (now last) tab
    TabInfo(SelIndex).FilePath = ""
    TabInfo(SelIndex).Text = ""
    TabInfo(SelIndex).SelStart = 0
    TabInfo(SelIndex).SelLength = 0
    'Above resets TabInfo entry
    CheckTabs 'Deals with Close Tab button (relocates and/or hides/displays)
    Exit Sub
End If

'The tab being closed is NOT the last. It can be anywhere before it and this code will work.
'Writing this code was the most time consuming and debugged routine of TabPad. It sucked.
SelIndex = Tabs.SelectedItem.Index
Tabs.Tabs(SelIndex + 1).Selected = True 'We select the tab AFTER the one being closed

'Tab 1
'Tab 2
'Tab 3 < Closing
'Tab 4
'Tab 5

For i = SelIndex To (Tabs.Tabs.Count - 1)
    'This will relocate tab info AHEAD of the closing tab to the location before it.
    'eg. If you're closing Tab 3, Tab 4's info will go to Tab 3's info location and Tab 5's info will go to where to Tab 4's was. If that make sense
    TabInfo(i - 1).Text = TabInfo(i).Text
    TabInfo(i - 1).SelStart = TabInfo(i).SelStart
    TabInfo(i - 1).SelLength = TabInfo(i).SelLength
    TabInfo(i - 1).FilePath = TabInfo(i).FilePath

    TabInfo(i).Text = ""
    TabInfo(i).SelStart = 0
    TabInfo(i).SelLength = 0
    TabInfo(i).FilePath = ""
Next i

Tabs.Tabs.Remove SelIndex 'Finally we remove it from the control

RetrieveMode = True
Tabs.Tabs(SelIndex).Selected = True
'We change our selected Tab... again.
'This is where all the issues occured. Thankfully, slapping a
'RetrieveMode boolean situation in allowed me to fix one bug I
'will never fully bother understanding. It was simply that I didnt
'know exactly how my array worked - but the code for when a Tab
'is clicked will show that it was probably the best way anyway
RetrieveMode = False

CheckTabs 'Yet again, locate / change visibility of Close tab button
End Sub

Private Sub cmdSaveFonts_Click()
SaveSetting "TabPad", "Settings", "Font", Me.cmbFonts.Text
SaveSetting "TabPad", "Settings", "FontSize", Me.cmbFSize.Text
SaveSetting "TabPad", "Settings", "FontBold", Me.chkBold
'We save our new font settings to registry

LoadFSettings
'Load them back from registry into control properties

Me.fraFontSettings.Visible = False
'Close frame
End Sub

Private Sub cmdSettingsCancel_Click()
'On tab pad settings frame
Me.fraSettings.Visible = False
End Sub

Private Sub cmdSettingsSet_Click()
'On tab pad settings frame
Me.fraSettings.Visible = False

SaveSetting "TabPad", "Settings", "CrashRecovery", Me.chkCrashRecovery

If IsNumeric(txtRecentDocuments) = True Then
    If txtRecentDocuments.Text > 32 Then
        MsgBox "Maximum recent documents TabPad will keep is 32"
        txtRecentDocuments.Text = 32
    End If
        SaveSetting "TabPad", "Settings", "RecentDocuments", Me.txtRecentDocuments.Text
        MaxRecentDocs = txtRecentDocuments.Text
End If
End Sub

Private Sub cmdTabKeyCancel_Click()
'On tab key settings frame
Me.fraTabKeySettings.Visible = False
End Sub

Private Sub cmdTabKeySet_Click()
SaveSetting "TabPad", "Settings", "TabAsSpaces", Me.chkTabIS.Value
Me.fraTabKeySettings.Visible = False
If Me.optTabKeyDefault = True Then
    SaveSetting "TabPad", "Settings", "TabKeyAmount", "Default"
Else
    SaveSetting "TabPad", "Settings", "TabKeyAmount", txtTabKeyCustAmount
End If
LoadTKSettings
End Sub


Private Sub Form_Load()
Dim Comm As String
Dim i As Integer
Comm = Command

If App.PrevInstance = True Then
    'TabPad is already running!
    If Comm = "" Then
        'We'll make our original instance open a new tab
        AddToPH "NEWBLANK"
        End
    Else
        'Command line includes filename (probably from Windows)
        If Left(Comm, 1) = Chr(34) And Right(Comm, 1) = Chr(34) Then
            'Remove quoation marks around path
            Comm = Mid(Comm, 2, Len(Comm) - 2)
        End If
        AddToPH Comm
        'AddToPH - PH stands for placeholder.
        'This new instance will write to a file available to the
        'old instance and the old will load it (thx 2 timer event)
        End
        'Don't need this instance
    End If
End If

txtNoWrap.Visible = False
txtWrap.Visible = False
'Depends on wordwrap setting

Me.Caption = "TabPad v" & App.Major & "." & App.Minor & App.Revision
'Setup our title

ReDim TabInfo(0)
'TabInfo needs to be created so it can be resized and used as required

RetrieveMode = True
'RetrieveMode is a small boolean that was made for one specific purpose early on in the app.
'I believe it was swapping text between Wrap and Non-Wrap textboxes
'It is now used to tell a few parts of the code that an automated event is occuring and that it should not be saved, etc.
Me.Height = GetSetting("TabPad", "Settings", "Height", Me.Height)
Me.Width = GetSetting("TabPad", "Settings", "Width", Me.Width)
RetrieveMode = False
Form_Resize 'Make sure everything fits

SetRandomTips
LoadSettings 'Load settings (includes Fonts + Tab Key prefs)

For i = 1 To MaxRecentDocs
'We pull recent documents from registry
    If Not GetSetting("TabPad", "Recent", "Recent" & i, "x") = "x" Then
        RecentDocs.Add GetSetting("TabPad", "Recent", "Recent" & i, "x")
    End If
Next i
PopulateRDMenu 'Fix up the Recent Documents slideout on File menu

CheckTabs 'Simply hides or displays a close tab button (also positions it)

If Not Comm = "" Then
    'If the command line arguements are not blank..
        If Left(Comm, 1) = Chr(34) And Right(Comm, 1) = Chr(34) Then
            'Remove quoation marks around path
            Comm = Mid(Comm, 2, Len(Comm) - 2)
        End If
    Me.Show
    'Otherwise Notepad will load then TabPad will superimpose over it
    LoadFile Comm
    'Load the file
    UntitledCount = 0
    'Next tab created will be Untitled 1, instead of 2.
Else
    UntitledCount = 1
    'We make a tab called Untitled 1.
End If
tmrLoadPH = True

'We check if Crash Recovery option has been set
If GetSetting("TabPad", "Settings", "CrashRecovery", "-1") = "-1" Then
    If MsgBox("Newer versions of TabPad now support 'Crash Recovery' - if TabPad crashes, or your computer loses power, your session can be recovered! To enable Crash Recovery, select Yes", vbQuestion + vbYesNo, "Enable Crash Recovery?") = vbYes Then
        SaveSetting "TabPad", "Settings", "CrashRecovery", "1"
    Else
        SaveSetting "TabPad", "Settings", "CrashRecovery", "0"
    End If
End If



Me.Show
'Show the form. Otherwise focusing will fail.
Form_Resize

Me.txtMain.SetFocus
'Focus


If FileExists(App.Path & "\TPLoadList.dat") = True Then
    If GetSetting("TabPad", "Settings", "CrashRecovery", "0") = "1" Then
        If FileExists(App.Path & "\TabPad-Recovery\Session.dat") = False Then
            MsgBox "It appears TabPad did not close correctly, and the crash recovery information was not retrievable. Sorry for any inconvenience"
        Else
            If MsgBox("It appears TabPad did not close correctly, would you like to recover your session using the crash recovery information?", vbQuestion + vbYesNo, "Recover last session?") = vbYes Then
                RecoverSession
            Else
                MsgBox "Ok, crash recovery information has been cleared."
            End If
        End If
    Else
        MsgBox "It appears TabPad did not close correctly, and you have crash recovery disabled. To enable this, go into File -> Settings, and next time this happens your session will be recoverable"
    End If
End If



MakePlaceHolder 'Make the file new instances will drop paths into


tmrAutoRecoverySave = True
tmrAutoRecoverySave_Timer

tmrRandomTip_Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Program wants to close
Dim OrigTabCount As Integer
Dim i As Integer

If Not Tabs.Tabs.Count = 1 Then
    If MsgBox("You are about to close " & Tabs.Tabs.Count & " tabs. Are you sure you want to continue?", vbQuestion + vbYesNo, "Close TabPad?") = vbNo Then
        Cancel = True
        Exit Sub
    End If
End If

For i = Tabs.Tabs.Count To 1 Step -1
    'We go from the last tab open to the first tab open
    Tabs.Tabs(i).Selected = True
    'We select it
    OrigTabCount = Tabs.Tabs.Count
    'Get the index
    cmdCloseTab_Click
    'Click the close button to prompt for save
    If OrigTabCount = Tabs.Tabs.Count And Not i = 1 Then
        'Tab didn't close, user cancelled - abort closing..
        Cancel = True
        'Tells windows we're not closing
        Exit Sub
    End If
Next i

If Right(Tabs.SelectedItem.Caption, 3) = "  *" Then
    'We're at the last tab, Close Tab button doesn't work anymore ;)
    Select Case MsgBox("Do you want to save the changes to " & Left(Tabs.SelectedItem.Caption, Len(Tabs.SelectedItem.Caption) - 3) & "?", vbQuestion + vbYesNoCancel)
        Case vbYes
        mnuFSave_Click
        
        Case vbCancel
        Cancel = True
        'User cancelled with 1 tab to go - Tell windows we're not closing
        Exit Sub
        
    End Select
End If


KillPH
'Kills the placeholding file that other instances writh paths too
End
'Shutdown entirely
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub 'It's minimised
If RetrieveMode = True Then Exit Sub 'Automated procedure
On Error Resume Next

'Tabs.Width = Me.Width - 120
Tabs.Width = Me.Width - 420
'Make it smaller so that the tabs dont go as far. it's cool so the button is perfecto

Tabs.Height = Me.Height
'Tabs control stretches to height of our form

CheckTabs
'Hide or show / position close tab button if needed

txtMain.Width = Me.Width - 110
'Set width of current active textbox

If StatusBarOn = False Then
    txtMain.Height = Me.Height - 1160
Else
    txtMain.Height = Me.Height - 1400
End If
'And height

'StatusBar.Panels(1).Width = Me.Width - StatusBar.Panels(2).Width - 500


If Not Me.WindowState = 0 Then Exit Sub 'If the state isn't normal (eg. It's maximised) then exit procedure
SaveSetting "TabPad", "Settings", "Width", Me.Width
SaveSetting "TabPad", "Settings", "Height", Me.Height
'Save positioning to registry for next run


End Sub

Private Sub mnuEAllowTabs_Click()
'Show settings
Me.cmdTabKeyCancel.Cancel = True
Me.cmdTabKeySet.Default = True

Me.fraTabKeySettings.Top = (Me.Height - Me.fraTabKeySettings.Height) / 2
Me.fraTabKeySettings.Left = (Me.Width - (Me.fraTabKeySettings.Width)) / 2
'Position in (almost) center of screen

    'Read settings for frame:
    Me.chkTabIS.Value = GetSetting("TabPad", "Settings", "TabAsSpaces", "1")
    
    If Me.chkTabIS.Value = 1 Then
        'Read more settings if tab key is enabled
        If GetSetting("TabPad", "Settings", "TabKeyAmount", "Default") = "Default" Then
            Me.optTabKeyDefault.Value = True
        Else
            Me.optTabKeyCustom.Value = True
            Me.txtTabKeyCustAmount = GetSetting("TabPad", "Settings", "TabKeyAmount", "4")
        End If
    End If

Me.fraTabKeySettings.Visible = True
End Sub

Private Sub mnuEClear_Click()
'Clear text on current tab
If Me.txtMain.Text = "" Then Exit Sub 'No need to
Me.txtMain.Text = ""
'Clear it
txtMain_Change
'Initiation change
If Not Right(Me.Tabs.SelectedItem.Caption, 3) = "  *" Then
    Me.Tabs.SelectedItem.Caption = Me.Tabs.SelectedItem.Caption & "  *"
End If
'Make sure it shows up as different ;)
End Sub

Private Sub mnuEFind_Click()
'Find in current tab text
Dim SearchStr As String
Dim ISResp As Long
SearchStr = InputBox("Search for what?", "Searching document") 'Shows generic inputbox
If SearchStr = "" Then Exit Sub 'If they enter nothing we can't search
Pos = 0 'Start from top of document
SrchStr = SearchStr 'Looking for..
HadResult = False 'Has this search yielded any finds?
mnuEFindNext_Click 'Click the FindNext menu item (search code is there)
End Sub

Private Sub mnuEFindNext_Click()
If SrchStr = "" Then
    mnuEFind_Click
    'Huh? No search term! Click Find button
    Exit Sub
End If
If txtMain.Text = "" Then Exit Sub 'No text to search


If Pos = 0 Then Pos = 1 'We can't search before the 1st character
'Match Case Searching Code:
'Place = InStr(Pos, txtMain.Text, SrchStr, vbBinaryCompare)

'Or Don't:
Place = InStr(Pos, txtMain.Text, SrchStr, vbTextCompare)
'InStr returns the position in the string (textbox) the first character of the search string is
If Place = 0 Then
    'Couldn't find Search String
    Pos = 0
    'Reset position
    If HadResult = False Then
        'No results for search
        MsgBox "Couldn't find " & Chr(34) & SrchStr & Chr(34), vbInformation
        SrchStr = ""
        'Reset search string, it's useless
    Else
        'We have had results, but there were no more after the last.
        If MsgBox("Finished searching to the end of the document - would you like to continue searching from the top?", vbQuestion + vbYesNo) = vbYes Then
            'Ask user if they want to start over
            'They said Yes
            txtMain.SelStart = 0
            txtMain.SelLength = 0
            mnuEFindNext_Click
            Exit Sub
        Else
            'Ask user if they want to start over
            'They said No
            SrchStr = ""
            'Reset string, user doesnt want to start over
            Exit Sub
        End If
    End If
    Exit Sub
End If

txtMain.SelStart = Place - 1
txtMain.SelLength = Len(SrchStr)
txtMain.SetFocus
'Highlight the search result
Pos = Place + 1
'Update position so we don't find it again
HadResult = True
'YES this search has produced results!
End Sub

Private Sub mnuEInsertDT_Click()
'Me.txtMain.SelText = Me.txtMain.SelText & Now
Dim ToD As String
If Hour(Time$) > 12 Then
    ToD = "PM"
Else
    ToD = "AM"
End If
Me.txtMain.SelText = Me.txtMain.SelText & Hour(Time) & ":" & Minute(Time) & " " & ToD & " " & Date
End Sub

Private Sub mnuEReplace_Click()
'String replacement
Dim OrigStr As String
Dim RepStr As String
OrigStr = InputBox("Enter string to be replaced", "String to replace") 'Generic input box
If OrigStr = "" Then Exit Sub 'They entered nothing or cancelled
RepStr = InputBox("Enter string to replace with", "String replacement") 'Generic input box
If RepStr = "" Then Exit Sub 'Same as above (2 up)
txtMain.Text = Replace(txtMain.Text, OrigStr, RepStr) 'Do replacements (simple)
End Sub

Private Sub mnuESA_Click()
'Select all
Me.txtMain.SelStart = 0
Me.txtMain.SelLength = Len(Me.txtMain.Text)
'Select from start to finish
End Sub

Private Sub mnuFCloseTab_Click()
cmdCloseTab_Click 'Close tab menu click prompts button to be ... cluck? lol clicked.
End Sub

Private Sub mnuFExit_Click()
Form_QueryUnload 0, 0 'Pretend Windows/User wants to shut it down so Save Yes/No/Cancel routine occurs
End Sub

Private Sub mnuFFont_Click()
'Edit Font frame to be shown
Dim SortTree As New clsTree 'Cool code obtained from Planet Source Code to alphabetically sort
Dim i As Integer
Me.fraFontSettings.Top = (Me.Height - Me.fraFontSettings.Height) / 2
Me.fraFontSettings.Left = (Me.Width - (Me.fraFontSettings.Width)) / 2
'Center (almost) frame on screen

Me.cmdCancel.Cancel = True
Me.cmdSaveFonts.Default = True

Me.cmbFonts.Clear
For i = 1 To Screen.FontCount
    SortTree.AddItem Screen.Fonts(i), ""
Next i
'Pull fonts from system into sorting mechanism

Dim Results() As String
Dim Tags() As String
Dim NoResults As Integer
    
    NoResults = SortTree.SortItems(False, Results, Tags) 'ORIGINAL PSCODE COMMENT: reverse returns the results in reverse alphabetical order, results and tags are dynamic arrays that will hold the sorted tags and results
    For i = 2 To NoResults 'Number 1 produces weird result for some reason
        cmbFonts.AddItem Results(i) 'Add to onscreen dropdown
        If Results(i) = Me.txtMain.Font Then
            'Hey, that's the font we're using in! go to it
            Me.cmbFonts.ListIndex = Me.cmbFonts.ListCount - 1
        End If
    Next i

'Go through size list and select current size:
For i = 0 To Me.cmbFSize.ListCount - 1
    If cmbFSize.List(i) = GetSetting("TabPad", "Settings", "FontSize", Me.txtMain.FontSize) Then
        'Registry size matches (pulling size from txtMain skews results as it uses decimals)
        Me.cmbFSize.ListIndex = i
        Exit For
    End If
Next i

If Me.cmbFSize.Text = "" Then
    Me.cmbFSize.ListIndex = 2
    'Default to size 10
End If

If Me.txtMain.FontBold = True Then 'Check or uncheck Bold checkbox
    chkBold.Value = 1
Else
    chkBold.Value = 0
End If
Me.fraFontSettings.Visible = True 'Show frame
End Sub

Private Sub mnuFLoadTabset_Click()
MsgBox "[Possible] Future Feature"
End Sub

Private Sub mnuFNew_Click()
'Open a new tab
If Tabs.SelectedItem.Index = Tabs.Tabs.Count Then
    'If we're at the last tab and it's blank, dont make a new one
    If TabInfo(Tabs.SelectedItem.Index - 1).FilePath = "" Then
        If txtMain.Text = "" Then Exit Sub
    End If
Else
    If TabInfo(Tabs.Tabs.Count - 1).Text = "" And TabInfo(Tabs.Tabs.Count - 1).FilePath = "" Then
        'If the last tab is blank, don't make a new one, just skip to the last blank
        Tabs.Tabs(Tabs.Tabs.Count).Selected = True
        Exit Sub
    End If
End If

If RetrieveMode = False Then
    'RetrieveMode is on when a file is being loaded. I like my untitled numbering to work correctly
    UntitledCount = UntitledCount + 1
    'Increment if the user has clicked it
End If
Tabs.Tabs.Add , , "Untitled " & UntitledCount
'Add tab with funky title
ReDim Preserve TabInfo(Tabs.Tabs.Count - 1)
'Make tab info
CheckTabs 'Fix that pesky close button
RetrieveMode = False
Tabs.Tabs(Tabs.Tabs.Count).Selected = True 'Jump to new tab
txtMain.Text = "" 'Reset text if not done
End Sub

Private Sub mnuFNotepad_Click()
'Open current file in Notepad
If TabInfo(LastTabIndex).FilePath = "" Then
    'File isn't saved
    MsgBox "You need to save your file before you can view it in Notepad.", vbExclamation
    Exit Sub
End If
mnuFSave_Click 'Update file on disk
On Error Resume Next
'Try and open it :D
Shell "C:\Windows\Notepad.exe " & TabInfo(LastTabIndex).FilePath, vbNormalFocus
End Sub

Private Sub mnuFOpen_Click()
'Open file
Dim FilePath As String
Dim Dia As New clsDialog
FilePath = Dia.ShowOpen(Me.hwnd, "Open a text document", App.Path, "Text Files (*.txt) | *.txt|All Files | *.*")
'Bring up dialog box via API

If FilePath = "" Then Exit Sub 'They didnt chose one

LoadFile FilePath 'File is loaded by this routine
End Sub

Private Sub mnuFRecentClear_Click()
'Clear recent documents list
Set RecentDocs = New Collection 'Reset internal list
DeleteRRK 'Delete Recent Registry Keys
PopulateRDMenu 'Show up blank menu
End Sub

Private Sub mnuFRecentList_Click(Index As Integer)
LoadFile mnuFRecentList(Index).Tag 'We load recent document :)
End Sub

Private Sub mnuFSave_Click()
'Save file
Dim i As Integer

If txtMain.Text = "" Then
    'Option of whether or not to save
    If MsgBox("There is no text entered - are you sure you want to save?", vbQuestion + vbYesNo) = vbNo Then
        'They chose not to
        Exit Sub
    End If
End If

'Should be an option:
'If txtMain.Text = "" Then Exit Sub 'No text, why save?
If TabInfo(Tabs.SelectedItem.Index - 1).FilePath = "" Then
    'Havent saved before, show dialog
    mnuFSaveAs_Click
    Exit Sub
End If

Open TabInfo(Tabs.SelectedItem.Index - 1).FilePath For Output As #1
    Print #1, txtMain.Text
Close #1
'Output textbox to file

Dim TabTitle As String
TabTitle = GetFileName(TabInfo(Tabs.SelectedItem.Index - 1).FilePath)
If LCase(Right(TabTitle, 4)) = ".txt" Then
    TabTitle = Left(TabTitle, Len(TabTitle) - 4)
End If
Tabs.SelectedItem.Caption = TabTitle
'Fix up tab title

If RecentDocs.Count = 0 Then
    RecentDocs.Add TabInfo(Tabs.SelectedItem.Index - 1).FilePath
Else
    RecentDocs.Add TabInfo(Tabs.SelectedItem.Index - 1).FilePath, , 1
End If
'Add our new save to recent documents, or atleast pop it at the top

reloopdel:
'Check if we have duplicate recents
For i = 2 To RecentDocs.Count
    If RecentDocs.item(i) = TabInfo(Tabs.SelectedItem.Index - 1).FilePath Then
        RecentDocs.Remove i
        'Remove older dupe.
        GoTo reloopdel
    End If
Next i

Do Until RecentDocs.Count <= MaxRecentDocs
    RecentDocs.Remove RecentDocs.Count
    'Remove recents that arent that recent anymore
Loop
PopulateRDMenu 'Re populate recent documents menu slideout
End Sub

Private Sub mnuFSaveAll_Click()
'Save all tabs
Dim OrigSelection As Integer
Dim i As Integer
OrigSelection = Tabs.SelectedItem.Index 'Original index selected so we can go back to it
For i = Tabs.Tabs.Count To 1 Step -1 'Go from last tab to first tab
    Tabs.Tabs(i).Selected = True 'Select it
    If Right(Tabs.SelectedItem.Caption, 3) = "  *" Then 'Check if it needs to be saved
        If MsgBox("Do you want to save the changes to " & Left(Tabs.SelectedItem.Caption, Len(Tabs.SelectedItem.Caption) - 3) & "?", vbQuestion + vbYesNo) = vbYes Then 'Popup dialog
            mnuFSave_Click 'Save if told to
        End If
    End If
Next i
Tabs.Tabs(OrigSelection).Selected = True 'Go back to original selection
End Sub

Private Sub mnuFSaveAs_Click()
'Save As
Dim FilePath As String
Dim Dia As New clsDialog
FilePath = Dia.ShowSave(Me.hwnd, "Save your text document", App.Path, "Text File (*.txt) | *.txt|All Files | *.*", "txt")
'Show save dialog using API
If FilePath = "" Then Exit Sub 'No name given, can't save
TabInfo(Tabs.SelectedItem.Index - 1).FilePath = FilePath 'Update array

'Tabs.SelectedItem.ToolTipText = FilePath

mnuFSave_Click 'Actually save with this new path :)
End Sub

Private Sub mnuFSaveTabset_Click()
MsgBox "[Possible] Future Feature"
Exit Sub

Dim i As Integer
Dim NotSaved As Integer
For i = 0 To UBound(TabInfo)
    If TabInfo(i).FilePath = "" Then
        NotSaved = NotSaved + 1
    End If
Next i
If NotSaved = Tabs.Tabs.Count Then
    MsgBox "None of your current tabs are saved. To make a tabset, physical file locations must be available for each tab. Save your documents then try again"
    Exit Sub
Else
    MsgBox NotSaved & " tabs (of " & Tabs.Tabs.Count & ") aren't saved. These tabs will not be included in the tabset, however it will still be created. You can save each document then resave your tabset if you wish."
End If

Dim FilePath As String
Dim Dia As New clsDialog
FilePath = Dia.ShowSave(Me.hwnd, "Save a tabset", App.Path, "TabPad Tabsets (*.tpt) | *.tpt|All Files | *.*", "tpt")
'Show save dialog using API
If FilePath = "" Then Exit Sub 'No name given, can't save

Open FilePath For Output As #22
    For i = 1 To UBound(TabInfo)
        If Not TabInfo(i).FilePath = "" Then
            Print #22, TabInfo(i).FilePath
        End If
    Next i
Close #22

End Sub

Private Sub mnuFSettings_Click()
Me.cmdSettingsSet.Default = True
Me.cmdSettingsCancel.Cancel = True
Me.fraSettings.Top = (Me.Height - Me.fraSettings.Height) / 2
Me.fraSettings.Left = (Me.Width - (Me.fraSettings.Width)) / 2

Me.chkCrashRecovery = GetSetting("TabPad", "Settings", "CrashRecovery", "1")
Me.txtRecentDocuments = GetSetting("TabPad", "Settings", "RecentDocuments", "10")

Me.fraSettings.Visible = True 'Show frame
End Sub

Private Sub mnuFWordWrap_Click()
'Change wordwrap settings
RetrieveMode = True
If GetSetting("TabPad", "Settings", "WordWrap", "1") = "0" Then
    mnuFWordWrap.Checked = True 'Change checking
    SaveSetting "TabPad", "Settings", "WordWrap", "1" 'Update reg
    TabInfo(LastTabIndex).Text = txtMain.Text 'Put text into array
    TabInfo(LastTabIndex).SelStart = txtMain.SelStart 'And position
    TabInfo(LastTabIndex).SelLength = txtMain.SelLength '..position len
    txtMain = "" 'Clear text
    txtMain.Visible = False 'Hide old box
    Set txtMain = txtWrap 'Change our little variable to new textbox
    txtMain.Visible = True 'Make it visible
    txtMain = TabInfo(LastTabIndex).Text 'Load text back from array
    txtMain.SelStart = TabInfo(LastTabIndex).SelStart 'Reposition
    txtMain.SelLength = TabInfo(LastTabIndex).SelLength 'and len..
Else
    'Read comments above
    mnuFWordWrap.Checked = False
    SaveSetting "TabPad", "Settings", "WordWrap", "0"
    TabInfo(LastTabIndex).Text = txtMain
    TabInfo(LastTabIndex).SelStart = txtMain.SelStart
    TabInfo(LastTabIndex).SelLength = txtMain.SelLength
    txtMain = ""
    txtMain.Visible = False
    Set txtMain = txtNoWrap
    txtMain.Visible = True
    txtMain = TabInfo(LastTabIndex).Text
    txtMain.SelStart = TabInfo(LastTabIndex).SelStart
    txtMain.SelLength = TabInfo(LastTabIndex).SelLength
End If
RetrieveMode = False
Form_Resize 'Fix sizing of new textbox
'Not sure why this code keeps being used:
'LoadSettings
'Should be:
LoadFSettings 'Get fonts
txtMain.SetFocus 'Give focus to box
End Sub

Private Sub mnuHAbout_Click()
'About menu click
RetrieveMode = True
mnuFNew_Click 'Make a new tab to show text in
RetrieveMode = False
TabInfo(Tabs.Tabs.Count - 1).FilePath = "" 'No path for this one
'Me.txtMain = "TabPad v" & App.Major & "." & App.Minor & " - written by Michael Dickens ( www.LFI.net )" & vbCrLf & vbCrLf & "TabPad lets people from today deal with tomorrow's problems using technology from yesterday. Tabbed solutions are nothing new, but the lack of ability to deal with text files in such an intelligent way is primitive. TabPad uses the standard TabStrip control built into the Visual Basic runtimes, along with a generic text box and a couple of command buttons." & vbCrLf & vbCrLf & "Thanks to Telly from PSCode.com for enhanced Find code." & vbCrLf & vbCrLf & "Written by Michael Dickens - www.LFI.net" & vbCrLf & "Email: MichaelDickens@gmail.com" 'Put text into box
Me.txtMain = "TabPad v" & App.Major & "." & App.Minor & " - written by Michael Dickens ( www.LFI.net )" & vbCrLf & vbCrLf & "TabPad lets you deal with as many text files as you want, at the same time, all in one window. Using code written to be speedy, TabPad has no problem dealing with dozens of tabs with content from anywhere on your PC. TabPad also helps to keep you on track by only allowing you to open a file in TabPad once, unlike Notepad which will let you accidentally edit the file in multiple windows and leave you with many different versions. TabPad is completely free, and the sourcecode for it can be downloaded from my website." & vbCrLf & vbCrLf & "Check for updates: http://lfi.net/LFI/prjTP.htm" & vbCrLf & vbCrLf & "Thanks to Telly from PSCode.com for enhanced Find code." & vbCrLf & vbCrLf & "Written by Michael Dickens - www.LFI.net" & vbCrLf & "Email: MichaelDickens@gmail.com"   'Put text into box
Tabs.SelectedItem.Caption = "About TabPad" 'Set title of tab
End Sub

Private Sub mnuHRandomTip_Click()
tmrRandomTip_Timer
tmrRandomTip = False
tmrRandomTip = True
End Sub

Private Sub mnuStatus_Click()
mnuVStatus_Click
End Sub

Private Sub mnuVBlackBack_Click()
If GetSetting("TabPad", "Settings", "BlackBack", "0") = "0" Then
    mnuVBlackBack.Checked = True 'Change checking
    SaveSetting "TabPad", "Settings", "BlackBack", "1" 'Save to reg
Else
    mnuVBlackBack.Checked = False 'Change checking
    SaveSetting "TabPad", "Settings", "BlackBack", "0" 'Save to reg
End If
LoadFSettings
End Sub

Private Sub mnuVDocInfo_Click()
MsgBox "Document '" & Tabs.SelectedItem.Caption & "'" & vbCrLf & vbCrLf & "Characters: " & Len(txtMain) & vbCrLf & "Lines: " & CharCount(txtMain.Text, vbCrLf) + 1 & vbCrLf & "Words: " & CharCount(txtMain.Text, " ") + 1
End Sub

Private Sub mnuVRandomize_Click()
Randomize
    Me.txtMain.BackColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
    Me.txtMain.ForeColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
End Sub

Private Sub mnuVSBRandTips_Click()

If GetSetting("TabPad", "Settings", "StatusBarRandomTips", "1") = "0" Then
    mnuVSBRandTips.Checked = True 'Change checking
    SaveSetting "TabPad", "Settings", "StatusBarRandomTips", "1" 'Save to reg
    StatusBarRandomTips = True 'Enable variable so when the user changes things in textbox we update it
    'StatusBar.Panels(1).Visible = True
Else
    mnuVSBRandTips.Checked = False 'Change checking
    SaveSetting "TabPad", "Settings", "StatusBarRandomTips", "0" 'Save to reg
    StatusBarRandomTips = False 'Disable variable (see above)
    'StatusBar.Panels(1).Visible = False
End If
LoadSettings
End Sub

Private Sub mnuVStatus_Click()
If GetSetting("TabPad", "Settings", "StatusBar", "1") = "0" Then
    mnuVStatus.Checked = True 'Change checking
    SaveSetting "TabPad", "Settings", "StatusBar", "1" 'Save to reg
    StatusBarOn = True 'Enable variable so when the user changes things in textbox we update it
    'Me.StatusBar.Visible = True
Else
    mnuVStatus.Checked = False 'Change checking
    SaveSetting "TabPad", "Settings", "StatusBar", "0" 'Save to reg
    StatusBarOn = False 'Disable variable (see above)
    'Me.StatusBar.Visible = False
End If
LoadSettings
'Form_Resize
End Sub

Private Sub optTabKeyCustom_Click()
If optTabKeyDefault.Value = True Then
    txtTabKeyCustAmount.Visible = False
Else
    txtTabKeyCustAmount.Visible = True
End If
End Sub

Private Sub optTabKeyDefault_Click()
If optTabKeyDefault.Value = True Then
    txtTabKeyCustAmount.Visible = False
Else
    txtTabKeyCustAmount.Visible = True
End If
End Sub

Private Sub Tabs_Click()
'New tab item clicked

If RetrieveMode = False Then
    TabInfo(LastTabIndex).Text = Me.txtMain.Text
    TabInfo(LastTabIndex).SelStart = Me.txtMain.SelStart
    TabInfo(LastTabIndex).SelLength = Me.txtMain.SelLength
    'Put text and positions into array, unless RetrieveMode = True
    'This occurs (=True) when closing a tab
End If
RetrieveMode = True
Me.txtMain = TabInfo(Tabs.SelectedItem.Index - 1).Text 'Put new tab's text into textbox
RetrieveMode = False
Me.txtMain.SetFocus
Me.txtMain.SelStart = TabInfo(Tabs.SelectedItem.Index - 1).SelStart 'Read position from last edit
Me.txtMain.SelLength = TabInfo(Tabs.SelectedItem.Index - 1).SelLength '..

LastTabIndex = Tabs.SelectedItem.Index - 1 'Change selected tab index (-1 because TabInfo and the Tabs control index scheme are different - described in another part of code)
End Sub

Private Sub Tabs_GotFocus()
Me.tmrCheckFocus = False
Me.tmrCheckFocus = True
'tmrCheckFocus ensures the Tabs_Click event takes place even if the tab is selected obscurely (eg. if it is clicked and dragged)
End Sub

Private Sub Tabs_LostFocus()
Me.tmrCheckFocus = False
'Don't need to check anymore
End Sub

Private Sub tmrAutoRecoverySave_Timer()
tmrAutoRecoverySave.Interval = 15000
Dim i As Integer
If GetSetting("TabPad", "Settings", "CrashRecovery", "1") = "1" Then
    CreatePathTree App.Path & "\TabPad-Recovery\"
    Open App.Path & "\TabPad-Recovery\Clear.dat" For Output As #1
    Close #1
    
    Kill App.Path & "\TabPad-Recovery\*.dat"
    
    'TabInfo(Tabs.SelectedItem.Index - 1).SelStart = txtMain.SelStart
    'TabInfo(Tabs.SelectedItem.Index - 1).SelLength = txtMain.SelLength
    Open App.Path & "\TabPad-Recovery\Session.dat" For Output As #5
        Print #5, Tabs.SelectedItem.Index
        For i = 0 To UBound(TabInfo)
            Print #5, Replace(TabInfo(i).FilePath, " ", "%20") & " " & Replace(Tabs.Tabs(i + 1).Caption, " ", "%20") & " " & TabInfo(i).SelStart & " " & TabInfo(i).SelLength
            'Open App.Path & "\TabPad-Recovery\Tab " & i & ".dat" For Output As #7
            Open App.Path & "\TabPad-Recovery\Tab " & i & ".dat" For Binary Access Write As #7
                If Tabs.SelectedItem.Index - 1 = i Then
                    Put #7, , txtMain.Text
                Else
                    Put #7, , TabInfo(i).Text
                End If
            Close #7
        Next i
    Close #5
    
    'Didn't work straight away, chucked it in as it was useless anyway
    'If ShowRecoverySaved = True Then
        'tmrShowRecoverySaved = True
        'ShowRecoverySaved = False
        'Me.StatusBar.Panels(3).Text = "Recovery Saved"
        'Me.tmrStatusBar.Enabled = False
        'Me.tmrStatusBar.Interval = 2000
        'Me.tmrStatusBar.Enabled = True
    'End If
End If
End Sub

Private Sub tmrCheckFocus_Timer()
'Explained in Tabs_GotFocus
tmrCheckFocus = False
'Don't need to check again because we'll be told when focus is received again
If Not Tabs.SelectedItem.Index - 1 = LastTabIndex Then
    'If the selected tab index isn't the same as the tab we have registered as being shown, click over to new tab
    Tabs_Click
    Exit Sub
End If
End Sub

Private Sub txtMain_Change()
'Raised by txtWrap and txtNoWrap to cut down on code usage
If RetrieveMode = True Then Exit Sub 'File is being loaded or wrap/nowrap mode being toggled
If txtMain.Text = "" And Right(Me.Tabs.SelectedItem.Caption, 3) = "  *" And TabInfo(LastTabIndex).FilePath = "" Then
    'If the textbox is empty, and the file hasn't been saved before (or loaded), yet it is noted as being unsaved, remove asterisk
    Me.Tabs.SelectedItem.Caption = Left(Me.Tabs.SelectedItem.Caption, Len(Me.Tabs.SelectedItem.Caption) - 3)
    Exit Sub
End If
'If Not txtMain.Text = "" And Not Right(Me.Tabs.SelectedItem.Caption, 3) = "  *" Then
If Not txtMain.Text = "" And Not Right(Me.Tabs.SelectedItem.Caption, 3) = "  *" Then
    'If text has changed and it isn't noted, add asterisk
    Me.Tabs.SelectedItem.Caption = Me.Tabs.SelectedItem.Caption & "  *"
End If
If txtMain.Text = "" And Not TabInfo(LastTabIndex).FilePath = "" And Not Right(Me.Tabs.SelectedItem.Caption, 3) = "  *" Then
    'If text has cleared and it isn't noted, add asterisk
    Me.Tabs.SelectedItem.Caption = Me.Tabs.SelectedItem.Caption & "  *"
End If

If GetSetting("TabPad", "Settings", "CrashRecovery", "1") = "1" Then
    'CreatePathTree App.Path & "\TabPad-Recovery\"
    'Open App.Path & "\TabPad-Recovery\Tab " & Tabs.SelectedItem.Index - 1 & ".dat" For Output As #7
        'Print #7, txtMain.Text
    'Close #7
    tmrAutoRecoverySave.Interval = 1000
    tmrAutoRecoverySave = False
    tmrAutoRecoverySave = True
End If
End Sub

Sub CheckTabs()
If Tabs.Tabs.Count = 1 Then
    'If there is only 1 tab
    Me.cmdCloseTab.Visible = False
    Me.mnuFCloseTab.Visible = False
    'Get rid of the close button/menu item - it's unneeded
Else
    'More than 1 tab
    Me.cmdCloseTab.Visible = True
    Me.mnuFCloseTab.Visible = True
    'Show close button and menu item
End If

'This code was for when the button slid along the tab control (buggy) :(
'Dim OverallWidth As Integer
'For i = 1 To Tabs.Tabs.Count
'    OverallWidth = OverallWidth + Tabs.Tabs(i).Width
'Next i

'If OverallWidth >= Tabs.Width - 150 Then
'    Me.cmdCloseTab.Left = Me.Width - 880
'Else
'    Me.cmdCloseTab.Left = Me.Width - 390
'End If

'Because tab control isn't fullsize, we slap it next to it
Me.cmdCloseTab.Left = Me.Width - 425
End Sub


Sub LoadSettings()
If GetSetting("TabPad", "Settings", "WordWrap", "1") = "1" Then
    'Word Wrap mode
    Me.mnuFWordWrap.Checked = True
    Set txtMain = txtWrap
    txtMain.Visible = True
    txtMain = TabInfo(LastTabIndex).Text
Else
    'No Word Wrap
    Me.mnuFWordWrap.Checked = False
    Set txtMain = txtNoWrap
    txtMain.Visible = True
    txtMain = TabInfo(LastTabIndex).Text
End If
'Me.mnuFAllowTabs.Checked = GetSetting("TabPad", "Settings", "TabAsSpaces", "0") 'Check or uncheck menu item

StatusBarOn = GetSetting("TabPad", "Settings", "StatusBar", "1")

Me.StatusBar.Visible = StatusBarOn
Me.mnuVStatus.Checked = StatusBarOn


StatusBarRandomTips = GetSetting("TabPad", "Settings", "StatusBarRandomTips", "1")

Me.StatusBar.Panels(1).Visible = StatusBarRandomTips
Me.mnuVSBRandTips.Checked = StatusBarRandomTips


If StatusBarOn = True Then
    Me.mnuVSBRandTips.Enabled = True
        If StatusBarRandomTips = True Then
            Me.mnuHRandomTip.Visible = True
            Me.mnuHBrk.Visible = True
        Else
            Me.mnuHRandomTip.Visible = False
            Me.mnuHBrk.Visible = False
        End If
Else
    Me.mnuHRandomTip.Visible = False
    Me.mnuHBrk.Visible = False
    Me.mnuVSBRandTips.Enabled = False
End If

If GetSetting("TabPad", "Settings", "BlackBack", "0") = "1" Then
    mnuVBlackBack.Checked = True
End If

LoadFSettings 'Load font settings
LoadTKSettings 'Load tab key settings

MaxRecentDocs = GetSetting("TabPad", "Settings", "RecentDocuments", "10")
If MaxRecentDocs > 32 Or MaxRecentDocs < 0 Then
    MaxRecentDocs = 10
End If

Form_Resize 'Resize our new text box :)
End Sub

Sub LoadTKSettings()
If GetSetting("TabPad", "Settings", "TabAsSpaces", "1") = "1" Then
    Me.cmdCloseTab.TabStop = False
    Me.Tabs.TabStop = False
    If GetSetting("TabPad", "Settings", "TabKeyAmount", "Default") = "Default" Then
        TabKeyMethod = 0
    Else
        TabKeyMethod = GetSetting("TabPad", "Settings", "TabKeyAmount", "4")
    End If
Else
    Me.cmdCloseTab.TabStop = True
    Me.Tabs.TabStop = True
End If
End Sub

Sub LoadFSettings()
Me.txtMain.FontName = GetSetting("TabPad", "Settings", "Font", Me.txtMain.FontName)
Me.txtMain.FontSize = GetSetting("TabPad", "Settings", "FontSize", Me.txtMain.FontSize)
Me.txtMain.FontBold = GetSetting("TabPad", "Settings", "FontBold", Me.txtMain.FontBold)

If GetSetting("TabPad", "Settings", "BlackBack", "0") Then
    Me.txtMain.BackColor = RGB(0, 0, 0)
    Me.txtMain.ForeColor = RGB(255, 255, 255)
Else
    Me.txtMain.BackColor = RGB(255, 255, 255)
    Me.txtMain.ForeColor = RGB(0, 0, 0)
End If
    
'Read settings from Registry into textbox properties
End Sub

Sub LoadFile(Path As String)
'Loads a file into a tab
On Error GoTo errTrap
Dim i As Integer
For i = 1 To Tabs.Tabs.Count
    'Check each tab to see if the file is already open
    If TabInfo(i - 1).FilePath = Path Then
        'We found it already open, jump to it
        Tabs.Tabs(i).Selected = True
        Exit Sub
    End If
Next i
RetrieveMode = True
mnuFNew_Click
'Make a new tab, if neccassary
RetrieveMode = False

If RecentDocs.Count = 0 Then
    RecentDocs.Add Path
Else
    RecentDocs.Add Path, , 1
End If
'Add the loading docuement to recent docs

reloopdel:
For i = 2 To RecentDocs.Count
    If RecentDocs.item(i) = Path Then
        RecentDocs.Remove i
        'Remove older dupe
        GoTo reloopdel
    End If
Next i
'Get rid of duplicates (older first)

Do Until RecentDocs.Count <= MaxRecentDocs
    RecentDocs.Remove RecentDocs.Count
Loop
'Make sure there are no more than 10 [Now MaxRecentDocs.. Up to 32] recents

PopulateRDMenu
'Fix up menu slideout

If FileLen(Path) >= 65536 Then 'Checks size of file to load (in bytes)
    'More than 64 kb.
    Me.Show
    txtMain = GetFileName(Path) & " is too large for TabPad to handle. The file has been opened in Notepad (assuming it is installed to the default location)."
    'Display error text
    Tabs.SelectedItem.Caption = GetFileName(Path) & " too large"
    'Fix title
    DoEvents
    'Make sure TabPad is all done before notepad is focused
    OpenInNP Path
    'Open it in Notepad
    Exit Sub
End If

Dim CurrLine As String
Dim WholeDoc As String
Open Path For Input As #1 'Open the path for Input (as opposed to Output)
    Do Until EOF(1) 'Do until we've reached the End Of File
        Line Input #1, CurrLine 'Retrieve one line at a time
        If Not WholeDoc = "" Then WholeDoc = WholeDoc & vbCrLf 'If our WholeDoc variable isn't empty, add a new line
        WholeDoc = WholeDoc & CurrLine '..and append string read
    Loop 'keep looping
Close #1 'close it when done

'Open Path For Binary Access Read As #1
    'WholeDoc = Space(LOF(1))
    'Get #1, , WholeDoc
'Close #1

TabInfo(Tabs.Tabs.Count - 1).FilePath = Path
TabInfo(Tabs.Tabs.Count - 1).Text = WholeDoc
'Set array settings
txtMain.Text = WholeDoc
'Load text into GUI
Dim TabTitle As String
TabTitle = GetFileName(Path)
If LCase(Right(TabTitle, 4)) = ".txt" Then
    'If the extension is .txt, we don't need to show it on tab's title
    TabTitle = Left(TabTitle, Len(TabTitle) - 4)
End If
Tabs.SelectedItem.Caption = TabTitle
'Display user-friendly tab title

'If Len(Path) <= 76 Then
    'Tabs.SelectedItem.ToolTipText = Path
'Else
    'Tabs.SelectedItem.ToolTipText = Left(Path, 3) & "..." & Right(Path, 72)
'End If

Exit Sub
errTrap:
'Failed to load file
MsgBox "Error loading " & GetFileName(Path) & ": (" & Err.Number & ") " & Err.Description
End Sub

Private Sub tmrLoadPH_Timer()
'We read the placeholder to see if new instances have been running
If Not FileLen(App.Path & "\TPLoadList.dat") = 0 Then
    'We shall LOAD the files :)
    Dim CurrLine As String
    Open App.Path & "\TPLoadList.dat" For Input As #2 'Open it for input (to read, not write)
        Do Until EOF(2) 'Do until we've reached the End Of File
            Line Input #2, CurrLine 'Read Line
            If CurrLine = "NEWBLANK" Then
                mnuFNew_Click 'New instance was called without command line arguements - give them a blank tab
            Else
                LoadFile CurrLine 'Otherwise we load their file
            End If
        Loop
    Close #2 'All done, close file
    MakePlaceHolder 'Reset it ;) (so that we don't load them over and over
    If Me.WindowState = 1 Then 'If it's minimized
        Me.WindowState = 0 'Make it so it's not
    End If
    Me.SetFocus 'Focus form
    Me.txtMain.SetFocus 'Focus textbox
    Me.ZOrder 'Put on top
End If
End Sub

Sub PopulateRDMenu()
'Gets recent documents from internal list for onscreen menu
Dim i As Integer
For i = 1 To Me.mnuFRecentList.UBound
    'Get rid of outstanding controls (from last PopulateRDMenu Call)
    Unload Me.mnuFRecentList(i)
Next i
If RecentDocs.Count = 0 Then 'There ARE no recent documents!
    Me.mnuFRecentList(0).Caption = "(none)"
    Me.mnuFRecentList(0).Enabled = False
    'Display an entry saying there are none
    Exit Sub
Else
    Me.mnuFRecentList(0).Caption = GetFileName(RecentDocs.item(1))
    Me.mnuFRecentList(0).Tag = RecentDocs.item(1)
    Me.mnuFRecentList(0).Enabled = True
    'We do our first recent document here because Array 0 menu item is perisistent
End If

For i = 1 To RecentDocs.Count - 1
    'We service all recent documents except the first
    Load Me.mnuFRecentList(i) 'Make a menu entry
    Me.mnuFRecentList(i).Caption = GetFileName(RecentDocs.item(i + 1)) 'Set the name of the file (viewer friendly)
    Me.mnuFRecentList(i).Tag = RecentDocs.item(i + 1) 'Complete path (for loading)
Next i

DeleteRRK 'Delete the registry keys
For i = 1 To RecentDocs.Count
    'So that we can rewrite them :)
    SaveSetting "TabPad", "Recent", "Recent" & i, RecentDocs(i) 'Write to registry
Next i
End Sub

Private Sub tmrRandomTip_Timer()
'Dim RandNum As Integer
'Randomize
'RandNum = Int(Rnd * RandomTips.Count) + 1

'Me.StatusBar.Panels(1).Text = "Random Tip: " & RandomTips.item(RandNum)

If ShuffledRandomTips.Count = 0 Then
    ShuffleCollection RandomTips, ShuffledRandomTips
End If

Me.StatusBar.Panels(1).Text = "Random Tip: " & ShuffledRandomTips.item(1)
ShuffledRandomTips.Remove 1
End Sub

Private Sub tmrStatusBar_Timer()
If StatusBarOn = True Then
    tmrStatusBar.Interval = 50
    'This is done here because on KeyPress and MouseUp events, different readings were being returned
    'Also home + end keys werent noticed
    If Me.WindowState = 1 Then Exit Sub
    Me.StatusBar.Panels(3).Text = "Line " & GetLineFromChar(txtMain.hwnd, -1) & ", Col " & GetColNum(txtMain)
End If
End Sub

Private Sub txtNoWrap_Change()
'Text changed, call single event
txtMain_Change
End Sub

Private Sub txtNoWrap_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtMain_MouseUp Button, Shift
End Sub

Private Sub txtWrap_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtMain_MouseUp Button, Shift
End Sub

Private Sub txtWrap_Change()
'Text changed, call single event
txtMain_Change
End Sub

Private Sub txtNoWrap_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
'Somebody dropped a file (or a few ;)) onto our document :)
Dim i As Integer
For i = 1 To Data.Files.Count
    LoadFile Data.Files(i) 'So we load each
Next i
End Sub

Private Sub txtWrap_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
'Somebody dropped a file (or a few ;)) onto our document :)
Dim i As Integer
For i = 1 To Data.Files.Count
    LoadFile Data.Files(i) 'So we load each
Next i
End Sub

Private Sub txtMain_Key(KeyCode As Integer, Optional Shift As Integer)
If KeyCode = 9 Then 'Tab Key
    If Not TabKeyMethod = 0 Then
        SendKeys "{backspace}"
        txtMain.SelText = txtMain.SelText & Space(TabKeyMethod)
    End If
End If
End Sub


Private Sub txtMain_MouseUp(Button As Integer, Shift As Integer)
Dim SelWord As String
Dim i As Long
Dim StartPos As Long
Dim EndPos As Long
Dim TempKS As Long
DoEvents
TempKS = GetAsyncKeyState(17) 'This just stops the keystate being one (meaning the key had been being pressed but isnt now)
If Not GetAsyncKeyState(17) = 0 And Button = 1 Then
    If txtMain.SelStart = 0 Then Exit Sub
    If txtMain.SelLength = Len(txtMain.Text) Then Exit Sub
    
    For i = txtMain.SelStart To 1 Step -1
        If Mid(txtMain.Text, i, 1) = " " Or Mid(txtMain.Text, i, 1) = vbLf Then
            StartPos = i + 1
            Exit For
        End If
    Next i
    
    For i = StartPos + 1 To Len(txtMain.Text)
        If Mid(txtMain.Text, i, 1) = " " Or Mid(txtMain.Text, i, 1) = vbCr Then
            EndPos = i
            Exit For
        End If
    Next i
    If EndPos = 0 Then EndPos = Len(txtMain.Text) + 1
    
    If StartPos = 0 Then StartPos = 1
    SelWord = Mid(txtMain.Text, StartPos, EndPos - StartPos)
    
    If SelWord = "" Then
        Exit Sub
    End If
    
    If Not LCase(Left(SelWord, 7)) = "http://" Then
        If MsgBox("You control-clicked a piece of text (" & SelWord & "). TabPad usually opens a browser to load the URL but this does not appear to be a URL. Would you like TabPad to make it look like one and try to load it?", vbYesNo + vbQuestion, "Open in webbrowser") = vbYes Then
            SelWord = "http://" & SelWord
        Else
            MsgBox "In the future try not to control-click items which you don't want to navigate to."
            Exit Sub
        End If
    End If
    
    
    ShellDef SelWord
End If
End Sub

Private Sub txtNoWrap_KeyPress(KeyAscii As Integer)
txtMain_Key KeyAscii
End Sub

Private Sub txtWrap_KeyPress(KeyAscii As Integer)
txtMain_Key KeyAscii
End Sub


Sub RecoverSession()
Dim i As Integer
Dim SelectedTab As Integer
Dim LineNum As Integer
Dim CLine As String
Dim TabInfoNum As Integer
Dim FText As String

Tabs.Tabs.Clear
RetrieveMode = True
Open App.Path & "\TabPad-Recovery\Session.dat" For Input As #5
    Do Until EOF(5)
        LineNum = LineNum + 1
        Line Input #5, CLine
        If LineNum = 1 Then
            SelectedTab = CLine
        Else
            'Replace(TabInfo(i).FilePath, " ", "%20") & " " & Replace(Tabs.Tabs(i + 1).Caption, " ", "%20") & " " & TabInfo(i).SelStart & " " & TabInfo(i).SelLength
            '
            'File%20Path Caption SelStart SelLength
            ReDim Preserve TabInfo(TabInfoNum)
            TabInfo(TabInfoNum).FilePath = Replace(Split(CLine, " ")(0), "%20", " ")
            Open App.Path & "\TabPad-Recovery\Tab " & TabInfoNum & ".dat" For Binary Access Read As #7
                FText = Space(LOF(7))
                Get #7, , FText
            Close #7
            'MsgBox FText
            TabInfo(TabInfoNum).Text = FText
            TabInfo(TabInfoNum).SelStart = Split(CLine, " ")(2)
            'MsgBox Split(CLine, " ")(2)
            TabInfo(TabInfoNum).SelLength = Split(CLine, " ")(3)
            Tabs.Tabs.Add , , Replace(Split(CLine, " ")(1) & "", "%20", " ")
            
            If TabInfoNum + 1 = SelectedTab Then
                txtMain.Text = FText
                txtMain.SelStart = TabInfo(TabInfoNum).SelStart
                txtMain.SelLength = TabInfo(TabInfoNum).SelLength
            End If
            
            TabInfoNum = TabInfoNum + 1
        End If
    Loop
Close #5

'For i = 0 To UBound(TabInfo)
    'MsgBox i & " = " & TabInfo(i).Text
'Next i

'MsgBox SelectedTab
Tabs.Tabs(SelectedTab).Selected = True
'Tabs_Click

RetrieveMode = False
End Sub
