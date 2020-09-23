VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "API Viewer"
   ClientHeight    =   4905
   ClientLeft      =   2175
   ClientTop       =   2220
   ClientWidth     =   6135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   1200
      Width           =   1095
   End
   Begin VB.PictureBox picListings 
      BorderStyle     =   0  'None
      Height          =   4200
      Left            =   510
      ScaleHeight     =   4200
      ScaleWidth      =   4125
      TabIndex        =   3
      Top             =   165
      Width           =   4120
      Begin VB.ListBox lstListing 
         Height          =   2010
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   325
         Width           =   4125
      End
      Begin VB.ListBox lstChosen 
         Height          =   1815
         Left            =   0
         TabIndex        =   6
         Top             =   2360
         Width           =   4125
      End
      Begin MSForms.ComboBox cmbSearch 
         Height          =   315
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   4125
         VariousPropertyBits=   142624795
         DisplayStyle    =   3
         MousePointer    =   3
         Size            =   "7276;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.PictureBox picFindError 
      BorderStyle     =   0  'None
      Height          =   4200
      Left            =   480
      ScaleHeight     =   4200
      ScaleWidth      =   4125
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4120
      Begin VB.CommandButton cmdGetLastError 
         Caption         =   "&Get Last Error"
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         ToolTipText     =   "Retrieves the value of the GetLastError API"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdFindError 
         Caption         =   "&Find"
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtErrNumber 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstErrors 
         Height          =   1230
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblErrorFound 
         BackStyle       =   0  'Transparent
         Height          =   795
         Left            =   165
         TabIndex        =   13
         Top             =   960
         Width           =   3885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Error No:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   270
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame fraScope 
      Caption         =   "Scope"
      Height          =   855
      Left            =   4920
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
      Begin VB.OptionButton optPrivate 
         Caption         =   "Pri&vate"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optPublic 
         Caption         =   "P&ublic"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ProgressBar pgbCopying 
      Height          =   195
      Left            =   4710
      TabIndex        =   4
      Top             =   4695
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer timResetStatus 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1800
      Top             =   2400
   End
   Begin MSComctlLib.TabStrip tabFormat 
      Height          =   4290
      Left            =   120
      TabIndex        =   1
      Top             =   105
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7567
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      Placement       =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Declares"
            Key             =   "ltDeclare"
            Object.ToolTipText     =   "Choose declare functions (Ctrl+D)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Types"
            Key             =   "ltType"
            Object.ToolTipText     =   "Choose types (Ctrl+T)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Constants"
            Key             =   "ltConstant"
            Object.ToolTipText     =   "Choose constants (Ctrl+O)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Enums"
            Key             =   "ltEnum"
            Object.ToolTipText     =   "Choose enumerations (Ctrl+E)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Find Error"
            Key             =   "ltFindError"
            Object.ToolTipText     =   "Find an error from its value (Ctrl+F)"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4650
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8202
            MinWidth        =   847
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Status"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit_AddItem 
         Caption         =   "&Add Item"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuEdit_RemoveItem 
         Caption         =   "&Remove Item"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_ClearItems 
         Caption         =   "C&lear  Items"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuEdit_Dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuView_ShowDeclares 
         Caption         =   "Show &Declares"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuView_ShowTypes 
         Caption         =   "Show &Types"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuView_ShowConstants 
         Caption         =   "Show &Constants"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuView_ShowEnums 
         Caption         =   "Show &Enums"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ListingType
    ltDeclare = 0
    ltType = 1
    ltConstant = 2
    ltEnum = 3
End Enum

Dim LastListing     As Long
Dim IsDelete        As Boolean

Private Sub cmbSearch_Change()

    Dim RetVal      As Long
    Dim CaratPos    As Long
    Static TxtLen   As Long

    If Len(cmbSearch.Text) > 0 Then
        RetVal = FindClosestEntry(lstListing, cmbSearch.Text)
        lstListing.ListIndex = RetVal
        cmbSearch.SetFocus
    End If

End Sub

Private Sub cmbSearch_KeyPress(KeyAscii As MSForms.ReturnInteger)

    If KeyAscii = vbKeyReturn Then lstListing_KeyPress (KeyAscii)
    
End Sub

Private Sub cmdAdd_Click()

    Dim FindText    As String
    Dim ListNo      As Long
    
    If lstListing.SelCount > 0 Then
        Select Case LastListing
            Case Is = ltDeclare
                FindText = "Declare: " & lstListing.List(lstListing.ListIndex)
            Case Is = ltType
                FindText = "Type: " & lstListing.List(lstListing.ListIndex)
            Case Is = ltConstant
                FindText = "Constant: " & lstListing.List(lstListing.ListIndex)
            Case Is = ltEnum
                FindText = "Enum: " & lstListing.List(lstListing.ListIndex)
        End Select
        
        For ListNo = 0 To lstChosen.ListCount - 1
            If lstChosen.List(ListNo) = FindText Then
                Beep
                Offence.DuplicateItem = Offence.DuplicateItem + 1
                stbStatus.Panels(1).Text = "Duplicate items not allowed"
                If Offence.DuplicateItem = 3 Then
                    MsgBox "Dupliacte items are not allowed", vbOKOnly + vbExclamation + vbApplicationModal, "API Viewer"
                    Offence.DuplicateItem = 1
                End If
                timResetStatus.Enabled = False
                timResetStatus.Enabled = True
                Exit Sub
            End If
        Next ListNo
        
        lstChosen.AddItem FindText
    End If
    
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Click to add selected item to the copy list"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub cmdClear_Click()

    lstChosen.Clear
    
End Sub

Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Click to clear all items in the copy list"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub cmdClose_Click()

    Set frmMain = Nothing
    Unload frmMain
    End
    
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Click to close API Viewer"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub cmdCopy_Click()

    Dim ListItem    As Long
    Dim TotalText   As String
    Dim Scope       As String
    
    If lstChosen.ListCount > 1 Then
        pgbCopying.Max = lstChosen.ListCount
        With pgbCopying
            .Value = 0
            .Visible = True
        End With
        stbStatus.Panels(2).Visible = True
    End If
    
    If optPrivate.Value = True Then
        Scope = "Private "
    Else
        Scope = "Public "
    End If
    
    For ListItem = 0 To lstChosen.ListCount - 1
        pgbCopying.Value = pgbCopying.Value + 1
        DoEvents
        Select Case Left(lstChosen.List(ListItem), 1)
            Case Is = "D"
                stbStatus.Panels(1).Text = "Finding " & Right(lstChosen.List(ListItem), Len(lstChosen.List(ListItem)) - Len("Declare:"))
                TotalText = TotalText & vbNewLine
                TotalText = TotalText & Scope & GrabDeclare(Right(lstChosen.List(ListItem), Len(lstChosen.List(ListItem)) - Len("Declare:")))
            Case Is = "T"
                stbStatus.Panels(1).Text = "Finding " & Right(lstChosen.List(ListItem), Len(lstChosen.List(ListItem)) - Len("Type:"))
                TotalText = TotalText & vbNewLine
                TotalText = TotalText & Scope & GrabType(Right(lstChosen.List(ListItem), Len(lstChosen.List(ListItem)) - Len("Type:")))
            Case Is = "E"
                stbStatus.Panels(1).Text = "Finding " & Right(lstChosen.List(ListItem), Len(lstChosen.List(ListItem)) - Len("Enum:"))
                TotalText = TotalText & vbNewLine
                TotalText = TotalText & Scope & GrabEnum(Right(lstChosen.List(ListItem), Len(lstChosen.List(ListItem)) - Len("Enum:")))
            Case Is = "C"
                stbStatus.Panels(1).Text = "Finding " & Right(lstChosen.List(ListItem), Len(lstChosen.List(ListItem)) - Len("Constant:"))
                TotalText = TotalText & vbNewLine
                TotalText = TotalText & Scope & GrabConstant(Right(lstChosen.List(ListItem), Len(lstChosen.List(ListItem)) - Len("Constant:")))
        End Select
    Next ListItem
    
    stbStatus.Panels(1).Text = "Copying..."
    
    If Right(TotalText, 2) = vbNewLine Then
        TotalText = Left(TotalText, Len(TotalText) - 2)
    End If
    
    Clipboard.Clear
    Clipboard.SetText TotalText
    
    With pgbCopying
        .Value = 0
        .Visible = False
        stbStatus.Panels(2).Visible = False
    End With
    stbStatus.Panels(2).Visible = False
    
    If lstChosen.ListCount > 1 Then
        stbStatus.Panels(1).Text = "Items copied"
    ElseIf lstChosen.ListCount = 1 Then
        stbStatus.Panels(1).Text = "Item copied"
    End If
    
    Call timResetStatus_Timer
    
End Sub

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Click to copy selected items"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub cmdFindError_Click()

    lblErrorFound.Caption = GrabError(txtErrNumber.Text)
    
End Sub

Private Sub cmdFindError_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Click to identify the chosen error number"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub cmdGetLastError_Click()

    Dim LastError   As Long
    
    LastError = GetLastError
    txtErrNumber.Text = LastError
    lblErrorFound.Caption = GrabError(txtErrNumber.Text)
    
End Sub

Private Sub cmdGetLastError_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Click to get the error number from the GetLastError API and identify it"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub cmdOptions_Click()

    frmOptions.Show vbModal, Me
    
End Sub

Private Sub cmdOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Click to display options for API Viewer"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub cmdRemove_Click()

    If lstChosen.SelCount > 0 Then
        lstChosen.RemoveItem lstChosen.ListIndex
    End If
    
End Sub

Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Click to remove the selected item from the copy list"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub Form_Load()
    
    LastListing = ltDeclare

    DoEvents
    
    LastListing = -1
    
    Call frmOptions.GetOptions(True)
    Call frmOptions.ImplyOptions
    Set tabFormat.SelectedItem = tabFormat.Tabs(1)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Me.optPrivate.Value = True Then
        SaveSetting App.Title, "Settings", "Scope", "Private"
    Else
        SaveSetting App.Title, "Settings", "Scope", "Public"
    End If
    
    End
    
End Sub

Private Sub lstChosen_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then cmdRemove_Click
    
End Sub

Private Sub lstChosen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "List of items to copy"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub lstListing_Click()

    If lstListing.ListIndex = 1 Then Exit Sub
    
    If cmbSearch.Text <> lstListing.List(lstListing.ListIndex) Then
        cmbSearch.Text = lstListing.List(lstListing.ListIndex)
        cmbSearch.SelStart = 0
        cmbSearch.SelLength = Len(cmbSearch.Text)
        cmbSearch.SetFocus
    End If
    
End Sub

Private Sub lstListing_DblClick()

    Call cmdAdd_Click
    
End Sub

Private Sub lstListing_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then cmdAdd_Click
    
End Sub

Private Sub lstListing_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "List of APIs to choose from"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub mnuEdit_AddItem_Click()

    Call cmdAdd_Click
    
End Sub

Private Sub mnuEdit_ClearItems_Click()

    Call cmdClear_Click
    
End Sub

Private Sub mnuEdit_Copy_Click()

    Call cmdCopy_Click
    
End Sub

Private Sub mnuEdit_RemoveItem_Click()

    Call cmdRemove_Click
    
End Sub

Private Sub mnuHelp_About_Click()

    frmAbout.Show vbModal, Me
    
End Sub

Private Sub mnuView_ShowConstants_Click()

    Set tabFormat.SelectedItem = tabFormat.Tabs(3)
    Call tabFormat_Click
    
End Sub

Private Sub mnuView_ShowDeclares_Click()

    mnuView_ShowDeclares.Checked = True
    mnuView_ShowTypes.Checked = False
    mnuView_ShowConstants.Checked = False
    mnuView_ShowEnums.Checked = False
    Set tabFormat.SelectedItem = tabFormat.Tabs(1)
    Call tabFormat_Click
    
End Sub

Private Sub mnuView_ShowEnums_Click()

    mnuView_ShowDeclares.Checked = False
    mnuView_ShowTypes.Checked = False
    mnuView_ShowConstants.Checked = False
    mnuView_ShowEnums.Checked = True
    Set tabFormat.SelectedItem = tabFormat.Tabs(4)
    Call tabFormat_Click
    
End Sub

Private Sub mnuView_ShowTypes_Click()

    mnuView_ShowDeclares.Checked = False
    mnuView_ShowTypes.Checked = True
    mnuView_ShowConstants.Checked = False
    mnuView_ShowEnums.Checked = False
    Set tabFormat.SelectedItem = tabFormat.Tabs(2)
    Call tabFormat_Click
    
End Sub

Private Sub optPrivate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Set copy scope to private"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub optPublic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    stbStatus.Panels(1).Text = "Set copy scope to public"
    timResetStatus.Enabled = False
    timResetStatus.Interval = 2500
    timResetStatus.Enabled = True
    
End Sub

Private Sub tabFormat_Click()

    Dim SelKey  As String
    
    If LastListing = tabFormat.SelectedItem.Index - 1 Then Exit Sub

    lstListing.Clear
    cmbSearch.Clear
    
    SelKey = tabFormat.SelectedItem.Key
    
    Select Case SelKey
        Case Is = "ltDeclare"
            LastListing = ltDeclare
            mnuView_ShowDeclares.Checked = True
            mnuView_ShowTypes.Checked = False
            mnuView_ShowConstants.Checked = False
            mnuView_ShowEnums.Checked = False
            picListings.Visible = True
            picListings.Enabled = True
            picFindError.Visible = False
            picFindError.Enabled = False
            Call ShowDeclares(lstListing, cmbSearch)
        Case Is = "ltConstant"
            LastListing = ltConstant
            mnuView_ShowDeclares.Checked = False
            mnuView_ShowTypes.Checked = False
            mnuView_ShowConstants.Checked = True
            mnuView_ShowEnums.Checked = False
            picListings.Visible = True
            picListings.Enabled = True
            picFindError.Visible = False
            picFindError.Enabled = False
            Call ShowConstants(lstListing, cmbSearch)
        Case Is = "ltType"
            LastListing = ltType
            mnuView_ShowDeclares.Checked = False
            mnuView_ShowTypes.Checked = True
            mnuView_ShowConstants.Checked = False
            mnuView_ShowEnums.Checked = False
            picListings.Visible = True
            picListings.Enabled = True
            picFindError.Visible = False
            picFindError.Enabled = False
            Call ShowTypes(lstListing, cmbSearch)
        Case Is = "ltEnum"
            LastListing = ltEnum
            mnuView_ShowDeclares.Checked = False
            mnuView_ShowTypes.Checked = False
            mnuView_ShowConstants.Checked = False
            mnuView_ShowEnums.Checked = True
            picListings.Visible = True
            picListings.Enabled = True
            picFindError.Visible = False
            picFindError.Enabled = False
            Call ShowEnums(lstListing, cmbSearch)
        Case Is = "ltFindError"
            mnuView_ShowDeclares.Checked = False
            mnuView_ShowTypes.Checked = True
            mnuView_ShowConstants.Checked = False
            mnuView_ShowEnums.Checked = False
            picListings.Visible = False
            picListings.Enabled = False
            picFindError.Visible = True
            picFindError.Enabled = True
            Call ShowFindError(lstListing, cmbSearch)
    End Select

End Sub

Private Sub timResetStatus_Timer()

    stbStatus.Panels(1).Text = "Ready"
    timResetStatus.Enabled = False
    
End Sub
