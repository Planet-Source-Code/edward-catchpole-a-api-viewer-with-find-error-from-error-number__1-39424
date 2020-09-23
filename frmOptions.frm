VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1470
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   3855
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUseAutocomplete 
      Caption         =   "Use autocomplete"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkShowParam 
      Caption         =   "Show parameter in constant name"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   975
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   975
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   975
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    
    Call SaveOptions
    Call ImplyOptions
    
End Sub

Private Sub cmdCancel_Click()

    Me.Hide
    
End Sub

Private Sub cmdOK_Click()

    Call SaveOptions
    Call ImplyOptions
    Me.Hide
    
End Sub

Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
End Sub

Public Sub SaveOptions()

    SaveSetting App.Title, "Settings", "ShowNameParam", chkShowParam.Value
    SaveSetting App.Title, "Settings", "UseAutocomplete", chkUseAutocomplete.Value

End Sub

Public Sub GetOptions(IsStartup As Boolean)

    Dim ScopeStr    As String
    
    If IsStartup = True Then
        ScopeStr = GetSetting(App.Title, "Settings", "Scope", "Private")
        Select Case LCase(ScopeStr)
            Case Is = "public"
                frmMain.optPublic.Value = True
            Case Else
                frmMain.optPrivate.Value = True
        End Select
    End If
    
    chkShowParam.Value = GetSetting(App.Title, "Settings", "ShowNameParam", 1)
    chkUseAutocomplete = GetSetting(App.Title, "Settings", "Autocomplete", 1)
    
End Sub

Public Sub ImplyOptions()

    Select Case chkUseAutocomplete.Value
        Case Is = vbChecked
            frmMain.cmbSearch.MatchEntry = fmMatchEntryComplete
        Case Is = vbUnchecked
            frmMain.cmbSearch.MatchEntry = fmMatchEntryNone
    End Select
    
End Sub
