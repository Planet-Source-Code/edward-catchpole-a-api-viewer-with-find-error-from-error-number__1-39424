Attribute VB_Name = "modShow"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F

Public conADO       As ADODB.Connection
Public rstListing   As ADODB.Recordset
Public strSQL       As String

Public Const DataFile = "apitext.mdb"

Private Function DataSource() As String

    If App.Path = "\" Then
       DataSource = App.Path & DataFile
    Else
       DataSource = App.Path & "\" & DataFile
    End If
   
End Function

Public Sub Data_Connect()

    strSQL = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source = " & DataSource
    
    Set conADO = New ADODB.Connection
    conADO.CursorLocation = adUseClient
    conADO.Open strSQL

End Sub

Public Sub ShowDeclares(lstList As ListBox, cmbList As Object, Optional strType As String = "")

    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Name FROM Declares"
    rstListing.Open strSQL, conADO
    
    Do Until rstListing.EOF
         lstList.AddItem strType & rstListing("Name")
         cmbList.AddItem rstListing("Name")
         rstListing.MoveNext
    Loop
    
    Call Data_Close
        
End Sub

Public Sub ShowTypes(lstList As ListBox, cmbList As Object, Optional strType As String = "")

    Call Data_Connect

    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Name FROM Types"
    rstListing.Open strSQL, conADO
    
    Do Until rstListing.EOF
         lstList.AddItem strType & rstListing("Name")
         cmbList.AddItem rstListing("Name")
         rstListing.MoveNext
    Loop
    
    Call Data_Close
        
End Sub

Public Sub ShowEnums(lstList As ListBox, cmbList As Object, Optional strType As String = "")

    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Name FROM Enums"
    rstListing.Open strSQL, conADO
    
    Do Until rstListing.EOF
         lstList.AddItem strType & rstListing("Name")
         cmbList.AddItem rstListing("Name")
         rstListing.MoveNext
    Loop
    
    Call Data_Close
        
End Sub

Public Sub Data_Close()

    rstListing.Close
    Set rstListing = Nothing
    Set conADO = Nothing

End Sub

Public Sub ShowConstants(lstList As ListBox, cmbList As Object, Optional strType As String = "")

    Dim FieldName   As String
    Dim TblName     As String
    
    TblName = "Constants"
    
DoNext:
    Call Data_Connect
    
    If frmOptions.chkShowParam.Value = True Then
        FieldName = "NameParam"
    Else
        FieldName = "Name"
    End If
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT " & FieldName & " FROM " & TblName
    rstListing.Open strSQL, conADO
    
    Do Until rstListing.EOF
         lstList.AddItem strType & rstListing(FieldName)
         cmbList.AddItem rstListing(FieldName)
         rstListing.MoveNext
    Loop
    
    Call Data_Close
    
    If TblName = "Constants" Then
        TblName = "Constants2"
        GoTo DoNext
    End If
    
End Sub

Public Sub ShowFindError(lstList As ListBox, cmbList As Object, Optional strType As String = "")

    Dim FieldName   As String
    
    Call Data_Connect
    
    FieldName = "ErrNumber"
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT " & FieldName & " FROM Errors"
    rstListing.Open strSQL, conADO
    
    Do Until rstListing.EOF
         lstList.AddItem strType & rstListing(FieldName)
         cmbList.AddItem rstListing(FieldName)
         rstListing.MoveNext
    Loop
    
    Call Data_Close

End Sub
