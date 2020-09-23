Attribute VB_Name = "modFunc"
Option Explicit

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Type Offences
    DuplicateItem           As Long
End Type

Public Offence  As Offences

Public Function GrabDeclare(sName As String) As String

    On Error GoTo HandleErr
    
    Dim RecordNo    As Long
    Dim FoundRecord As Long
    
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    
    strSQL = "SELECT Name FROM Declares"
    rstListing.Open strSQL, conADO
    
    Do Until Trim(rstListing("Name")) = Trim(sName)
         rstListing.MoveNext
         RecordNo = RecordNo + 1
         If RecordNo > rstListing.RecordCount Then GoTo HandleErr
    Loop
    
    Call Data_Close
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Text FROM Declares"
    rstListing.Open strSQL, conADO
    
    rstListing.MoveFirst
    Call rstListing.Move(RecordNo)
    
    GrabDeclare = rstListing("Text")
    
    Call Data_Close
    
    Exit Function
    
HandleErr:
    
    GrabDeclare = "There was an error finding " & sName
    
    Call Data_Close
    
End Function

Public Function GrabConstant(sName As String) As String

    Dim Character   As String
    
    Character = LCase(Left(Trim(sName), 1))
    
    Select Case Character
        Case "a" To "m"
            GrabConstant = GrabConstant1stHalf(sName)
        Case "n" To "z"
            GrabConstant = GrabConstant2ndHalf(sName)
    End Select
    
End Function

Public Function GrabEnum(sName As String) As String

    On Error GoTo HandleErr
    
    Dim RecordNo    As Long
    Dim FoundRecord As Long
    
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    
    strSQL = "SELECT Name FROM Enums"
    rstListing.Open strSQL, conADO
    
    Do Until Trim(rstListing("Name")) = Trim(sName)
         rstListing.MoveNext
         RecordNo = RecordNo + 1
         If RecordNo > rstListing.RecordCount Then GoTo HandleErr
    Loop
    
    Call Data_Close
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Param FROM Enums"
    rstListing.Open strSQL, conADO
    
    rstListing.MoveFirst
    Call rstListing.Move(RecordNo)
    
    GrabEnum = "Enum" & sName & vbNewLine & AddSpaces(rstListing("Param")) & "End Enum"
    
    Call Data_Close
    
    Exit Function

HandleErr:
    
    GrabEnum = "There was an error finding " & sName
    
    Call Data_Close
    
End Function

Public Function GrabType(sName As String) As String

    On Error GoTo HandleErr
    
    Dim RecordNo    As Long
    Dim FoundRecord As Long
    
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    
    strSQL = "SELECT Name FROM Types"
    rstListing.Open strSQL, conADO
    
    Do Until Trim(rstListing("Name")) = Trim(sName)
         rstListing.MoveNext
         RecordNo = RecordNo + 1
         If RecordNo > rstListing.RecordCount Then GoTo HandleErr
    Loop
    
    Call Data_Close
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Param FROM Types"
    rstListing.Open strSQL, conADO
    
    rstListing.MoveFirst
    Call rstListing.Move(RecordNo)
    
    GrabType = "Type" & sName & vbNewLine & AddSpaces(rstListing("Param")) & "End Type"
    
    Call Data_Close
    
    Exit Function

HandleErr:
    
    GrabType = "There was an error finding " & sName
    
    Call Data_Close
    
End Function

Public Function AddSpaces(sData As String) As String

    Dim TempStr     As String
    Dim CharPos     As Long
    Dim TempData    As String
    
    TempData = sData
    
    Do
        CharPos = InStr(TempData, Chr(10))
        If CharPos = 0 Then Exit Do
        TempStr = TempStr & "    " & Left(TempData, CharPos - 1) & vbNewLine
        TempData = Right(TempData, Len(TempData) - CharPos)
    Loop
    
    AddSpaces = TempStr
    
End Function

Public Function FindClosestEntry(listbx As ListBox, strPrefix As String) As Long

    Dim RetVal      As Long
    Dim strEntry    As String
    Dim EntryNo     As Long
    
    RetVal = -1
    strEntry = strPrefix
    
    Do
        For EntryNo = 0 To listbx.ListCount - 1
            If Left(LCase(listbx.List(EntryNo)), Len(strEntry)) = LCase(strEntry) Then
                RetVal = EntryNo
                Exit For
            End If
        Next EntryNo
        If Len(strEntry) = 1 Then
            FindClosestEntry = -1
            Exit Function
        End If
        strEntry = Left(strEntry, Len(strEntry) - 1)
    Loop While RetVal = -1
    
    FindClosestEntry = RetVal

End Function

Public Function GrabConstant1stHalf(sName As String) As String

    On Error GoTo HandleErr

    Dim RecordNo    As Long
    Dim FoundRecord As Long
    Dim FieldName   As String
    
    Call Data_Connect
    
    If frmOptions.chkShowParam.Value = True Then
        FieldName = "NameParam"
    Else
        FieldName = "Name"
    End If
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT " & FieldName & " FROM Constants"
    rstListing.Open strSQL, conADO
    
    Do Until Trim(rstListing(FieldName)) = Trim(sName)
         rstListing.MoveNext
         RecordNo = RecordNo + 1
         If RecordNo > rstListing.RecordCount Then GoTo HandleErr
    Loop
    
    Call Data_Close
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Param FROM Constants"
    rstListing.Open strSQL, conADO
    
    rstListing.MoveFirst
    Call rstListing.Move(RecordNo)
    
    GrabConstant1stHalf = rstListing("Param")
    
    Call Data_Close
    
    Exit Function
    
HandleErr:
    
    GrabConstant1stHalf = "There was an error finding " & sName
    
    Call Data_Close
    
End Function

Public Function GrabConstant2ndHalf(sName As String) As String

    On Error GoTo HandleErr

    Dim RecordNo    As Long
    Dim FoundRecord As Long
    Dim FieldName   As String
    
    Call Data_Connect
    
    If frmOptions.chkShowParam.Value = True Then
        FieldName = "NameParam"
    Else
        FieldName = "Name"
    End If
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT " & FieldName & " FROM Constants2"
    rstListing.Open strSQL, conADO
    
    Do Until Trim(rstListing(FieldName)) = Trim(sName)
         rstListing.MoveNext
         RecordNo = RecordNo + 1
         If RecordNo > rstListing.RecordCount Then GoTo HandleErr
    Loop
    
    Call Data_Close
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Param FROM Constants2"
    rstListing.Open strSQL, conADO
    
    rstListing.MoveFirst
    Call rstListing.Move(RecordNo)
    
    GrabConstant2ndHalf = rstListing("Param")
    
    Call Data_Close
    
    Exit Function
    
HandleErr:
    
    GrabConstant2ndHalf = "There was an error finding " & sName
    
    Call Data_Close
    
End Function

Public Function GrabError(sName As String) As String

    On Error GoTo HandleErr

    Dim RecordNo    As Long
    Dim FoundRecord As Long
    Dim FieldName   As String
    
    Call Data_Connect
    
    FieldName = "ErrNumber"
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT " & FieldName & " FROM Errors"
    rstListing.Open strSQL, conADO
    
    Do Until Trim(rstListing(FieldName)) = Trim(sName)
         rstListing.MoveNext
         RecordNo = RecordNo + 1
         If RecordNo > rstListing.RecordCount Then GoTo HandleErr
    Loop
    
    Call Data_Close
    Call Data_Connect
    
    Set rstListing = New ADODB.Recordset
    strSQL = "SELECT Name FROM Errors"
    rstListing.Open strSQL, conADO
    
    rstListing.MoveFirst
    Call rstListing.Move(RecordNo)
    
    GrabError = "Error identified as: " & vbNewLine & rstListing("Name")
    
    Call Data_Close
    
    Exit Function
    
HandleErr:
    
    GrabError = "Error " & sName & " not found"
    
    Call Data_Close

End Function
