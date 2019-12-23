Attribute VB_Name = "Import_Points"
Public wb As Workbook
Public ws As Worksheet
Public eventWs As Worksheet

' Password (Change here)
Public Const password As String = "password"
' Location of Point Tracking sheet and Sign-In Sheet (Change here)
Public Const pointSheet As String = "C:\Users\a2_mi\Personal\Tau Beta Pi Points Sheet\Point Tracking Sheet.xlsm"

' Public variables for event specific details
Public ePoints As Integer    ' Points for the event
Public eName As String      ' Name of the event
Public EDate As String      ' Date of the event
Public eType As String      ' Type of the event

Public col As Integer       ' Column of event

Public status As Integer    ' Status of import: 0 -> GOOD/SUCCESS, 1 -> STOP/FAILURE, 2 -> SUCCESS BUT MISMATCH
Public mismatch As Integer  ' Number of unmatched sign-ins

' Subroutine to open the Point Tracking Sheet workbook
Sub Open_Wb()
    ' Open the point tracking sheet
    Set wb = Workbooks.Open(pointSheet)
    Set wb = ActiveWorkbook
End Sub

' Subroutine to save and close out the Point Tracking Sheet workbook
Sub Save_Wb()
    ' Save the file if changes were made
    If status Mod 2 = 0 Then
        wb.Save
    End If
    ' Only close the worbook if no errors
    If status = 0 Then
        wb.Close
    End If
End Sub

' Subroutine to record points for everyone on the sign-up sheet
Sub Move_Points()
    ' Flag variable, record found
    Dim found As Integer
    
    ' Sign-in sheet variables
    Dim rowCount As Integer
    Dim row As Integer
    Dim firstName As String
    Dim lastName As String
    Dim netid As String
    
    'Point tracking sheet variables
    Dim pRowRange As Range
    Dim pRow
    Dim LastRow As Integer
    
    mismatch = 0
    
    ' Count the number of sign-in entries
    Workbooks("Sign-In Sheet.xlsm").Activate
    eventWs.Select
    
    rowCount = Range("A1").End(xlDown).row
    
    ' Iterate through each sign-in entry and check info
    For row = 3 To rowCount
        found = 0
        ' Reactivate the sign-in sheet
        Workbooks("Sign-In Sheet.xlsm").Activate
        eventWs.Select
        
        ' Check the student's name and netid
        firstName = Cells(row, 1).Value
        lastName = Cells(row, 2).Value
        netid = Cells(row, 3).Value
        
    ' Open the appropriate tab
        ' Check if entry is a member
        If Cells(row, 4).Value = "M" Or Cells(row, 4).Value = "m" Then
            Workbooks("Point Tracking Sheet.xlsm").Activate
            Worksheets(3).Activate
        ' If entry is an intitiate
        Else
            ' Check if graduate student
            If Cells(row, 5).Value = "G" Or Cells(row, 5).Value = "g" Then
                Workbooks("Point Tracking Sheet.xlsm").Activate
                Worksheets(2).Activate
            ' If undergrad
            Else
                Workbooks("Point Tracking Sheet.xlsm").Activate
                Worksheets(1).Activate
            End If
        End If
        
    ' Add information
        If Range("A4").Value = "" Then
            LastRow = 4
        ElseIf Range("A5").Value = "" Then
            LastRow = 5
        Else
            LastRow = Range("A4").End(xlDown).row
        End If
        
        ' Iterate through each row to find the matching name
        For pRow = 4 To LastRow
            If lastName = Cells(pRow, 1).Value And netid = Cells(pRow, 4).Value Then
                Cells(pRow, col).Value = ePoints
                
                ' Check for the type of event
                If eType = "Social" Then
                    Cells(pRow, 10).Value = Cells(pRow, 10).Value + ePoints
                ElseIf eType = "Professional" Then
                    Cells(pRow, 11).Value = Cells(pRow, 11).Value + ePoints
                Else
                    Cells(pRow, 12).Value = Cells(pRow, 12).Value + ePoints
                End If
                
                ' Exit the loop
                found = 1
                Exit For
            End If
        Next pRow
        
        ' If no matching entry found, create a list of the mismatches
        If found = 0 Then
            Cells(LastRow + 1, 1).Value = lastName
            Cells(LastRow + 1, 2).Value = firstName
            Cells(LastRow + 1, 3).Value = netid
            Cells(LastRow + 1, 4).Value = eType
            Cells(LastRow + 1, 5).Value = ePoints
            
            mismatch = mismatch + 1
        End If
    Next row
End Sub

' Subroutine to transfer data to the appropriate sheets
Sub Import()
    ' Set import status
    status = 0
    
    Set eventWs = ThisWorkbook.Worksheets(ActiveSheet.Name)
   
    ' Prompt user with password
    Call Event_Data.Password_Check
    
    If status = 0 Then
        ' Activate the point tracking workbook
        Call Import_Points.Open_Wb
    
        ' Fill the point tracking sheet with event details
        Call Event_Data.Open_Form
        
        ' Record points
        Call Move_Points
        
        If status = 0 Then
            If mismatch = 0 Then
                ' If no mismatches display confirmation
                MsgBox "Import Successful", vbInformation, "Success"
            Else
                ' If a name signed up does not match list, display the number of mismatches
                MsgBox "Import Successful. " & mismatch & " names were not found. The Names are listed at the bottom of the Point Tracking Sheet.", vbExclamation, "Success with Errors"
                status = 2
            End If
        End If
    Else
        ' Display cancellation message
        MsgBox "Import Canceled.", vbCritical, "Canceled"
    End If
    ' Save and close the point tracking sheet and display a message
        Call Save_Wb
End Sub
