Attribute VB_Name = "Event_Data"

' Subroutine to open a userform for event specifics
Sub Open_Form()
    ' Add options to the list box "Points"
    With ImportPoints.Points
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
    End With
    ' Open the form
    ImportPoints.Show
    
    If status = 0 Then
        ' Import event details
        Call Event_Data
    End If
End Sub

' Subroutine to ensure import is intentional
Sub Password_Check()
    Dim userPass As String
    
    ' Prompt user with a password
    userPass = InputBox("Please enter password", "Password Prompt", "*************************")
    If userPass = password Then
        MsgBox "Access Granted", vbInformation, "Accepted"
        
    Else
        MsgBox "Acces Denied", vbCritical, "Error"
        status = 1
    End If
End Sub


' Subroutine to copy the event details onto the points sheet
Sub Event_Data()
    ' Worksheet index
    Dim i As Integer
    i = 1
    
    ' Find the next available column
    col = 1 + ActiveSheet.Cells(1, 1).End(xlToRight).Column
      
    ' Loop through the 3 tabs (undergrad, grad, members)
    For Each ws In ActiveWorkbook.Sheets
        Worksheets(i).Activate
        
        ' Enter event info into the appropriate cells
        Cells(1, col).Value = eName
        Cells(2, col).Value = EDate
        Cells(3, col).Value = eType
        
        ' Increment worksheet index
        i = i + 1
    Next ws
End Sub

