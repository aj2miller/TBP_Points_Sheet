VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportPoints 
   Caption         =   "Point Information"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3900
   OleObjectBlob   =   "ImportPoints.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Subroutine to cancel the import
Private Sub Cancel_Click()
'Clicking "Cancel" closes the form
    Unload ImportPoints
    status = 1
End Sub

' Subroutine to save all event specifics as public variables
Private Sub Import_Click()
    Dim errors As Integer
    errors = 0
    
' Check for issues in entered specifics
    If ImportPoints.evName.Value = "" Then
        MsgBox ("Please enter name of event")
        errors = errors + 1
    End If
    If ImportPoints.Points.Value = "" Then
        MsgBox ("Please enter the point amount")
        errors = errors + 1
    End If
    If ImportPoints.SocEvent.Value = False And ImportPoints.ProfEvent = False And ImportPoints.ServEvent.Value = False Then
        MsgBox ("Please enter the type of event")
        errors = errors + 1
    End If
  
' Save details as public variables if no errors
    If errors = 0 Then
        eName = ImportPoints.evName.Value
        EDate = ImportPoints.evDate.Value
        Dim Poin As String
        Poin = ImportPoints.Points.Value
        ePoints = CInt(Poin)
        
        If SocEvent = True Then
            eType = "Social"
        ElseIf ProfEvent = True Then
            eType = "Professional"
        Else
            eType = "Service"
        End If
        'Close the form
        Unload ImportPoints
        
' Errors
    Else
        status = 1
        Exit Sub
    End If
End Sub
