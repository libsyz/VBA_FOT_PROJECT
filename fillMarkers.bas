Attribute VB_Name = "fillMarkers"


Option Explicit

'Pseudo Code
Dim currentExercise As String
Dim currentExerciseRow As Integer
Dim currentExerciseColumn As Integer

Dim currentCompetency As String
Dim targetRow As Integer
Dim targetColumn As Integer

Dim markers() As String

Dim ExIter As Integer
Dim CompIter As Integer
Dim i As Integer



Sub fillMarkers()
    ReDim markers(1 To 1) As String

    
    For ExIter = 0 To 4
    Worksheets("2-Do EX-C Matrix").Activate
    Range("C8").Select
    ActiveCell.Offset(0, ExIter).Select
    currentExercise = ActiveCell.Value
    If currentExercise <> "" Then
        
    
    
    currentExerciseColumn = ActiveCell.Column
    currentExerciseRow = ActiveCell.Row
    
        For CompIter = 1 To 4
        Worksheets("2-Do EX-C Matrix").Activate
        Cells(currentExerciseRow, currentExerciseColumn).Select
        
        ActiveCell.Offset(CompIter, 0).Select
        currentCompetency = ActiveCell.Value
        Debug.Print (currentCompetency)
        Debug.Print (currentExercise)
        
        'Find where the markers are
        'Go to Library Sheet
        Worksheets("Marker Library Simulations").Activate
        findLibraryRow (currentCompetency)
        Range("A1").Select
        Debug.Print (targetRow)
        findLibraryColumn (currentExercise)
        Call SetStartPoint
        
        'Collect Behavioral Markers
        Call collectMarkers
        
        
        'Go to test Sheet
        Worksheets("test").Activate
        findTestRow (currentExercise)
        findTestColumn (currentCompetency)
        Call SetStartPoint
        
        
        'Add those mutherfucking markers 6 freaking times
        Call createMarkerDropdowns
        
    
        
        Next
        
    End If
    
    Next
    
    
    
    


End Sub


Sub findLibraryRow(arg As String)
    Range("A1").Select
    While ActiveCell.Value <> arg
        Debug.Print ("competency not found, moving down")
        ActiveCell.Offset(1, 0).Select
    Wend
    targetRow = ActiveCell.Row
    Debug.Print (targetRow)

End Sub


Sub findLibraryColumn(arg As String)
    Range("A2").Select
    While ActiveCell.Value <> arg
        Debug.Print ("exercise not found, moving right")
        ActiveCell.Offset(0, 1).Select
    Wend
    targetColumn = ActiveCell.Column
    Debug.Print (targetColumn)

End Sub


Sub findTestRow(arg As String)
    Range("A1").Select
    While ActiveCell.Value <> arg
        Debug.Print ("exercise not found, moving down")
        ActiveCell.Offset(1, 0).Select
    Wend
    targetRow = ActiveCell.Row
    Debug.Print (targetRow)

End Sub


Sub findTestColumn(arg As String)
    Range("A1").Select
    ActiveCell.Offset(targetRow - 2, 0).Select
    While ActiveCell.Value <> arg
        Debug.Print ("competency not found, moving right")
        ActiveCell.Offset(0, 1).Select
    Wend
    targetColumn = ActiveCell.Column
    Debug.Print (targetColumn)

End Sub


Sub SetStartPoint()
    Debug.Print (targetRow)
    Debug.Print (targetColumn)
    Cells(targetRow, targetColumn).Select
End Sub


Sub collectMarkers()
    While ActiveCell.Value <> ""
        markers(UBound(markers)) = ActiveCell.Value
        ReDim Preserve markers(1 To UBound(markers) + 1) As String
        ActiveCell.Offset(1, 0).Select
    Wend
End Sub
 

Sub createMarkerDropdowns()
    For i = 1 To 6
     With ActiveCell.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, Formula1:=Join(markers, ",")
     End With
     ActiveCell.Offset(1, 0).Select
    Next
    ReDim markers(1 To 1) As String
    

End Sub
