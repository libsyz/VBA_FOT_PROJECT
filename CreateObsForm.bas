Attribute VB_Name = "CreateObsForm"
Option Explicit

Dim wordApp As Word.Application
Dim currentCompetency As String
Dim currentExercise As String
Dim currentCompetencyDescription As String
Dim markerCollection(23) As String 'Array of 24 Values
Dim referenceCellAddress As String


Dim markerCollectionReference As Variant

Dim i As Integer 'throwaway variables i to iterate vertically
Dim j As Integer                    ' j to iterate horizontally

Dim item As Variant
Dim itemReference As Integer





Sub createForms()
    Worksheets("test").Activate
    

    'Early binding declaration- Needs to be changed to late
    'binding for production

    

    
    'Loop 5 times
    'Look for the Ex + i Cell on Range A
    ' If the cell underneath is 0, next iteration
    ' If it's something else, start creating a form
    For i = 1 To 5
        Worksheets("test").Activate
        Columns(1).Find(what:="Ex" & i, MatchCase:=True).Select
        If ActiveCell.Offset(1, 0) = 0 Then
        'Do nothing
        Else
        referenceCellAddress = ActiveCell.Address
     
     
        'Create an Observation Form
     
        Set wordApp = New Word.Application
        wordApp.Visible = True
        wordApp.Activate
        wordApp.Documents.add "C:\Users\migue\Documents\Custom Office Templates\Evaluation_Form_Template.dotx"
     
        'Add Exercise Title to the Form
        Call writeTitle
        
        'Add The 4 Competencies And Its Descriptions to the form
        j = 1
        For j = 1 To 4
           Debug.Print ("Finished Writing Title")
           Debug.Print (j)
           
           Worksheets("test").Activate
           Range(referenceCellAddress).Select
           currentCompetency = ActiveCell.Offset(0, j).Value
           Call writeCompetency
           Call fetchCompetencyDescription
           Call writeCompetencyDescription
           
        Next
        'Add The Markers To The Form
        Call writeMarkers
        'Save Template as New Docx with Exercise Name
        wordApp.Activate
        wordApp.Documents(wordApp.Documents.Count).SaveAs2 Filename:=currentExercise & "_EvaluationForm"
        
        'Close Word Application
        wordApp.Quit
    
        End If
     
    Next
    
    
    Worksheets("test").Activate
    MsgBox "Evaluation Forms Have Been Created"
    
    

    
End Sub



Sub writeTitle()

    currentExercise = ActiveCell.Offset(1, 0).Value
    Debug.Print (currentExercise)
    wordApp.Selection.GoTo what:=-1, Name:="ExerciseTitle"
    wordApp.Selection.InsertAfter Text:=currentExercise
    
End Sub

Sub fetchCompetencyDescription()
    
    Worksheets("Marker Library Simulations").Activate
    Columns(1).Find(what:=currentCompetency).Select
    currentCompetencyDescription = ActiveCell.Offset(0, 1).Value
    
End Sub

Sub writeCompetency()

    wordApp.Activate
    With wordApp.Selection
        .GoTo what:=-1, Name:="CompetencyTitle" & j & "A"
        .InsertAfter Text:=currentCompetency
        .GoTo what:=-1, Name:="CompetencyTitle" & j & "B"
        .InsertAfter Text:=currentCompetency
    End With
    
End Sub

Sub writeCompetencyDescription()
    wordApp.Activate
    Debug.Print ("About to write competency")
    Debug.Print (j)
    
    With wordApp.Selection
        .GoTo what:=-1, Name:="CompetencyDesc" & j & "A"
        .InsertAfter Text:=currentCompetencyDescription
        .GoTo what:=-1, Name:="CompetencyDesc" & j & "B"
        .InsertAfter Text:=currentCompetencyDescription
    End With
    
End Sub

Sub writeMarkers()

    Sheets("test").Activate
    markerCollectionReference = Range("markerRange" & i).Value2
    itemReference = 1
    wordApp.Activate
    For Each item In markerCollectionReference
        wordApp.Selection.GoTo what:=-1, Name:="marker" & itemReference
        wordApp.Selection.InsertBefore Text:=item
        itemReference = itemReference + 1
    Next
End Sub
