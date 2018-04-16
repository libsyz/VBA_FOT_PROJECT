Attribute VB_Name = "Repo_CreateInterview"
Option Explicit

    Public InterviewDoc As Word.Application
    
Sub CreateInterview()
    
    Dim InterviewCell() As Variant
    Dim InterviewCompetenciesRange As Range
    Dim InterviewCompetencies As Variant
    Dim CurrentCompetency As String
    Dim CurrentDescription As String
    Dim CurrentQuestionsRange As Range
    Dim CurrentQuestionsArray As Variant
    
    Set InterviewDoc = New Word.Application
    InterviewDoc.Visible = True

    InterviewDoc.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Interview_Guide_Template.dotx"
    
    '### Did not know how to store a range after using .Find, so went back to .Select
    '### To get myself moving forward - Would be good to do it more efficiently
    
    'Fetch the competency names and push them into an array
    Worksheets("2-Do EX-C Matrix").Activate
    Range("C8").Select
    
    While ActiveCell.Value <> "Interview"
        ActiveCell.Offset(0, 1).Select
    Wend
    
    Set InterviewCompetenciesRange = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 0).End(xlDown))
    
    InterviewCompetencies = InterviewCompetenciesRange
    
    Dim Competency As Variant
    Dim LoopCounter As Integer
    LoopCounter = 1
    'Iterate through the array
        For Each Competency In InterviewCompetencies
        
        CurrentDescription = Worksheets("Marker Library Interview").Columns(1).Find(Competency).Offset(0, 1).Value
        Set CurrentQuestionsRange = Range(Worksheets("Marker Library Interview").Columns(1).Find(Competency).Offset(0, 2), _
                                      Worksheets("Marker Library Interview").Columns(1).Find(Competency).Offset(0, 2).End(xlDown))
                                      
        CurrentQuestionsArray = CurrentQuestionsRange
        
        Call WriteInterviewContent(Competency, CurrentDescription, CurrentQuestionsArray, LoopCounter)
        LoopCounter = LoopCounter + 1
    Next Competency
    
    Call SaveInterview
    
    'For Each Iteration
    'Insert the Competency Name
    'Insert the Competency Description
    'Insert the Questions
    
    'Inserting should be enabled by numbered bookmarks
    'i.e. on the first iteration, we have markers called
    ' CompetencyName1
    ' CompetencyDescription1
    ' CompetencyQuestions1
    
    'We have a counter variable that starts at 1 and increases
    ' After every iteration
    
    
End Sub

Sub WriteInterviewContent(Comp, CompDesc, Questions, Counter)

        InterviewDoc.Activate
        
        With InterviewDoc
        
        .Selection.GoTo what:=-1, Name:="Competency" & Counter
        .Selection.InsertBefore Text:=Comp
        
        .Selection.GoTo what:=-1, Name:="CompetencyDescription" & Counter
        .Selection.InsertBefore Text:=CompDesc
        
        End With

        Dim Question As Variant
        
        For Each Question In Questions
        InterviewDoc.Selection.GoTo what:=-1, Name:="CompetencyQuestions" & Counter
        InterviewDoc.Selection.InsertBefore Text:=Question
        Next Question

End Sub

Sub SaveInterview()

    Dim FileNombre As String

    FileNombre = Application.ActiveWorkbook.Path & "\" & "InterviewGuide"

    InterviewDoc.Documents(InterviewDoc.Documents.Count).SaveAs2 _
        Filename:=FileNombre, _
        FileFormat:=wdFormatDocumentDefault, _
        ReadOnlyRecommended:=False
    Set InterviewDoc = Nothing
    
End Sub
