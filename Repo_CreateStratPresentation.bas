Attribute VB_Name = "Repo_CreateStratPresentation"
' This Exercise does not need any customization

Option Explicit

    Private ParticipantInstructionsDoc As Word.Application
    Private CaseStudyDoc As Word.Application
    Private AnnexDoc As Word.Application
   
    Private CompetencyName As String
    Private CompetencyId As String
    Private ExerciseId As String
    Private SnippetIdArray() As String
    
    Private FinancialAcumenChecked As Boolean
    Private OperationalDecisionMakingChecked As Boolean
    Private CustomerFocusChecked As Boolean
    Private LeadingAndManagingChangeChecked As Boolean
    Private TargetedContentPresent As Boolean
    
    Private CompetenciesRange As Range
    Private Competencies() As Variant
    
    

Sub CreateCaseStudyPresentation()
    
    Call CheckCompetencies
    Call CreateParticipantInstructions
    Call CreateCaseStudy
    Call CreateAnnexes
    
End Sub
    
Sub CheckCompetencies()

    ExerciseId = "EX05"
    'Go to the sheet where the exercise selection is and assign the range of the competencies selected to Competencies()
    'Iterate through competencies
    '   - If any of the competencies is part of the targeted competencies
    '   - We add its snippetID to the SnippetArray
    
    Worksheets("2-Do EX-C Matrix").Activate
    Worksheets("2-Do EX-C Matrix").Rows(8).Find("Case Study Presentation").Select
    Set CompetenciesRange = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 0).End(xlDown))
    Competencies() = CompetenciesRange
    
    Dim Competency As Variant
    ReDim SnippetIdArray(1 To 1) As String
    For Each Competency In Competencies
    
        Select Case Competency
        
        Case Is = "Financial Acumen"
            FinancialAcumenChecked = True
            Call AddSnippetId(Competency)
        
        Case Is = "Operational Decision Making"
            OperationalDecisionMakingChecked = True
            Call AddSnippetId(Competency)
            
        Case Is = "Customer Focus"
            CustomerFocusChecked = True
            Call AddSnippetId(Competency)

        Case Is = "Leading & Managing Change"
            LeadingAndManagingChangeChecked = True
            Call AddSnippetId(Competency)
        
        End Select
    Next Competency
    
    If UBound(SnippetIdArray) > 1 Then
        TargetedContentPresent = True
    End If
        
End Sub

Sub AddSnippetId(Comp)

    Dim SnippetId As String
    
    Worksheets("Marker Library Simulations").Activate
    Columns(1).Find(Comp).Select
    CompetencyId = ActiveCell.Offset(0, 1).Value
    
    SnippetId = ExerciseId & CompetencyId
    SnippetIdArray(UBound(SnippetIdArray)) = SnippetId
    ReDim Preserve SnippetIdArray(1 To UBound(SnippetIdArray) + 1) As String
    
End Sub


Sub CreateParticipantInstructions()
    'Create a Word Document with the participant instructions
    Set ParticipantInstructionsDoc = New Word.Application
    ParticipantInstructionsDoc.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Participant Instructions_Case Study Presentation.dotx"
    ParticipantInstructionsDoc.Visible = True
    
    If TargetedContentPresent = True Then
    
        Worksheets("Case Study Presentation Library").Activate
        Dim SnippetTargetedIntro As String
        SnippetTargetedIntro = Worksheets("Case Study Presentation Library").Range("C11").Value
        
        Call WriteSnippet(ParticipantInstructionsDoc, SnippetTargetedIntro, "TargetedIntroBookmark")
        
        Dim SnippetId As Variant
        
        For Each SnippetId In SnippetIdArray
        
            Dim Snippet As String
            Snippet = Worksheets("Case Study Presentation Library").Columns(2).Find(SnippetId).Offset(0, 1).Value
            Call WriteSnippet(ParticipantInstructionsDoc, Snippet, "TargetedGoalBookmark")
            
        Next SnippetId
 
    End If
    'Add Call to Save Documents
    Call SaveInstructions(ParticipantInstructionsDoc, "Participant_Instructions")
    
End Sub

Sub CreateCaseStudy()
    
        Set AnnexDoc = New Word.Application
        AnnexDoc.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Maskabbah_Case_Study.dotx"
        AnnexDoc.Visible = True
    
    Call SaveInstructions(AnnexDoc, "Maskabbah_Case_Study")
    
End Sub

Sub CreateAnnexes()
    
    If OperationalDecisionMakingChecked = True Or FinancialAcumenChecked = True Then
        Set AnnexDoc = New Word.Application
        AnnexDoc.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Case Study Presentation Annex.dotx"
        AnnexDoc.Visible = True
    End If

    Call SaveInstructions(AnnexDoc, "Maskabbah_Case_Study_Annex")
    
End Sub

Sub WriteSnippet(WordDoc, Snippet, BookMark)

    WordDoc.Activate
    With WordDoc.Selection
        .GoTo what:=-1, Name:=BookMark
        .InsertParagraphAfter
        .InsertAfter Text:=Snippet
    End With
    
End Sub

Sub SaveInstructions(WordDoc, ExerciseName)

Dim FileNombre As String

FileNombre = Application.ActiveWorkbook.Path & "\" & "Case Study Presentation"

    WordDoc.Documents(WordDoc.Documents.Count).SaveAs2 _
        Filename:=FileNombre & "_" & ExerciseName, _
        FileFormat:=wdFormatDocumentDefault, _
        ReadOnlyRecommended:=False
        
    WordDoc.Documents(WordDoc.Documents.Count).Close
    WordDoc.Quit
    Set WordDoc = Nothing
    
End Sub







