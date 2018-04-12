Attribute VB_Name = "AE_Create_RolePlays"
Option Explicit

'The big picture

'This macro creates role plays based on the exercises that has been
' selected.

' The Macro fetches the array of exercises selected and
' iterates through it. Depending on the iteration, creates
' a set of Documents


' We will apply different methods depending on what has been selected. We can
' work through it with a Case Statement that then calls on other Macros to do the job

Private SelectedExercises As Range
Private InstructionsDocument As Word.Application
Private CurrentPath As String
Private Exercise As Variant


    
Sub CreateExercises()

    
CurrentPath = Application.ActiveWorkbook.Path


Set SelectedExercises = Worksheets("2-Do EX-C Matrix").Range("C8", Range("C8").End(xlToRight))
Set InstructionsDocument = Nothing
    
    
    
    For Each Exercise In SelectedExercises
        
        
        
        Select Case Exercise
            
            
            Case Is = "Role Vision Presentation" 'Done
                Call CreateRoleVisionPresentation
                
            Case Is = "Case Study Presentation"
                Call CreateCaseStudyPresentation
                
            Case Is = "1st Line Manager Conversation"
                Call CreateFirstLineManagementRolePlay
                
            Case Else
                Debug.Print ("some other exercise")
        End Select
                
        
        
        
    Next Exercise
    
    

End Sub


Sub CreateRoleVisionPresentation()

            Dim FileNombre As String
            
            FileNombre = CurrentPath & "\" & Exercise & "_Participant_Instructions"
            
            'Fetch Role Vision Template, and save it in project folder
             Set InstructionsDocument = New Word.Application
             InstructionsDocument.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Participant Instruction_Role Vision Presentation.dotx"
             InstructionsDocument.Documents(InstructionsDocument.Documents.Count).SaveAs2 _
                Filename:=FileNombre, _
                FileFormat:=wdFormatDocumentDefault, _
                ReadOnlyRecommended:=False
             Set InstructionsDocument = Nothing

End Sub
           
           
Sub CreateCaseStudyPresentation()

    'Check which competencies have been selected
    'If there are targeted competencies,
    MsgBox "Feature will be implemented soon!"
    
End Sub



