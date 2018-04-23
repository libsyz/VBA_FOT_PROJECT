Attribute VB_Name = "Repo_CreateCustomerConversation"
Option Explicit
    Private CompetencyName As String
    Private CompetencyId As String
    Private ParticipantInstructionsDoc As Word.Application
    Private RolePlayerInstructionsDoc As Word.Application
    Dim ExerciseId As String
    Dim SnippetIdArray() As String

Sub CreateFirstLineManagementRolePlayWord()
    
    Dim IssueMatrix() As Variant
    Dim IssueRange As Range
    
    
    
    Dim isX As Boolean
    Dim i As Integer
    Dim Counter As Integer
    
    
    Dim CompetencyNamesArray() As String
    Dim CompetencyIdsArray() As String
    
    Dim SingleCell As Range
    ReDim SnippetIdArray(1 To 1) As String
    
    Worksheets("Role Play Selector").Activate
    ExerciseId = "EX03"
    Set IssueRange = Range("A35:C39")
    IssueMatrix = IssueRange
    
    Counter = 0
    
    For i = LBound(IssueMatrix, 1) To UBound(IssueMatrix, 1)
        If IssueMatrix(i, 1) = "X" Or IssueMatrix(i, 1) = "x" Then
            Counter = Counter + 1
            ReDim CompetencyNamesArray(1 To Counter)
            CompetencyNamesArray(Counter) = IssueMatrix(i, 3)
        End If
        
    
    Next i
        
    For Each SingleCell In Range("A35:C39")
        
        SingleCell.Select
        'Call CheckValue
        isX = IssueIsChecked()
        If isX = True Then
            Call FetchCompetencyId(CompetencyName)
            Call AddSnippetId
        End If
        
    Next SingleCell
    
    Call CreateInstructions
    Call SaveInstructions
    
End Sub
    

Function IssueIsChecked() As Boolean

    Dim isX As Boolean
    isX = False
    
    If ActiveCell.Value = "X" Or ActiveCell.Value = "x" Then
        isX = True
        CompetencyName = ActiveCell.Offset(0, 2).Value
    End If
    
    IssueIsChecked = isX
   
End Function

Sub FetchCompetencyId(CompetencyName)
    
    Worksheets("Marker Library Simulations").Activate
        If CompetencyName <> "Selling the Vision / Leading Change / Leading & Managing Change" Then
            Columns(1).Find(CompetencyName).Select
            CompetencyId = ActiveCell.Offset(0, 1).Value
        Else
            CompetencyId = Range("B94").Value
    End If

End Sub

Sub AddSnippetId()
    Dim SnippetId As String
      If ActiveCell.Value = "" Then
    
    Else
        SnippetId = ExerciseId & CompetencyId
        
        SnippetIdArray(UBound(SnippetIdArray)) = SnippetId
        ReDim Preserve SnippetIdArray(1 To UBound(SnippetIdArray) + 1) As String
    End If
    Worksheets("Role Play Selector").Activate
End Sub


Sub CreateInstructions()
 
    Set ParticipantInstructionsDoc = New Word.Application
    Set RolePlayerInstructionsDoc = New Word.Application
    ParticipantInstructionsDoc.Visible = True
    RolePlayerInstructionsDoc.Visible = True
  
    ParticipantInstructionsDoc.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Participant Instructions_Customer Conversation.dotx"
    RolePlayerInstructionsDoc.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Role Player Instructions_Customer Conversation.dotx"
 
    Dim SnippetId As Variant
    Dim SnippetAbout As String
    Dim SnippetGoal As String
    Dim SnippetBehavior As String
    Dim SnippetPerspective As String
    Dim SnippetSuccessBehavior As String
    Dim SnippetDerailmentBehavior As String
    
 
    For Each SnippetId In SnippetIdArray
    
        Dim ReferenceCell As Range
        Set ReferenceCell = Worksheets("Customer Conversation Library").Columns(2).Find(SnippetId)
        
        SnippetAbout = ReferenceCell.Offset(0, 2).Value
        Call WriteSnippet(ParticipantInstructionsDoc, SnippetAbout, "AboutBookmark")

        If ReferenceCell.Offset(0, 3).Value <> "" Then
            SnippetGoal = ReferenceCell.Offset(0, 3).Value
            Call WriteSnippet(ParticipantInstructionsDoc, SnippetGoal, "GoalBookmark")
        Else
            Debug.Print ("No goals to add")
        End If
        
        If ReferenceCell.Offset(0, 4).Value <> "" Then
            SnippetBehavior = ReferenceCell.Offset(0, 4).Value
            Call WriteSnippet(RolePlayerInstructionsDoc, SnippetBehavior, "BehaviorBookmark")
        End If
        
        If ReferenceCell.Offset(0, 5).Value <> "" Then
            SnippetPerspective = ReferenceCell.Offset(0, 5).Value
            Call WriteSnippet(RolePlayerInstructionsDoc, SnippetPerspective, "PerspectiveBookmark")
        End If
        
        If ReferenceCell.Offset(0, 6).Value <> "" Then
            SnippetSuccessBehavior = ReferenceCell.Offset(0, 6).Value
            Call WriteSnippet(RolePlayerInstructionsDoc, SnippetSuccessBehavior, "SuccessBehaviorBookmark")
        End If
        
        If ReferenceCell.Offset(0, 6).Value <> "" Then
            SnippetDerailmentBehavior = ReferenceCell.Offset(0, 7).Value
            Call WriteSnippet(RolePlayerInstructionsDoc, SnippetDerailmentBehavior, "DerailmentBehaviorBookmark")
        End If
        
    Next SnippetId
    
    
    
  '### All this can be refactored with another loop and using a hash/dictionary for snippets
  '### How to Delete Last Array Element?
  '### Think about how to delete First Or Last Line

End Sub


Sub WriteSnippet(WordDoc, Snippet, BookMark)

    WordDoc.Activate
    With WordDoc.Selection
        .GoTo what:=-1, Name:=BookMark
        .InsertParagraphAfter
        .InsertAfter Text:=Snippet
    End With
    
End Sub


Sub SaveInstructions()

Dim FileNombre As String

FileNombre = Application.ActiveWorkbook.Path & "\" & "Customer Conversation"

    ParticipantInstructionsDoc.Documents(ParticipantInstructionsDoc.Documents.Count).SaveAs2 _
        Filename:=FileNombre & "_Participant_Instruction", _
        FileFormat:=wdFormatDocumentDefault, _
        ReadOnlyRecommended:=False
    Set ParticipantInstructionsDoc = Nothing
    'Filename = 1stLine Management Conversation_Participant Instruction.Docx
    'Filename = 1stLine Management Conversation_RolePLayer Instruction.Docx
    RolePlayerInstructionsDoc.Documents(RolePlayerInstructionsDoc.Documents.Count).SaveAs2 _
        Filename:=FileNombre & "_Role_Player_Instruction", _
        FileFormat:=wdFormatDocumentDefault, _
        ReadOnlyRecommended:=False
    Set RolePlayerInstructionsDoc = Nothing


End Sub





