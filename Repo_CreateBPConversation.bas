Attribute VB_Name = "Repo_CreateBPConversation"
' This Exercise does not need any customization

Option Explicit

    Private ParticipantInstructionsDoc As Word.Application
    Private RolePlayerInstructionsDoc As Word.Application

Sub CreateBusinessPartnerRolePlayWord()

    Call CreateInstructions
    Call SaveInstructions
    
End Sub
    

Sub CreateInstructions()
 
    Set ParticipantInstructionsDoc = New Word.Application
    Set RolePlayerInstructionsDoc = New Word.Application
    ParticipantInstructionsDoc.Visible = True
    RolePlayerInstructionsDoc.Visible = True
  
    ParticipantInstructionsDoc.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Participant Instruction_Business Partner Conversation.dotx"
    RolePlayerInstructionsDoc.Documents.Add "C:\Users\migue\Documents\Custom Office Templates\Role Player Instructions_Business Partner Conversation.dotx"

End Sub


Sub SaveInstructions()

Dim FileNombre As String

FileNombre = Application.ActiveWorkbook.Path & "\" & "Business Partner Conversation"

    ParticipantInstructionsDoc.Documents(ParticipantInstructionsDoc.Documents.Count).SaveAs2 _
        Filename:=FileNombre & "_Participant_Instruction", _
        FileFormat:=wdFormatDocumentDefault, _
        ReadOnlyRecommended:=False
    Set ParticipantInstructionsDoc = Nothing

    RolePlayerInstructionsDoc.Documents(RolePlayerInstructionsDoc.Documents.Count).SaveAs2 _
        Filename:=FileNombre & "_Role_Player_Instruction", _
        FileFormat:=wdFormatDocumentDefault, _
        ReadOnlyRecommended:=False
    Set RolePlayerInstructionsDoc = Nothing


End Sub




