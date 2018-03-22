Attribute VB_Name = "Init"
'Private i As Integer
Private competencies() As String

Sub init()
    ReDim competencies(1 To 1) As String
    Dim my_name As String
    Dim business_drivers() As String
    ReDim business_drivers(1 To 1) As String
    Dim businessDriver As String
    Dim i As Integer
    
    Worksheets("1-Select Business Drivers").Activate
    Range("A7").Select
    
    name = ActiveCell.Offset(0, 1).Value
    
    '### hard coded, find last value i column A
    For i = 0 To 20
        If ActiveCell.Value = "x" Or ActiveCell.Value = "X" Then
            'Debug.Print ActiveCell.Offset(0, 1).Value
            businessDriver = ActiveCell.Offset(0, 1)
            business_drivers(UBound(business_drivers)) = businessDriver '### do we need it?
            ReDim Preserve business_drivers(1 To UBound(business_drivers) + 1) As String
            Call store_competency(businessDriver)
        End If
        Worksheets("1-Select Business Drivers").Activate
        ActiveCell.Offset(1, 0).Select
        'Debug.Print competencies(1)
    Next i
    
    'see competencies on excel ### get rid of it
    'Range("E2").Select
    'For i = 1 To UBound(competencies)
        'Range(ActiveCell, ActiveCell) = competencies(i)
        'ActiveCell.Offset(1, 0).Select
    'Next i
    
    'create dropdowns
    Call create_drop_downs
    
    
End Sub

Sub store_competency(businessDriver As String)
    Dim competency As String
    
    'move to appropiate sheet
    Worksheets("Z1 - Lib Business Drivers").Activate
    'find businessDriver
    Range("A2").Select
    
    For i = 1 To 60
        If ActiveCell.Offset(0, 0).Value = businessDriver Then
            
            competency = ActiveCell.Offset(0, 1)
            
            While competency <> ""
                If is_in_array(competency, competencies) Then
                    
                Else
                    competencies(UBound(competencies)) = ActiveCell.Offset(0, 1)
                    ReDim Preserve competencies(1 To UBound(competencies) + 1) As String
                    
                End If
                ActiveCell.Offset(1, 0).Select
                competency = ActiveCell.Offset(0, 1)
                
            Wend
            
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
End Sub

Function is_in_array(string_to_be_found As String, arr As Variant) As Boolean
    is_in_array = (UBound(Filter(arr, string_to_be_found)) > -1)
End Function

Sub create_drop_downs()
    Dim i As Integer
    Dim j As Integer
    
    Worksheets("2-Do EX-C Matrix").Activate
    Range("C9").Select
    For j = 0 To 4
        Range("C9").Offset(0, j).Select
        For i = 1 To 4
            With ActiveCell.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, Formula1:=Join(competencies, ",")
            End With
            ActiveCell.Offset(1, 0).Select
        Next
    Next
End Sub








