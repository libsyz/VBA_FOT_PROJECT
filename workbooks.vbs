'Ay dios, ay mi cabeza\

Sub referringToWorkBooksByName ()

  workbooks("book2.xlsx").activate

End Sub

Sub referringToWorkBooksByIndexNumber ()

  workbooks(2).activate

End Sub


Sub UsingActiveWorkBook()
  Workbooks("Book2.xlsx").Activate
  Workbooks("Book2.xlsx").Close 'another workbook will become active
                                'after this line runs

  Workbooks("A_workbook.xlsx").Close True 'closes and saves changes'

  ThisWorkbook.Close '<- closes the workbook in which the sub is stored'
End



Sub openExistingWorkbook()

  Workbooks.Open "Enter Absolute Path"
end sub


Sub createNewWorkbook()

  Workbooks.Add "A new workbook"

End Sub


Sub savingWorkbookThatHasAlreadyBeenSaved ()

  Workbooks("workbook").save

End Sub



Sub savingNewWorkbook ()

  Workbooks.Add
  Workbooks.Save

End Sub

Sub savingOnSpecificDestination ()

  Workbooks.Add
  Workbooks.SaveAs 'Enter absolute path as parameter
                   'Parameters on save as also allow you to save
                   'different filetypes

End Sub

