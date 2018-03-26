
'Unlike standard variables, objct variables store
'references to entire datasets

'Storing a range of cells into a variable'

option Explicit

sub StoreRangeOfCells()

  dim FilmNameCells as Range

  'Object Variables always need to be SET'
  'FIlmNameCells will always reference the original range, in its original
  'workbook

  set FilmNameCells = Range("A2:A15")

  FilmNameCells.Font.Color = rbgWhite


End Sub

'You can also reference workbooks'

Sub WorkbooksReference ()

   Dim myWorkbook as Workbook

End Sub


'Finding a Range'


Sub FindingARange ()
  Dim FilmToFind as String
  Dim FilmCell as Range

  FilmToFind = InputBox("Type in A Film Name")

  set FilmCell =
      Range("B3", Range("B3").end(xlDown)).find(FilmToFind)

      msgBox FilmCell.Value & " was found in cell " & FilmCell.Address

'### Check out find method, seems efficient

End Sub
