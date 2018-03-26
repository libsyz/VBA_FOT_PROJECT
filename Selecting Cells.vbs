
' Shortcut Alt + F11 to go to Developer Mode Directly'

'Selecting individual cells'

'Range'
'Range is an VBA Object that can be used to refer to a single or a list of Cells'

''

Sub SelectIndividualCellsByPosition()
  Range.("A13").Select
  ActiveCell.Value = 11

  'Selecting cells by row and column'
  'cells(row,col) returns a reference to a Range object, the same methods
  'can be used on it

  cells(13, 2).Select
  ActiveCell.Value = "The Lorax"

  'You could also use shorthand'

  [B14].Select
  ActiveCell.Value = #2 Mar 2008# 'Nice way to enter dates'

End Sub


Sub selectCellsFromWorkbook ()

 'Method 1 '

 Worksheets("workshetname").activate

 'from anotherworkbook'

 Workbooks("workbookname").activate

End Sub

'Change values from cells without selection'

Sub changeValuesWithoutSelect ()
  range("A14").value = "My value"
  cells(3,3).value = "Another value"

'chang values in another worksheet without selection'

  Worksheets("workshetname").range("A14").value = 14

End Sub


'Selecting Multiple Cells'

sub selectMultipleCells()
  Range("A14:C14").Select
  'when you have more than a cell selected, use Selection'
  Selection.Interior.color = rgbDarkBlue

  range("A1:C1").Font.Color = rbgWhite

  'Shorthand also works'

  [A1:C1].font.size = 30

  'You can also do range(topleft corner, bottom right corner')

  Range("A2", "C4").Interior.color = rgbDarkBlue

end sub


'Finding the end of a list'

sub addItemToEndofList()
  Range("A1").Select

  ActiveCell.end(xlDown).select
  'once selected, you can move relatively with .offset(row,col)'
  ActiveCell.Offset(3, 0).select

  'Using last element and offset together'

  range("A1").end(xlDown).offset(1, 0).select

  ActiveCell.Value = ActiveCell(-1, 0).value + 1

end sub



sub copyingAndPasting()

  Worksheets("workshetname").activate
  'Current Region selects the whole block of data '
  range("A1").CurrentRegion.Copy

  Worksheets("workshetname2").activate
  range("a1").PasteSpecial xlPasteFormats
  range("a1").PasteSpecial xlPasteColumnWidth 'This needs to happen on a separate step'

  'You can also do the copy paste directly with an optional parameter for copy'

  range("A1").CurrentRegion.Copy Worksheets("worksheetname2").Range("A1")
  columns("A:C").Autofit
  'Also, to access intellisense'

  range("B:C").EntireColumn.Autofit
end sub


'### Questions '
'It is possible to query for a list of range names on VBA?'
' something like Worksheets('myworksheet').ranges?
