
'All worksheet methods'

'Selecting a worksheet'

Sub selectWorksheet()
  worksheets("name").Activate

  worksheets("name").Select


end

'Selecting a chart'

Sub selectChart ()
  Charts("chart1").Activate
  Charts("chart1").Select
End Sub

'Selecting a sheet of any type'

Sub moveToAnySheet ()
  Sheets("Chart1").Activate
  Sheets("Chart1").Select
End Sub


'Selecting multiple Sheets'

Sub SelectMultipleSheets ()
  worksheets("sheet1").Select
  worksheets("sheet2").Select False 'false parameter is needed to extend selection'
End Sub


'Referring to Sheets by index'

Sub referToSheetByIndex ()
  'Excel creates indexes for its worksheets starting at 1'
  'This method is unreliable - users might move your sheets around'
  worksheets(2).Select

End Sub

'referring to sheets by codenames'

sub referToSheetByCodename()

  'sheet codenames can be changed on the properties window of your sheet object'
  sheet2.Activate
  'Sheet codenames require developer specific knowledge and are harder to change,
   'use it as a best practice
End Sub



'inserting sheets'

sub addSheets()

  Worksheets.Add
  'Clicking space on this will allow to pass different arguments'
  'Before / After - where it should appear, only one can be called'
  Worksheets.Add after:=Worksheets("Sheet3")
  'Adding all the way to the end'
  Worksheets.Add before:=Sheets(1)

  'Add all the way to the beginning'
  Worksheets.Add after:=Sheets(sheets.count)

end Sub

'inserting Charts'

sub addCharts()

  Charts.Add after:=Worksheets("Sheet3")

  Sheets.Add type:=xlSheetType.xlChart

end Sub


'Deleting Sheets'

sub DeleteSheets()

  'This will prompt a confirmation dialog box'
  worksheets("Sheet5").delete

  'This disables alert prompts/dialog boxes'
  Application.displayAlerts = False
  worksheets("Sheet5").Delete

end Sub


'Delete all sheets'

Sub deleteAllSheets ()
  Sheets.delete

End Sub


'Copying & Pasting Sheets'

Sub CopySheets ()
  'As before, you can use before/after arguments to define a paste destination'
  worksheets("sheet1").copy '[Before], [after]'

  'If you want to copy into a specific workbook'
    worksheets("sheet1").copy Before:=Workbooks("Book1").Worksheets("Sheet2")

End Sub


'Moving a Sheet '

Sub movingSheets ()
  sheets("sheet1").Move after:=Sheets(Sheets.Count)
End Sub


'Rename a Sheet'

Sub renameSheet ()
  'You need to call the sheet on its codename to change its visible name'
  Sheet1.name = "New Name"
End Sub

Sub hideSheet ()

  Sheets("Sheetname").visible = xlHiddenSheet 'xlSheetVisible
                                              'xlSheetVeryHidden <- Perfect to
                                              'protect data from users'

End Sub

