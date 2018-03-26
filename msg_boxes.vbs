Option Explicit
'diplaying a simple message'

Sub displayingSimpleMessage ()

  msgBox "This is a message"

End Sub

'diplaying a message that concatenates a variable'

Sub displayConcatMessage ()

  Dim name as String
  name = "Miguel"
  msgBox "my name is " & name

  'add a new line'
  msgBox "my name is " & name & vbNewLine & "Nice to mee you"

End Sub

'customizing a message'

Sub customMessage ()

'Several parameters will come up on intellisense
'example

  msgBox prompt:="I like Pizza", title:="this is a box", Buttons:="vbInformation"

End Sub

'Getting input from users'

Sub AskQuestion ()

  dim Buttonclicked as VbMsgBoxResult

  'vbQuestion offers lots of different parameters to play with'
  'when you store the result of messagebox into a variable, parameters need to be
  'enclosed in brackets

  ButtonClicked = msgBox("Do you like pizza", vbQuestion + vbYesNo, "Food Question")

  'declaring VbMsgBox will allow Intellisense to display which values are accepted'


End Sub


