Option Explicit '<- Key to avoid variable naming errors

'### Check out developer reference to understand VBA datatypes
'### How do you transform some data types into others'
'### How do you call subroutines from other modules'

Public superVariable as string '<-variable available in the whole project'
                               'declare all of them on a separate module'
Dim globalVariable as Variant '<-Variable available in the module scope'

Sub dealingWithVariables ()
  'Some variable examples'
  Dim newBall as String
  Dim newDate as Date
  Dim newLength as Integer '<- variables available in the subroutine scope

  'declaring a variable and then interpolating it'
  newBall = "football"
  msgBox newBall & " has been added to the list"


  'Autoincrease is possible'
  newLength = 4
  newLength = newLength + 1

End Sub
