Attribute VB_Name = "VBALib_MathFunctions"
' Common VBA Library - MathFunctions
' Provides useful mathematical functions.

Option Explicit

' Returns the lesser of its two arguments.
Public Function Min(a As Double, b As Double) As Double
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

' Returns the greater of its two arguments.
Public Function Max(a As Double, b As Double) As Double
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

' Returns its argument truncated (rounded down) to the given significance or
' the given number of decimal places.
' @param significance: The significance, or step size, of the function.  For
' example, a step size of 0.2 will ensure that the number returned is a
' multiple of 0.2.
' @param places: The number of decimal places to keep.
Public Function Floor(num As Double, _
    Optional significance As Double = 1, _
    Optional places As Integer = 0) As Double
    
    ValidateFloorCeilingParams significance, places
    Floor = Int(num / significance) * significance
End Function

' Returns its argument rounded up to the given significance or the given number
' of decimal places.
' @param significance: The significance, or step size, of the function.  For
' example, a step size of 0.2 will ensure that the number returned is a
' multiple of 0.2.
' @param places: The number of decimal places to keep.
Public Function Ceiling(num As Double, _
    Optional significance As Double = 1, _
    Optional places As Integer = 0) As Double
    
    ValidateFloorCeilingParams significance, places
    Ceiling = Floor(num, significance)
    If num <> Ceiling Then Ceiling = Ceiling + significance
End Function

Private Sub ValidateFloorCeilingParams( _
    ByRef significance As Double, _
    ByRef places As Integer)
    
    If places <> 0 Then
        If significance <> 1 Then
            Err.Raise 32000, Description:= _
                "Pass either a number of decimal places or a significance " _
                & "to Floor() or Ceiling(), not both."
        Else
            significance = 10 ^ -places
        End If
    End If
End Sub
