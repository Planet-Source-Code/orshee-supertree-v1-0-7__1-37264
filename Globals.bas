Attribute VB_Name = "Globals"
Option Explicit

Public Function NeverZero(Value As Long) As Long
    If Value = 0 Then
        NeverZero = 1
    Else
        NeverZero = Value
    End If
End Function

Private Function ConvertHexCharacters(sChar As String) As Integer
    Select Case sChar
        Case "0" To "9"
            ConvertHexCharacters = Val(sChar)
        Case "A" To "F"
            ConvertHexCharacters = Asc(sChar) - 55
        Case Else
            Exit Function
    End Select
End Function

Public Function GetRGBColor(ByVal LongColorValue As Long) As String
   Dim myRedValue As String
   Dim myGreenValue As String
   Dim myBlueValue As String
   Dim myHexString1 As String
   Dim myHexString2 As String
   Dim myString As String * 6
   
   myHexString2 = Hex$(LongColorValue)
   myString = "000000"
   
   RSet myString = myHexString2
   
   myRedValue = 16 * ConvertHexCharacters(Mid$(myString, 5, 1)) + ConvertHexCharacters(Mid$(myString, 6, 1))
   myGreenValue = 16 * ConvertHexCharacters(Mid$(myString, 3, 1)) + ConvertHexCharacters(Mid$(myString, 4, 1))
   myBlueValue = 16 * ConvertHexCharacters(Mid$(myString, 1, 1)) + ConvertHexCharacters(Mid$(myString, 2, 1))
   
   GetRGBColor = Str$(myRedValue) & Str$(myGreenValue) & Str$(myBlueValue)
End Function

