'Generate a random case sensitive alpha-numeric string
'with specified number of postitons.
'Copy key and paste it into key.txt


Option Explicit
Dim Title, NumChar, Count, strRdm, intRdm
Title = "Gen Key for Encrypton"

NumChar = InputBox("Enter a number between 1 and 1000 to generate a " & _
                   "case sensitive string with that number of characters:", _
                   Title, 8)

If NOT IsNumeric(NumChar) Then
  MsgBox Chr(34) & NumChar & Chr(34) & " is invalid input." & vbcrlf & _
         vbcrlf & "Input must be a number between 1 and 1000",, Title
  WScript.Quit
Else
  NumChar = CInt(NumChar)
  If NumChar < 1 OR NumChar > 1000 Then
    MsgBox Chr(34) & NumChar & Chr(34) & " is invalid input." & vbcrlf & _
           vbcrlf & "Input must be a number between 1 and 1000",, Title
    WScript.Quit
  End If
End If

Randomize Timer

Do Until Count = NumChar
  Count = Count + 1
  GetRdm
  strRdm = strRdm & Chr(intRdm)
Loop

InputBox NumChar & " character case sensitive string:" & vbcrlf & vbcrlf & _
         "(Press Ctrl + C to copy results to Clipboard)", Title, strRdm

Sub GetRdm
  intRdm = Int((122 - 49) * Rnd + 48)
  If intRdm > 57 And intRdm < 65 Or intRdm > 90 And intRdm < 97 Then GetRdm
End Sub