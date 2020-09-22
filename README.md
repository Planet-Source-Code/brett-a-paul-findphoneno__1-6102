<div align="center">

## FindPhoneNo


</div>

### Description

This function searches a text string for possible phone numbers and returns an array of those numbers. It allows you to specify a default area code, too. (If you know an easier or more elegant way to do this, let me know!) Doesn't work for international numbers.
 
### More Info
 
Text - Text to be searched for phone numbers

DefaultAreaCode - if a 7-digit phone number is found, this goes on the front of it

Variant: Array of phone numbers. The array is dimensioned at 0 to start, so if the return's UBound is 0, no phone numbers were found.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brett A\. Paul](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brett-a-paul.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brett-a-paul-findphoneno__1-6102/archive/master.zip)





### Source Code

```
Function FindPhoneNo(ByVal strAdText As String, _
    strDefaultAreaCode As String) As Variant
' By Brett A. Paul - http://www.mitagroup.com/
' This routine takes the incoming ad text and abstracts it out
' (strAbstract) to perform some basic pattern matching. It also
' builds a parallel real string (strReal) so that it knows where the
' patterns came from and what they really are. Using this technique,
' the routine builds patterns, then examines them for phone number
' patterns.
Dim aPossible() As String ' This will hold the result set
Dim strReal As String ' This will hold the pattern-modified real numbers
Dim strAbstract As String ' This holds the pattern of the string
Dim strChar As String * 1 ' Holds 1 letter at a time from input string
Dim ptrWhere As Long ' Used in InStr functions
Dim ptrChar As Integer
Dim ptrPossible As Integer ' Points to last used possible array loc
ReDim aPossible(0) ' Will return array with element 0 if no #s found
' Remove dollar amounts from string
Do
  ptrWhere = InStr(strAdText, "$")
  If ptrWhere Then
    ' If a "$" is found, remove all numbers that appear after the
    ' "$". Note: This would need to be changed to allow for
    ' decimal places.
    Do While IsNumeric(Mid$(strAdText, ptrWhere + 1, 1))
      strAdText = Left$(strAdText, ptrWhere) & Right$(strAdText, _
          Len(strAdText) - (ptrWhere + 1))
    Loop
    ' Once the numbers are gone, take off the "$", too
    strAdText = Left$(strAdText, ptrWhere - 1) & Right$(strAdText, _
        Len(strAdText) - ptrWhere)
  End If
Loop Until ptrWhere = 0
' Begin building abstract and real strings for pattern matching
strReal = ""
strAbstract = ""
For ptrChar = 1 To Len(strAdText)
  ' Pick up the next character in the input string
  strChar = Mid$(strAdText, ptrChar, 1)
  If InStr(",-() :;!#%&*/", strChar) Then
    ' If character is one of these symbols, add a "-"
    ' This allows for phone numbers like (800) 555-1212
    ' or 800/555-1212, or however else people like to write
    ' phone numbers
    If Right$(strAbstract, 1) <> "-" And _
        Right$(strAbstract, 1) <> ">" Then
      strAbstract = strAbstract & "-"
      strReal = strReal & "-"
    End If
  ElseIf IsNumeric(strChar) Then
    ' If character is numeric, add a "#"
    strAbstract = strAbstract & "#"
    strReal = strReal & strChar
  Else
    ' If the character is something else, add "-" for the first
    ' character, or <-> for more than one character.
    Select Case Right$(strAbstract, 1)
      Case ",", "#", ""
        strAbstract = strAbstract & "<->"
        strReal = strReal & "<->"
      Case ">" ' Nothing to do - already has delimiter
      Case "-"
        strAbstract = Left$(strAbstract, _
            Len(strAbstract) - 1) & "<->"
        strReal = Left$(strReal, Len(strReal) - 1) & "<->"
    End Select
  End If
Next ptrChar
' When two phone numbers appear right next to each other, they may
' blend together in the pattern. To isolate each phone number,
' separate the two with a delimiter <->. This is done by looking for
' places where a dash and four numbers in a row are followed by
' another dash in the abstract pattern
Do
  ptrWhere = InStr(strAbstract, "-####-")
  If ptrWhere Then
    strAbstract = Left$(strAbstract, ptrWhere + 4) & "<->" & _
        Right$(strAbstract, Len(strAbstract) - (ptrWhere + 5))
    strReal = Left$(strReal, ptrWhere + 4) & "<->" & _
        Right$(strReal, Len(strReal) - (ptrWhere + 5))
  End If
Loop Until ptrWhere = 0
' Now that the patterns are ready, search for phone number patterns.
ptrPossible = 0
Do
  ' Begin by searching for ###-####
  ptrWhere = InStr(strAbstract, "###-####")
  If ptrWhere Then ' Found a phone number
    If Mid$(strAbstract, ptrWhere + 8, 1) = "#" Then
      ' Too many numbers; this is not really a phone number.
      ' Remove the substring
      strAbstract = Left$(strAbstract, ptrWhere - 1) & _
          Right$(strAbstract, Len(strAbstract) - _
              (ptrWhere + 7))
      strReal = Left$(strReal, ptrWhere - 1) & _
          Right$(strReal, Len(strReal) - (ptrWhere + 7))
    Else
      If ptrWhere > 4 Then ' Check for inclusion of area code
        If Mid$(strAbstract, ptrWhere - 4, 4) = "###-" Then
          ' Area code included
          ' Add phone number to list of possibles
          ptrPossible = ptrPossible + 1
          ReDim Preserve aPossible(ptrPossible)
          aPossible(ptrPossible) = Mid$(strReal, ptrWhere - 4, 12)
          ' Extract the substring from the abstract and
          ' real string so they don't get in the way of the
          ' next search
          strAbstract = Left$(strAbstract, ptrWhere - 5) & _
              Right$(strAbstract, Len(strAbstract) - _
                  (ptrWhere + 7))
          strReal = Left$(strReal, ptrWhere - 5) & _
              Right$(strReal, Len(strReal) - _
                  (ptrWhere + 7))
        Else
          ' Area code not included - use default
          ' Add phone number to list of possibles
          ptrPossible = ptrPossible + 1
          ReDim Preserve aPossible(ptrPossible)
          aPossible(ptrPossible) = strDefaultAreaCode & _
              "-" & Mid$(strReal, ptrWhere, 8)
          ' Extract the substring from the abstract
          ' and real string so they don't get in the way of
          ' the next search
          strAbstract = Left$(strAbstract, ptrWhere - 1) & _
              Right$(strAbstract, Len(strAbstract) _
                  - (ptrWhere + 7))
          strReal = Left$(strReal, ptrWhere - 1) & _
              Right$(strReal, Len(strReal) - _
              (ptrWhere + 7))
        End If
      Else
        ' Too close to the front of the string - can't
        ' have area code
        ' Use default area code
        ' Add phone number to list of possibles
        ptrPossible = ptrPossible + 1
        ReDim Preserve aPossible(ptrPossible)
        aPossible(ptrPossible) = strDefaultAreaCode & "-" & _
            Mid$(strReal, ptrWhere, 8)
        ' Extract the substring from the abstract
        ' and real string so they don't get in the way
        ' of the next search
        strAbstract = Left$(strAbstract, ptrWhere - 1) & _
            Right$(strAbstract, Len(strAbstract) - _
                (ptrWhere + 7))
        strReal = Left$(strReal, ptrWhere - 1) & _
            Right$(strReal, Len(strReal) - (ptrWhere + 7))
      End If
    End If
  End If
Loop Until ptrWhere = 0
' Finished! Set function result to the array of possible phone numbers
FindPhoneNo = aPossible
Exit_FindPhoneNo:
  Exit Function
End Function
Function TestIt()
Dim aPhoneNumbers() As String
Dim ptrNumber As Long
aPhoneNumbers = FindPhoneNo("blah blah blah (800) - 555 - 1212 blah 555 1212 blah 350319 340193 blah blah 800/349/49/40 bl 800/349/0044 ah ", "800")
For ptrNumber = 1 To UBound(aPhoneNumbers)
  Debug.Print aPhoneNumbers(ptrNumber)
Next ptrNumber
End Function
```

