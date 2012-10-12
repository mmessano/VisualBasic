 Option Explicit
'********************************************************************
'*
'* Sub Quicksort
'*
'* Purpose: Sort any array, any part of any array, in either direction,
'* as either case sensitive or insensitive. You can also choose
'* to display percent complete to the screen.
'*
'* Input: aryAny The array to be sorted
'* loBound    The index of the first (LBound) item to be sorted. Useful
'* if the array is only partially filled.
'* hiBound    The index of the first (LBound) item to be sorted. Useful
'* If the array is only partially filled.
'* blnDesc    A boolean indicating the direction to sort
'* True = Sort Descending
'* False = Sort Ascending
'* blnCase    A boolean indicating a Case Sensitive or Insensitive Sort
'* True = Case Sensitive Sort
'* False = Case Insensitive Sort
'* blnProgress    A boolean indicating display progress to the screen.
'* True = Display Progress
'* False = Do Not Display Progress
'* lngProgress Used to calculate progress, pass as an "Empty" variable
'* lngPrevious Used to calculate progress, pass as an "Empty" variable
'* lngBigO Used to calculate progress, pass as an "Empty" variable
'*
'* Output: aryAny is returned as a sorted array.
'*
'* SubRoutines Called:
'* StrFormat: Used for case insensitive sorts.
'*
'********************************************************************
Sub QuickSort(aryAny, loBound, hiBound, blnDesc, blnCase, blnProgress, lngProgress, lngPrevious, lngBigO)
Dim pivot, loSwap, hiSwap, temp, lovalue, hivalue
Dim blnFirst, orgpivot, lngEls

blnFirst = False

If blnProgress Then
'In a QuickSort BigO=(n*Log(n))
'In pratical testing BigO appears to be BigO=((n*Log(n))/3) + (n/3)
If IsEmpty(lngBigO) Then
lngEls = hiBound - loBound + 1
lngBigO = Round((lngEls * Log(lngEls))/3) + (lngEls/3)
blnFirst = True
End If 'IsEmpty(lngBigO)
If IsEmpty(lngProgress) Then
lngProgress = 0
End If '
'Wscript.Echo "lngBigO: " & lngBigO
'Wscript.Echo "lngProgress: " & lngProgress
If IsEmpty(lngPrevious) Then
lngPrevious = -1
End If '
End If 'blnProgress

'Set the variables used to provide case-insensitive searchs.
If blnCase Then
lovalue = aryAny(loBound)
hivalue = aryAny(hibound)
Else
lovalue = StrFormat(aryAny(loBound),1)
hivalue = StrFormat(aryAny(hibound),1)
End If 'blnCase

'== Two items to sort
If hiBound - loBound = 1 Then
If lovalue > hivalue Then
temp=aryAny(loBound)
aryAny(loBound) = aryAny(hiBound)
aryAny(hiBound) = temp
End If
End If

'== Three or more items to sort
pivot = aryAny(int((loBound + hiBound) / 2))
orgpivot = pivot
aryAny(int((loBound + hiBound) / 2)) = aryAny(loBound)
aryAny(loBound) = orgpivot
loSwap = loBound + 1
hiSwap = hiBound

Do
If blnProgress Then
lngProgress = lngProgress + 1
End If 'blnProgress
'Set the variables used to provide case-insensitive searchs.
If blnCase Then
lovalue = aryAny(loSwap)
hivalue = aryAny(hiSwap)
Else
pivot = StrFormat(pivot,1)
lovalue = StrFormat(aryAny(loSwap),1)
hivalue = StrFormat(aryAny(hiSwap),1)
End If 'blnCase

If blnDesc Then
'Sort Descending
'== Find the right loSwap
While (loSwap < hiSwap) and (lovalue >= pivot)
loSwap = loSwap + 1
If blnCase Then
     lovalue = aryAny(loSwap)
Else
lovalue = StrFormat(aryAny(loSwap),1)
End If 'blnCase
WEnd

'== Find the right hiSwap
While hivalue < pivot
hiSwap = hiSwap - 1
If blnCase Then
     hivalue = aryAny(hiSwap)
Else
hivalue = StrFormat(aryAny(hiSwap),1)
End If 'blnCase
WEnd
Else
'Sort Ascending
'== Find the right loSwap
While (loSwap < hiSwap) and (lovalue <= pivot)
loSwap = loSwap + 1
If blnCase Then
     lovalue = aryAny(loSwap)
Else
lovalue = StrFormat(aryAny(loSwap),1)
End If 'blnCase
WEnd

'== Find the right hiSwap
While hivalue > pivot
hiSwap = hiSwap - 1
If blnCase Then
     hivalue = aryAny(hiSwap)
Else
hivalue = StrFormat(aryAny(hiSwap),1)
End If 'blnCase
WEnd
End If 'blnDesc
'== Swap values If loSwap is less then hiSwap
If loSwap < hiSwap Then
temp = aryAny(loSwap)
aryAny(loSwap) = aryAny(hiSwap)
aryAny(hiSwap) = temp
End If
Loop While (loSwap < hiSwap)

aryAny(loBound) = aryAny(hiSwap)
aryAny(hiSwap) = orgpivot

If blnProgress Then
If Round((lngProgress/lngBigO),2) * 100 > lngPrevious Then
Wscript.Echo Cstr(Round((lngProgress/lngBigO),2) * 100) & "%"
lngPrevious = Round((lngProgress/lngBigO),2) * 100
End If 'Round((lngProgress/lngBigO),2) * 100 > intPrevious
End If 'blnProgress

'== Recursively call function .. the beauty of Quicksort
'== 2 or more items in first section
If loBound < (hiSwap - 1) then Call QuickSort(aryAny,loBound,hiSwap-1, blnDesc, blnCase, blnProgress, lngProgress, lngPrevious, lngBigO)
'== 2 or more items in second section
If hiSwap + 1 < hibound then Call QuickSort(aryAny,hiSwap+1,hiBound, blnDesc, blnCase, blnProgress, lngProgress, lngPrevious, lngBigO)
If blnProgress And blnFirst Then
Wscript.Echo "100% Complete"
End If 'blnProgress
End Sub 'QuickSort

'********************************************************************
'*
'* Function StrFormat
'*
'* Purpose: Formats a given string into the specified case.
'* intFormat = 1 = ALL UPPER CASE
'* intFormat = 2 = all lower case
'* intFormat = 3 = Every Word Starts Capitalized
'* intFormat = 4 = Only the first word is capitalized.
'*
'* Input: strAny A string that is to be formated.
'* intFormat An integer (1-4) representing the format to change
'* the string to.
'*
'* Output: Returns the given string in the given format.
'*
'********************************************************************
Public Function StrFormat(ByVal strAny, ByVal intFormat)

    Dim strTemp1     'As String
    Dim strTemp2     'As String
    Dim intCount     'As Integer
    Dim strWord     'As String
    
    Const vbFirstWord = 4
    Const vbProperCase = 3
    Const vbLowerCase = 2
    Const vbUpperCase = 1
    
    strTemp1 = strAny
    
    Select Case intFormat
     Case vbUpperCase
'UCase will convert numbers to strings.
If NOT IsNumeric(strTemp1) Then
     strTemp1 = UCase(strTemp1)
     End If 'NOT IsNumeric(strTemp1)
    
     Case vbLowerCase
     'LCase will convert numbers to strings.
     If NOT IsNumeric(strTemp1) Then
     strTemp1 = LCase(strTemp1)
     End If 'NOT IsNumeric(strTemp1)
    
     Case vbProperCase
     'LCase and UCase will convert numbers to strings.
     If NOT IsNumeric(strTemp1) Then
strTemp1 = LCase(strTemp1)
strTemp2 = Split(strTemp1)
For intCount = 0 To UBound(strTemp2)
strWord = strTemp2(intCount)
If Len(Trim(strWord)) > 0 Then
strWord = UCase(Left(strWord, 1)) & _
Right(strWord, Len(strWord) - 1)
strTemp2(intCount) = strWord
End If 'Len(Trim(strWord)) > 0
Next 'intCount
     strTemp1 = Join(strTemp2)
     End If 'NOT IsNumeric(strTemp1)
    
     Case vbFirstWord
     'LCase and UCase will convert numbers to strings.
     If NOT IsNumeric(strTemp1) Then
strTemp2 = UCase(Left(strTemp1, 1))
strTemp1 = LCase(Right(strTemp1, Len(strTemp1) -1))
strTemp1 = strTemp2 & strTemp1
End If 'IsNumeric(strTemp1)
    
     Case Else
    End Select 'intFormat
    
    StrFormat = strTemp1
End Function 'StrFormat

Randomize

Dim aryMy(1000)
Dim z, x, strTemp, myStrAry, aryAny

For z = 0 to 1000
aryMy(z) = Int((5000 - 1) * Rnd + 1)
Next
strTemp = "this,is,a,test,of,The,sorting,abilities,of,this,quick,sort,procedure,and,the,ability,to,find,duplicates,a,b,c,d,a,B,c"
myStrAry = Split(strTemp, ",")

aryAny = aryMy
'Wscript.Echo "strTemp: " & strTemp
'Sort Ascendig, Case Insensitive, and display progress
Quicksort aryAny, LBound(aryAny), UBound(aryAny), False, False, True, Empty, Empty, Empty

'Remove comments on the loop to verify the sort.
'For x = 0 To UBound(aryAny)
' Wscript.Echo "aryAny(" & x & "): " & aryAny(x)
'Next 'x

Wscript.Echo
'Sort Descendig, Case Sensitive , and display progress
Quicksort aryAny, LBound(aryAny), UBound(aryAny), True, True, True, Empty, Empty, Empty

'Remove comments on the loop to verify the sort.
'For x = 0 To UBound(aryAny)
' Wscript.Echo "aryAny(" & x & "): " & aryAny(x)
'Next 'x
