'==========================================================================
'
' NAME: sort.vbs
'
' AUTHOR: Mark D. MacLachlan , The Spider's Parlor
' URL: http://www.thespidersparlor.com
' DATE  : 2/10/2004
'
' COMMENT: Reads a file into an array, sorts it and writes the data back to the file.
'
'==========================================================================


Option Explicit
Dim oFSO, ForReading, ForWriting, sortFile, MyList, myArray, ts, i, j, temp, line, report
ForReading = 1
ForWriting = 2
Set oFSO=CreateObject("Scripting.FileSystemObject")
'comment out the next line if you want to supress prompting for the file location
sortFile = InputBox("What file should I sort?  Full path please!", "File To Sort")
'uncomment the next line if you want to have a static file location
'sortFile = "C:\Test.log"
MyList= ofso.OpenTextFile(sortFile, ForReading).ReadAll
myArray=Split(MyList,vbCrLf, -1, vbtextcompare)
    
'bubble sort thanks to Richard Lowe, 4GuysFromRolla.com
'what he does here is check each element in the array
'against the next value to see if it is greater than it.
'If location1 is > location2 write location1 to temp,
'then write location2 to location1 and finally write
'temp to location2

for i = UBound(myArray) - 1 To 0 Step -1
    for j= 0 to i
        if myArray(j)>myArray(j+1) then
            temp=myArray(j+1)
            myArray(j+1)=myArray(j)
            myArray(j)=temp
        end if
    next
next
'end bubble sort.  Thanks Richard!

For Each line In myArray
    'Check for blank lines and ignore them
    If Len(line) <> 0 Then
    report = report & line & vbcrlf
    End If
Next
   
MsgBox "The following will be written to " & sortfile & vbCrLf & report, vbOkOnly
  
'Now write the data back to the original file in sorted order
Set ts = oFSO.CreateTextFile (sortFile, ForWriting)
ts.write report

