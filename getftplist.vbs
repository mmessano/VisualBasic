' Chris Crowe
' IISFAQ Web Site
' http://www.iisfaq.com
' September 24, 2000
'
' Show ALL FTP Sites

'Set IISOBJ = getObject("IIS://alya/MSFTPSVC")
'For each Object in IISOBJ
'	if (Object.Class = "IIsFtpServer") then
'		WScript.Echo "FTP Site: " & Object.Name & " - " & Object.ServerComment
'	end if
'next



on error resume next

' get a list of servers from david's/doug's server sheet on serrano
' (now on xfs3, not serrano - jce 7/13/05)
' it's on the second sheet of the workbook (svrsheet=2)
'
' changed the spreadsheet location to xfs as serrano is shutting down

dim svrsheet, xlo
dim r, c
dim ssht
dim objiis, objsite, object, server
dim fso, outfile, iislist, line
dim today, yesterday, yesterfile, remdir

remdir = "\\messano338\dexma\logs\"

today = date
yesterday = dateadd("d",-1,today)

outfile = "ftpsites-" & month(date) & "-" & day(date) & ".txt"
yesterfile = "ftpsites-" & month(yesterday) & "-" & day(yesterday) & ".txt"
set fso = createobject("scripting.filesystemobject")
set iislist = fso.createtextfile(outfile, true)

ssht = "C:\Dexma\test_servers2.xls"
svrsheet = 1

set xlo = wscript.createobject("excel.application")
xlo.workbooks.open ssht
xlo.sheets(svrsheet).activate

' the list is in the C column (column 3) and
' starts on the eighth row

c = 1
r = 1

' loop through the list and use that entry as a server name
' to see if there are any web sites on it
' until you hit a blank cell

do until xlo.cells(r, c).value = ""
	server = xlo.cells(r,c).value
	object = "IIS://" & server & "/MSFTPSVC"
	set objiis = getobject(object)

	wscript.echo server & " FTP sites..."

	if (err <> 0) then
		wscript.echo "error: " & hex(err.number) & "(" & err.description & ")"
	else
		for each objsite in objiis
			if objsite.class = "IIsFtpServer" then
				line = server & "," & objsite.name & "," & objsite.servercomment
				iislist.writeline(line)
			end if
		next
	end if
	err.clear
	r = r + 1
loop

' copy the file to a public server for the webpage

fso.copyfile outfile, remdir

if err <> 0 then
	wscript.echo "error:" & hex(err.number) & "(" & err.description & ")"
end if

' if yesterday's file isn't from monday, delete it

if weekday(yesterday) <> 2 then
	fso.deletefile yesterfile
	fso.deletefile remdir & yesterfile
end if

xlo.quit
set xlo = nothing
set objiis = nothing
