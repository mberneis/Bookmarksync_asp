<%

' reptext(line, prefix)
'
' split a single line on word boundaries.  Lines are as long
' as they can get without exceeding 70 characters.  Lines are
' prefixed with the string specified by the pretxt argument,
' and are suffixed with CR LF.
'
function reptext (iline,pretxt)
	Dim ltxt, i

	'Split single line on word boundaries
	const cutoff = 70
	ltxt = ""
	while len(iline) > cutoff
		i = cutoff
		while i > 0 and mid(iline,i,1) <> " "
			i = i - 1
		wend
		if i = 0 then i = 70 'If no space at all then just cut off
		ltxt = ltxt & pretxt & left(iline,i) & vbCrLf
		iline = right(iline,len(iline)-i)
	wend
	if iline <> vbCrLf then
		reptext = ltxt & pretxt  & iline &  vbCrLf
	else
		reptext = ltxt
	end if
end function

' makereply(input, prefix)
'
' creates a word-wrapped, prefixed set of lines of text
' based on an input string.  The idea is to create a
' typical e-mail reply.  For instance:
'    makereply("this is a test message.", "> ")
' returns
'    "> this is a test message"
'
' lines are wrapped at 70 characters
'
function makereply (inptxt,pretxt)
	dim outtxt, i

	outtxt = ""
	do
		i = instr(inptxt,vbCrLf)
		if i = 0 then exit do
		outtxt = outtxt & reptext(left(inptxt,i-1),pretxt)
		inptxt = right(inptxt,len(inptxt) - i -1)
	loop
	makereply = outtxt & reptext(inptxt,pretxt)
end function

sub domail (xfrom,xto,xsubj,xbody)
	If RunningLocal() Then
		response.write "<script language=javascript>" & vbCrLf
		response.write "alert(""From: " & xfrom & "\nTo: " & xto & "\nSubject: " & xsubj & "\nBody: " & replace(xbody,vbCrLf,"\n") & """);" & vbCrLf
		response.write "</script>"
	Else
'		dim ObjMail
'
'		Set ObjMail = Server.CreateObject("CDONTS.NewMail")
'		objMail.From = xfrom
'		objMail.Subject = xsubj
'		objMail.Body = xbody & vbCrLf & "-------------------------------------------------------------------------" & vbCrLf & "Keep all your Browser Bookmarks synchronized - no matter where you are..." & vbCrLf & "http://www.bookmarksync.com" & vbCrLf
'		objMail.To = xto
'		objmail.Send
'		set objmail = nothing
	end if
end sub

%>