<%
option explicit

Dim mainmenu, submenu, SUBMEN, doUpload 

Sub OnTransactionAbort()
	response.write "Our website experiences temporary problems.<br>The site administrator has been notified.<hr>"	
	response.write "Please try again shortly..."
	on error resume next
	conn.close
	set conn=Nothing
	response.end	
end sub


Function BrowserType()
	dim userAgent

	userAgent = request.servervariables("HTTP_USER_AGENT")

	If (Instr(userAgent, "Opera") > 0) Then
		BrowserType = "op"
	ElseIf (Instr(userAgent, "MSIE") > 0) Then
		BrowserType = "ie"
	Else
		BrowserType = "ns"
	End If
End Function

'********************************************
'*** Some useful functions...
'********************************************

function htmltxt (h)
	Dim hx,i,j
	hx=h
	hx = replace(hx,"<p>",vbCrLf & vbCrLf,1,-1,1)
	hx = replace(hx,"<br>",vbCrLf,1,-1,1)
	hx = replace(hx,"<hr>",vbCrLf & "--------------------------------------------------" & vbCrLf,1,-1,1)
	do
		i = instr(hx,"<")
		if i=0 then exit do
		j = instr(i,hx,">")
		if j=0 then exit do
		hx = left(hx,i-1) & right(hx,len(hx)-j)
	loop
	htmltxt=hx
end function

function y2k(MyDate)
	if Year (MyDate) < 2000 then
		y2k = Replace (MyDate & "",Right(Year(MyDate),2),Year(MyDate))
	else
		y2k = MyDate
	end if
end function

Function strip (txt)
	strip = replace (txt,"'","''")
End Function

Function strip2 (txt)
	strip = replace(replace(replace (txt,"'","''"),"%","[%]"),"_","[_]")
End Function

Function prep (s)
	prep = replace(replace(s,"%","%%"),"\","%_")
	'prep = replace(replace(s,"\","%_"),"%","%%")
end function

Function Unprep(s)
	unprep = replace(replace(s,"%_","\"),"%%","%")
	'unprep = replace(replace(s,"%%","%"),"%_","\")
End Function

Function preps (s)
	prep = replace(replace(replace(s,"'","''"),"%","%%"),"\","%_")
end function

Function Unpreps(s)
	unprep = replace(replace(replace(s,"'","''"),"%_","\"),"%%","%")
End Function

function crypt (x) ' poor mans encryption ;-)
	Dim i, s

	for i = 1 to len(x)
		s = s & hex(asc(mid(x,i,1))-30+i)
	next
	crypt = s
end function

function decrypt (x)
	dim s, k, i, h
	s = ""
	k = 30
	for i = 1 to len(x) step 2
		h =  Int("&H" & mid(x,i,2))
		k = k - 1
		s = s & chr(h+k)
	next
	decrypt = s
end function

function ncrypt (x)
	ncrypt = crypt(CStr(Clng (x) * 18))
end function

function nohack(s)
	nohack = replace (replace(s,">","&gt;"),"<","&lt;")
end function

function nojhack(s)
	nojhack = replace(replace(replace(replace(s,"|",""),"""",""),"/","-"),"/","-")
end function

function ndecrypt (x)
	dim ntemp, ntemp2

	if x = "" then
		ndecrypt = 0
	else
		on error resume next
		ntemp = CLng(decrypt(x))
		if err then ntemp=-1234
		ntemp2 = CLng(ntemp / 18)
		if ntemp2 * 18 = ntemp then
			ndecrypt = ntemp2
		else
			ndecrypt = 0
		end if
	end if
end function

function xormail(s)
	dim x, i

	x = 0
	for i = 1 to len(s)
		x = x xor asc(mid(s,i,1))
	next
	xormail = hex(x)
	if len(xormail) = 1 then xormail = "0" & xormail
end function

' MenuItem(mtitle, murl, mimages)
'
' Show a single line of a menu.
' The mtitle parameter can be "-", then the output is "<BR>"
' The mtitle parameter can be a single top-level menu.
'     If it is equal to the global menuitem variable, then
'     mimages(0) is shown, otherwise mimages(1) is shown.
' The mtitle parameter can be a two-level menu separated by VBAR |
'     If the part after the | is equal to the global submenu variable,
'     then mimages(2) is shown, otherwise mimages(3) is shown.
' If murl is not empty, then a link is generated
'
Sub MenuItem(mtitle, murl, mimages)
	Dim B

	B = Instr(mtitle,"|")
	if (SUBMEN) or B = 0 then
		if mtitle <> "-" then
			response.write "<img "
			If B > 0 Then
				mtitle = right(mtitle,len(mtitle)-b)
				if mtitle = submenu then
					response.write mimages(2)
					murl=""
				else
					response.write mimages(3)
				end if
			ElseIf mtitle = mainmenu then
				response.write mimages(0)
				SUBMEN = TRUE
			Else
				response.write mimages(1)
				SUBMEN = FALSE
			End If

			response.write ">&nbsp;&nbsp;"

			if murl = "" then
				response.write mtitle
			else
				if mtitle = "HELP" then
					response.write "<a style=""cursor:help"" href=""" & murl & """>" & mtitle & "</a>"
				else
					response.write "<a href=""" & murl & """>" & mtitle & "</a>"
				end if
			end if
		end if

		response.write "<br>" & vbCrLf
	end if
End Sub

' ShowMenu(menu, mimages)
'
' Show a menu.  The menu parameter is an array of
' title, url values.  The mimages parameter is an
' array of 4 images (top level current, top level,
' submenu current, submenu).
'
Sub ShowMenu(menu, mimages)
	Dim i
	For i = 0 to UBound(menu)-1 step 2
		MenuItem menu(i), menu(i + 1), mimages
	next

	if Session("ID") = "" or Session("ID") = "0" then
		MenuItem "LOG IN", "../common/login.asp", mimages
	else
		MenuItem "LOG OUT", "../common/login.asp?logout=true", mimages
	end if
End Sub

%>
