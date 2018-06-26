<%
Dim conn, ID, NewsID

If RunningLocal() Then
	NewsID = 2
Else
	NewsID = 694
End If

Function RunningLocal()
	dim testip
	
	if not DoUpload then testip = left(Request("REMOTE_ADDR"), 3)
	RunningLocal = (testip = "10." or testip = "127")
End Function


' myredirect(url)
'
' Close the database connection, then redirect to 
' another page.
' 
sub myredirect (url)
	on error resume next	
	conn.close
	set conn=Nothing
	response.redirect (url)
	response.end
end sub

' MD5Pwd()
'
' The MD5Pwd() function returns the base-64 encoded MD5 hash
' of the user's e-mail address and password combined with a
' site-specific hash.
'
Function MD5Pwd(email, pass)
   Dim md5

   Set md5 = Server.CreateObject("FishX.MD5")
   md5.Init()
   md5.Update("syncit.com")   ' site-specific salt
   md5.Update(email)
   md5.Update(pass)

   MD5Pwd = md5.Final

   Set md5 = Nothing
End Function

' GetID()
'
' The GetID function attempts to retrieve the currently logged-in
' user's personid.
'
' If the user is currently logged in and has a current session, then
' the "ID" session variable is returned.
'
' If the user's session has timed out (say, perhaps the user stepped
' away from the computer, or has bookmarked the page), then a log
' in is attempted using the 'email' and 'pass' cookies.  If these
' cookies are present, email is a valid entry in the person table,
' and pass is a valid MD5 hash of the email and *real* password,
' then the personid value is returned.
'
' Otherwise, an empty string is returned.
'
Function GetID()
	dim email, pass, id, rs

	id = CLng(("0" & Session("ID")))

	if id = 0 then
		email = Request.cookies("email")
		pass  = Request.cookies("md5")
		If pass <> "" Then
			set rs = conn.execute ("SELECT pass, personid, name, partner_id,token,pc_license,mac_license,linux_license,osx_license,daysregistered=DATEDIFF(day,registered,GETDATE()) FROM person WHERE email='" & email & "'")

			If Not rs.eof Then
				If pass = MD5Pwd(email, rs("pass")) Then
					id = rs("personid")
					Session("ID") = id
					Session("name") = rs("name")
					Session("email") = email
					Session("partner_id") = rs("partner_id") & ""
					Session("token") = CLng("0" & rs("token"))
					Session("license") = not (IsNull(rs("pc_license")) and IsNull(rs("mac_license")) and IsNull(rs("linux_license")) and IsNull(rs("osx_license")))
					Session("daysregistered") = rs("daysregistered")
					if Session("license") then Session("popup") = "true"
				End If
			End If

			rs.close
			Set rs = Nothing
		End If
	End If

	GetID = id
End Function

' checkuser()
' Place in front of a web-page that isn't password protected.
'
' This function sets the ID global variable to the personid
' of the logged-in user.  The ID variable is ideally kept in
' a session variable, but if the user's session has timed out
' (say, perhaps the user stepped away from the computer, or
' has bookmarked the page), the the ID variable is retrieved
' if the email, pass cookies are set, and are correct.
'
' The ID variable is set if logged in, or 0 if not.
'
Sub checkuser()
	ID = GetID()
End Sub

' restricted(url)
' Place in front of a web page that should be password protected.
' The url parameter is the relative URL of the page being protected.
'
' The basic upshot is that the user is allowed access (the subroutine
' returns normally) if:
' 1) The user is already in a 'session' (the "ID" session variable is
'    currently set; or
' 2) The user has the 'email' and 'pass' cookies set, and they're
'    correct.  The pass is really an MD5 hash of the email and *real*
'    password.
'
' All other conditions (no cookies, email address doesn't refer to a
' person in the BookmarkSync registration db, or the MD5 password is
' incorrect) result in a redirection to login.asp, with a parameter
' so login.asp can redirect back to the proper page when done.
'
Sub restricted(url)
	ID = GetID()

	If ID = 0 Then
		myredirect "../common/login.asp?refer=" & url
	End If
	if Session("token") > 20 and Session("daysregistered") > 30 and (Session("license") = false) then myredirect "../common/presales.asp?expired=true&token=" & Session("token")
End Sub

sub dberror(errstr)
	response.write "Our database is experiencing temporary problems.<br>The site administrator has been notified.<hr>"	
	response.write "Please try again shortly..."
	response.write "<br><h6><font color=#FFFFFF>" & errstr & "</font></h6>"
	response.write errstr
	response.AppendToLog errstr
	conn.close
	set conn=Nothing
	'response.status = "500 Database error"
	response.end
end sub	

sub connexecute (sql)
	Dim sql_result
	on error resume next
	set sql_result = conn.execute (sql)
	if err then dberror err.description & " [" & err.number & "]" & sql
	on error goto 0
end sub

'********************************************
'* Actual Code to run
'********************************************
on error resume next
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "syncit","syncit","syncit"
if err then
	response.write "Our database is temporarily down. <br>The site administrator has been notified.<hr>"	
	response.write "Please try again shortly...<br><h6><font color=#FFFFFF>" &  err.description & " [" & hex(err.number) & "]</font></h6>"
	response.AppendToLog err.description & " [" & hex(err.number) & "]"
	conn.close
	set conn=Nothing
	response.status = "500 Database error"
	response.end
end if
on error goto 0
checkuser()
%>
