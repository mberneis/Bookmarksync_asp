<!-- #include file = "../inc/asp.inc" -->
<% response.expires=-1 %>
<!-- #include file = "../inc/db.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<!-- #include file = "../inc/tree.inc" -->
<html>
<head>
<title>SyncIT Remote</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
<%
Dim target
target= Session("target")
if request("target") <> "" then
	target = request("target")
	Session("target") = target
end if
if target="" then 
	target="_main"
	Session("target") = target
end if
response.write "<base target=""" & target & """>"
%>
</head>

<body bgcolor="#FFFFFF">
<%
Dim pubid
Dim email, pass, rs

pubid = request.form("pubid")
ID = ndecrypt (request.form("XID"))
if ID = 0 then ID = CLng("0" & (Session("RID")))
if ID = 0 then
	ID = Session("ID") + 0
end if
email = request.form("email")
if ID = 0 or email <> "" then
	pass = request.form("pass")
	Session("ID") = 0
	Session("RID") = 0
	If email <> "" Then
		on error resume next
		set rs = conn.execute ("SELECT personid, pass FROM person WHERE email='" & strip(email) & "'")
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		If Not rs.eof Then
			If pass = rs("pass") Then 
				Session("RID") = rs("personid")
			else
				response.write "<font color=#FF0000><b>Invalid email or password</b></font><hr size=1>"
			end if
		else
			response.write "<font color=#FF0000><b>Invalid email or password</b></font><hr size=1>"
		End If
		rs.close
	else
		email = Request.Cookies("email")
	end if
	ID = Session("RID") + 0
else
	Session("RID") = ID + 0
end if

If ID <>0 Then
	on error resume next
	Set rs = conn.execute ("SELECT name, email FROM person WHERE personid = " & ID)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	If Not rs.eof Then
		Session("RID") = ID
	End If

	rs.close
End If

If ID = 0 or request.querystring("logout") <> "" Then
%>
<table>
  <tr align="center">
    <td><a href="/"><img border=0 src="../images/logo1.gif" width="112" height="52" alt="SyncIT.com"></a></td>
  </tr>
</table>
<form method="POST" action="remote.asp" target="_self">
  <span class="menub">E-MAIL:</span><br>
                    <input type="text" name="email" class="login" size="12" maxlength="50" value="<%= email %>">
                  <br>
  <span class="menub">PASSWORD:</span><br>
                    <input type="password" name="pass" class="login" size="12" maxlength="50">

  <p><input border="0" src="../images/btnlogin.gif" name="I1" type="image" width="53" height="16" alt="Login"></p>
</form>
<% Else %>
<form method="POST" action="remote.asp" name="frmsel" target="_self">
<table>
  <tr align="center">
    <td><a target="_top" href="/"><img src="../images/logo1.gif" width="112" height="52" border="0" alt="SyncIT.com"></a></td> 
    <td><a href="remote.asp?logout=true" target="_top" class="menub">CHANGE USER</a><br>
    <% if target <> "_main" then %>
    <script language=javascript>
    <!--
    	if (parent.location.href != window.location.href) {
    		document.write ('<a href="../bms/default.asp" target="_top" class="menub">CLOSE WINDOW</a>');
    	} else {
    		document.write ('<a href="javascript:window.close();" target="_top" class="menub">CLOSE WINDOW</a>');   		
    	}
    //-->
    </script>
    <% end if %>
    </td>
  </tr>
</table>
<br>
<input type=hidden name="XID" value="<%= ncrypt (ID) %>">
<select size="1" class="f11" name="pubid" onChange="document.frmsel.submit();">
	<%
	dim sql
	response.write "<option value=0"
	if ndecrypt(pubid) = 0 then response.write " selected"
	response.write ">My Bookmarks</option>"
	sql = "SELECT publish.PublishID, publish.Title FROM subscriptions,publish where publish.PublishID = subscriptions.publish_id and (subscriptions.person_id)=" & ID
	on error resume next
	set rs = conn.execute (sql)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	while not rs.eof
		Dim pid

		response.write "<option"
		pid = ncrypt (rs("publishid"))
		if pid = pubid then response.write " selected"
		response.write " value=""" & pid & """>" & rs("title") & "</option>" & vbCrLf
		rs.movenext
	wend

	response.write "</select><br></form><nobr>" & vbCrLf

	tree pubid,false
end if

conn.close
set conn=Nothing
%>
</body>




