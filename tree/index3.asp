<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->

<%
Dim	email, pass, rs, person, publication

email = request.form("email")
pass = request.form("pass")
publication = Request.QueryString("publication")

If publication <> "" then
	

ElseIf ((email <> "") or ((ndecrypt(Request.cookies("pid"))<> 0) or (ndecrypt(Request.cookies("publication")) <> 0))) Then
	If email <> "" then
		on error resume next
		set rs = conn.execute ("SELECT personid, pass FROM person WHERE email='" & strip(email) & "'")
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		If Not rs.eof Then
			If pass = rs("pass") Then 
				person = rs("personid")
			' must do something here for incorrect password or login	
			else

			End If	
		End If
		rs.close
	Else 
		person = ndecrypt(Request.cookies("pid"))
		publication = ndecrypt(Request.Cookies("publication"))
		person=2
	End If
%>
	<html>
	<head>
	<title>syncit test</title>
	<!--
	<script language="Javascript">
		alert("<%=person%>");
	</script>
	-->
	<script language="JavaScript" src="tocTab2.asp?person=<%=ncrypt(person)%>&publication=<%=ncrypt(publication)%>"></script>
	<script language="JavaScript" src="tocParas.js"></script>
	<script language="JavaScript" src="displayToc3.js"></script>
	<script language="JavaScript">
	<!--
	function newtree(f) {
	//	toc.document.location ='index2.asp?pubid=' + f.pubid.options[f.pubid.selectedIndex].value;
		tocTab.length = 0;
		 
	}
	
	document.cookie = "pid=<%=ncrypt(person)%>";
	document.cookie = "publication=<%=ncrypt(publication)%>";
	
	function Logout() {
		var yesteryear = new Date();
		yesteryear.setYear(1999);
		parent.document.cookie = "pid=blah;expires=yesteryear.toGMTString";
		parent.document.cookie = "publication=;expires=yesteryear.toGMTString";
		toc.location.href = "index3.asp";	
	}
	//-->
	</script>
	<frameset border=0 onload="reDisplay('0',true);">
	<frame src="blank.htm" name="toc">
	</frameset>
	</head>
	
	<%
	' cols="200,*"
'	<br>
'	<form method="GET" action="index2.asp" name="someform">
'	<select size="1" class="f11" name="publication" onchange="newtree(this.form);">
	

'		dim sql
	
'		response.write "<option value=0"
'		if ndecrypt(publication) = 0 then response.write " selected"
'		response.write ">My Bookmarks</option>"
'		sql = "SELECT publish.PublishID, publish.Title FROM subscriptions,publish where publish.PublishID = subscriptions.publish_id and (subscriptions.person_id)=" & person
'		on error resume next
'		set rs = conn.execute (sql)
'		if err then dberror err.description & " [" & err.number & "]"
'		on error goto 0
'		while not rs.eof
'			Dim pid
'	
'			response.write "<option"
'			pid = ncrypt (rs("publishid"))
'			if pid = publication then response.write " selected"
'			response.write " value=""" & pid & """>" & rs("title") & "</option>" & vbCrLf
'			rs.movenext
'		wend
'	
'		response.write "</select><br></form>" & vbCrLf
	%>

<%
Else 
%>
	<html>
	<head>	
	<title>BookmarkSync</title>
	<meta http-equiv="content-type" content="text/html; charset=us-ascii">
	</head>
	<table>
	<tr align="center">
	    <td><img src="../images/logo1.gif" width="112" height="52" alt="SyncIT.com"></td>
	</tr>
	</table>
	<form method="POST" action="index3.asp" name=form1>
	<img src="tinyemail.gif"><br>
						<input type="text" name="email" class="login" size="12" maxlength="50">
					<br>
	<img src="tinypass.gif"><br>
						<input type="password" name="pass" class="login" size="12" maxlength="50">
	
	<p><input border="0" src="../images/btnlogin.gif" name="I1" type="image" width="53" height="16" alt="Login"></p>
	</form>
	
<%	
End if
%>
