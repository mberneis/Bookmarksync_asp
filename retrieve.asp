<!-- #include file = "../common/asp.inc" -->
<!-- #include file = "../common/db.inc" -->
<!-- #include file = "../common/mail.inc" -->
<%
Dim action, valfolder, email, pass, code, title, activity, fsize
Dim eid, rs, p, dist, path, fp, i, subj, ps, msg, sql, distribution

mainmenu = "RETRIEVE"
submenu  = ""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>SyncIT Express: Retrieve Delivery</title>
<link rel="StyleSheet" href="../common/syncit.css" type="text/css">
</head>
<script language="JavaScript">
<!--
// Validate Form
//~~~~~~~~~~~~~~
function checkit(f) {
	if (f.email.value.indexOf('.') == -1 || f.email.value.indexOf('@') == -1) {
		alert ("Please enter a valid E-mail address");
		f.email.focus();
		return false;
	}

	return true;
}
//-->
</script>

<body bgcolor="#FFFFFF" text="#000000" link="#3300CC" vlink="#990000" alink="#CC0099">
<!-- #include file = "../common/header.inc" -->
<table width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="117" valign="top">
<!-- #include file = "emenu.inc" -->
    </td>
    <td width="513" valign="top">
      <table width="513" border="0" cellspacing="0" cellpadding="0">
<!-- #include file = "ehead.inc" -->
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br><h2>Retrieve Your File Here</h2>
<%
	if request.querystring("err") = "1" then
		action=6
	else
		valfolder = "../"
		action = 0
		email = request.form("email")
		if email = "" then
			code = request.querystring("code")
			eid = ndecrypt (code)
			on error resume next
			set rs = conn.execute ("Select distribution,fsize,filename,password from express where expressid=" & eid)
			if err then dberror err.description & " [" & err.number & "]"
			on error goto 0
			if rs.eof then
				action = 1
			else
				action = 2
				distribution = rs("distribution")
				fsize = rs("fsize")
				fp = rs("filename")
				do
					I = Instr(fp,"\")
					if i = 0 then exit do
					fp = right(fp,len(fp)-i)
				loop
				pass=rs("password")
			end if
			rs.close
		else
			code = request.form("code")
			eid = ndecrypt(code)
			email = lcase(trim(request.form("email")))
			p = trim(request.form("pass"))
			on error resume next
			set rs = conn.execute( "Select distribution,password,filename,sender_id,title,activity from express where expressid=" & eid)
			if err then dberror err.description & " [" & err.number & "]"
			on error goto 0
			dist = lcase(rs("distribution"))
			if Instr(dist,email) =< 0 then
				action = 3
			else
				pass = rs("password")
				if pass <> "" and p <> pass then
					action = 4
				else
					path = rs("filename")
					fp = path
					do
						I = Instr(fp,"\")
						if i = 0 then exit do
						fp = right(fp,len(fp)-i)
					loop
					subj = "SyncIT File Delivery Receipt"
					on error resume next
					set ps = conn.execute ("Select name,email from person where personid = " & rs("sender_id"))
					if err then dberror err.description & " [" & err.number & "]"
					on error goto 0
					msg = ps("name") & "," & vbCrLf
					msg = msg & "Your Delivery was retrieved: " & vbCrLf
					msg = msg & "   File: " & fp & vbCrLf
					msg = msg & "   Time: " & Now() & vbCrLf
					msg = msg & "   Recipient: " & email & vbCrLf
					msg = msg & "-._.-._.-._.-._.-._.-._.-._.-._.-._.-._.-._.-._.-._.-._.-._.-._." & vbCrLf
					msg = msg & "Please keep in touch to let us know what you think of our service." & vbCrLf & vbCrLf
					msg = msg & "Your BookmarkSync support team." & vbCrLf
					msg = msg & "support@bookmarksync.com"
					domail "support@bookmarksync.com",ps("email"),subj,msg
					title = rs("title")
					activity = rs("activity") &  "<br>Downloaded at " & Now() & " from " & email
					sql = "update express set activity = '" & activity & "' where expressid=" & eid
					'response.write (sql)
					connexecute (sql)
					ps.close
					action = 5
					rs.close
					Session("syncitexpresspath") = path
				end if
			end if
		end if
	end if
	if email = "" then email = Session("email")
	if email = "" then email = request.cookies("email")

If action = 1 Then
	If code <> "" Then
		response.write "<h4><font color=""#FF0000"">No delivery available.</font></h4>" & vbCrLf
		response.write "<p>The delivery might have been expired or the tracking number might be invalid.<br>" & vbCrLf
		response.write "Please contact the original sender."
	End If
ElseIf action = 2 Then
	response.write "<h4>Please fill out the form to retrieve your delivery.</h4>" & vbCrLf
ElseIf action = 3 Then
	response.write "<h4><b><font color=""#FF0000"">Your e-mail address does not match</font></b></h4>" & vbCrLf
ElseIf action = 4 Then
	response.write "<h4><b><font color=""#FF0000"">Your password does not match</font></b></h4>" & vbCrLf
ElseIf action = 5 Then
	response.write "<h3><b>Thank you for using SyncIT Express.</b></h3>" & vbCrLf
	response.write "<p><b>Retrieve your document here: <a href=""download.asp"">"
	response.write title
	response.write "</a></p>" & vbCrLf
ElseIf action = 6 Then
	response.write "<h4><b><font color=""#FF0000"">The delivery does not exist anymore</font></b></h4>" & vbCrLf
End If
if action < 5 then %>
            <form method="post" action="retrieve.asp" name="frmupload" onsubmit="return checkit(this)">
            <table border="0">
              <tr>
                <td  class="menub" align="right">Tracking #</td>
                <td class="register"><input type="text" name="code" size="30" class="register" value="<%= code %>"></td>
              </tr>
              <tr>
                <td class="menub" align="right">Your E-mail</td>
                <td class="register"><input type="text" name="email" size="30" class="register" value="<%= email %>"></td>
              </tr>
<% if pass <> "" or action=1 then %>
              <tr>
                <td class="menub" align="right">Password</td>
                <td class="register"><b><input type="password" name="pass" size="20" class="register"></b></td>
              </tr>
<% end if %>
              <tr>
                <td>&nbsp;</td>
                <td><input border="0" src="../images/btnok.gif" name="I1" type="image" width="62" height="16" alt="OK"></td>
              </tr>
            </table>
	</form>
<% end if %>
            </td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../common/footer.inc" -->
</body>
</html>














