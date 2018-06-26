<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<%
Dim email, udeleted, sql, rs, reason, popup

mainmenu = "PROFILE"
submenu  = "REMOVE ME"
email = trim(request.form("email"))
udeleted=0
if email = "" then
	email = trim(Request.Cookies("email"))
else
	reason = trim(request.form("reason"))
	if request.form("doit") = "ON" then
		SQL =  "SELECT personid, name, pass, pc_license,mac_license,linux_license FROM person WHERE email = '" & replace(email,"'","''") & "'"
		on error resume next
		set rs = conn.execute (sql)
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		If rs.eof Then
			udeleted = 1
		ElseIf Request("license") = "" and not (IsNull(rs("pc_license")) and IsNull(rs("mac_license")) and IsNull(rs("linux_license"))) then
			Session("license")=true
			udeleted=3
		ElseIf Request.Form("pass") <> rs("pass") Then
			udeleted = 1
		Else
			Dim name

			
			ID = rs("personid")
			name = rs("name")
			on error resume next
			set rs = conn.execute ("SELECT publishid FROM publish WHERE user_id = " & ID)
			if err then dberror err.description & " [" & err.number & "]"
			on error goto 0
			while not rs.eof
				connexecute "DELETE FROM fmessages WHERE forum_id = " & rs("publishid")
				rs.movenext
			wend
			connexecute "{ CALL removeuser(" & ID & ", '" & replace(reason,"'","''") & "') }"
			ID = ""
			Response.Cookies("email").Expires = Date - 365
			Response.Cookies("MD5").Expires = Date - 365
			Response.Cookies("email") = ""
			Response.Cookies("MD5") = ""
			Session("ID") = ""
			on error goto 0
			Session.Abandon

			domail "support@bookmarksync.com",email,"BookmarkSync account removed",name & "," & vbCrLf & "Your BookmarkSync user account <" & email & "> has been removed at " & now() & vbCrLf
			if reason <> "" then domail email, "removed@bookmarksync.com", "Removed user: " & name, reason

			udeleted = 2
		end if
		rs.close
	else
		popup = true
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>Remove User</title>
<link rel="StyleSheet" href="syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#660066" vlink="#990000" alink="#CC0099">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="117" valign="top">
<!-- #include file = "../inc/bmenu.inc" -->
    </td>
    <td width="513" valign="top">
      <table width="513" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="2" bgcolor="#6633FF" width="503"><a href="../bms/"><img src="../images/horbtn_bmup.gif" width="120" height="18" border="0" alt="BookmarkSync"></a><a href="../collections/"><img src="../images/horbtn_collup.gif" width="147" height="18" border="0" alt="MySync Collections"></a><a href="../news/"><img src="../images/horbtn_newsup.gif" width="124" height="18" border="0" alt="QuickSync News"></a><a href="../express/"><img border="0" src="../images/horbtn_exup.gif" width="112" height="18" alt="SyncIT Express"></a></td>
          <td rowspan="2" width="10" valign="top"><img src="../images/rightbar.gif" width="10" height="55" alt=""></td>
        </tr>
        <tr>
          <td colspan="2" width="503" height="37" bgcolor="#000066"><img src="../images/heading_remove.gif" width="503" height="37" alt="SyncIT Remove User"></td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
<%
if udeleted = 2 then
	response.write " <br><h2>Goodbye !</h2>"
	response.write "You have removed your profile and all your bookmarks from our database."
else
	response.write " <br><h2>Remove User</h2>"
	if udeleted = 1 then
 		 response.write "<p><b><font color=#FF0000>Wrong e-mail/password combination !</font></b></p>"
	end if
	if udeleted = 3 then
 		 response.write "<p><b><font color=#FF0000>You have a payed SyncIT license. <br>If you remove yourself your license will forfeit</font></b></p>"
	end if
%>
		<h4>Please re-enter your e-mail address and password to confirm your cancellation.</h4>
		<p>Once you remove yourself all your bookmarks on our server will be removed...
		and there is no way of getting them back!</p>
		<p><b>Note: You do not have to remove yourself from the database if you just want to
		reinstall a new client. - <a href="/download/">Download here</a>.</b></p>
		<h4>Your local bookmarks and favorites will not be affected.</h4>
<script language="JavaScript"><!--
function validate(theForm)
{

  if (theForm.email.value == "")
  {
    alert("Please enter a value for the 'E-mail Address' field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.pass.value == "")
  {
    alert("Please enter a value for the 'Password' field.");
    theForm.pass.focus();
    return (false);
  }
  return (true);
}
//--></script><form method="POST" action="removeuser.asp" onsubmit="return validate(this)">
           <table border="0" width="485">
            <tr>
              <td width="50%" align="right" class="menub">E-mail:</td>
              <td width="50%" colspan="2"><input type="text" name="email" size="20" class="register" value="<%= email %>"></td>
            </tr>
            <tr>
              <td width="50%" align="right" class="menub">Password:</td>
              <td width="50%" colspan="2"><input type="password" name="pass" size="20" class="register"></td>
            </tr>
            <tr>
              <td width="50%" align="right" class="menub">Could you please tell us the reason
                for leaving so we can improve it for future members?</td>
              <td width="50%" colspan="2"><textarea rows="6" name="reason" cols="20" class="register"><%= reason %></textarea></td>
            </tr>
            <% if Session("license") = true then %>
            <tr>
              <td align="right" class="menub"><font color='red'>I forfeit my license(s):</a></td>
              <td><input type="checkbox" name="license" value="<%= Session("license") %>"></td>
              <td></td>
            </tr>          
            <% end if %>
            <tr>
              <td width="50%" align="right" class="menub">I really want to do this:</td>
              <td width="25%"><input type="checkbox" name="doit" value="ON"></td>
              <td width="25%" align="right"><input border="0" src="../images/btn_remove.gif" name="I2" type="image" width="62" height="16" alt="Remove"></td>
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
<!-- #include file = "../inc/footer.inc" -->
</body>
<%
if popup then
	response.write "<script>"& vbCrLf & "<!--" & vbCrLf
	response.write "alert('You must check ""I REALLY WANT TO DO THIS"" to remove yourself');" & vbCrLf
	response.write "//-->" & vbCrLf & "</script>"
end if
%>
</html>
