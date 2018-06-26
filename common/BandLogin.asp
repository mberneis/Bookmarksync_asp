<%@ Language=VBScript %>
<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<%
Dim email, pass, errmsg, url, sendpassemail

if request.querystring("logout") <> "" then
	Response.Cookies("md5") = ""
	Response.Cookies("email") = ""
	Response.Cookies("pass") = ""
	Response.Cookies("ID") = ""
	Response.Cookies("md5").Expires = Date - 1000
	Response.Cookies("email").Expires = Date - 1000
	Response.Cookies("pass").Expires = Date - 1000
	Response.Cookies("ID").Expires = Date - 1000
	Session.abandon
	Session("ID") = ""
	Session("email") = ""
	Session("pass") = ""
	Session("partner_id") = ""
	Session("partnerurl") = ""
	Session("partnerlogo") = ""
	ID = 0
	errmsg = "LOGGED OUT"
	email=""
end if

md5 = Request.Cookies("md5") 
mainmenu = "LOG IN"
submenu  = ""
email = request.form("email")
pass = request.form("pass")
errmsg = request.querystring("err")
sendpassemail = request.querystring("sendpass")
If sendpassemail <> "" Then
	dim Subj, body, sql, rs

	
	Subj = "Your SyncIT Password"
	body = "You have requested that your SyncIT password be sent to you" & vbCrLf
	sql = "SELECT pass FROM person WHERE email='" & sendpassemail & "'"
	on error resume next
	set rs = conn.execute (sql)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	If rs.eof Then
		body = body & "Your e-mail address could not be found" & vbCrLf
		body = body & "If you have not registered yet please visit http://syncit.com/common/login.asp" & vbCrLf
	Else
		body = body & "Password: " & rs("pass")& vbCrLf
	End If

	rs.close
	errmsg = "Your account information has been sent"
	body = body & "Your BookmarkSync Support Team" & vbCrLf

	domail "support@bookmarksync.com",sendpassemail,subj,body
End If

url = Request.QueryString("refer")
if url <> "" and errmsg = "" then errmsg = "RESTRICTED ACCESS"

If email <> "" Then
	on error resume next
	Set rs = conn.execute ("SELECT personid, name, email, pass,partner_id FROM person WHERE email = '" & strip(email) & "'")
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	If rs.eof Then
		ID = 0
	ElseIf rs("pass") <> pass Then
		ID = 0
	Else
		ID = rs("personid")
	End If

	If ID > 0 Then
		Session("ID") = ID
		Session("name") = rs("name")
		Session("email") = rs("email")
		Response.Cookies("email") = email
		If request.form ("savepass") = "ON" then
			Dim md5
			set md5 = Server.CreateObject("FishX.MD5")
			md5.init
			md5.Update("syncit.com") ' site-specific salt
			md5.Update(email)
			md5.Update(rs("pass"))
			Response.Cookies("email").Expires = Date + 365
			Response.Cookies("md5")=md5.Final
			Response.Cookies("md5").Expires = Date + 365
			'response.write "Saved pass = " & md5.Final
			set md5 = Nothing
		End If
		Session ("partner_id") = rs("partner_id") 
		rs.close
		conn.close
		Set conn=Nothing
		response.redirect request.form("url")
		response.Flush
		response.end
	else
		errmsg = "INCORRECT LOGIN"
	end if
end if
if email = "" then email = request.cookies("email")
if url = "" then url = "../bms/"
%>
<HTML>
<HEAD>
<title>BookmarkSync WebBand</title>
</HEAD>
<BODY>
<script language="JavaScript">
<!--
function sendpass() {
	email = band_login.email.value;
	if (email.indexOf('.') == -1 || email.indexOf('@') == -1) {
		alert ("Please enter a valid E-mail address");
		band_login.email.focus();
	} else {
		location.href='login.asp?sendpass=' + email;
	}
}
	
function validate_login(theForm)
{

  if (theForm.email.value == "") {
    alert("Please enter a value for the 'e-mail' field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.email.value.length > 50) {
    alert("Please enter at most 50 characters in the 'e-mail' field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.pass.value.length > 50) {
    alert("Please enter at most 50 characters in the 'pass' field.");
    theForm.pass.focus();
    return (false);
  }
  return (true);
}
//-->
</script>
<table border="0" cellpadding="0" cellspacing="0" width="120">
	<tr>
		<td width="120">
			<img src="../images/logo1.gif" width="113" height="52">
		</td>
	</tr>
	<tr>
		<td width="120">
			<form action="BandLogin.asp" method="Post" name="band_login" onSubmit="validate_login(this)">
			<input type="hidden" name="url" value="<%= url %>">
			Please enter your email address and password below<br>
			Email:<br>
			<input type="text" name="email" class="login" size="15" maxlength="50" value="<%= email %>" onkeypress="if (event.which==13) {document.login_form.pass.focus()}"><br>
			Password:<br>
			<input type="password" name="pass" class="login" size="15" maxlength="50" onkeypress="if (event.which == 13) {document.login_form.submit()}"><br>
			Would you like to save your password to enable automatic login in the future? <input type=checkbox name="savepass" value="ON" checked> Yes!<br>
			If you forgot your password, please	enter your email address and <a href="javascript:sendpass()">Click Here.</a> Your password will be sent to the email address typed above.
			</td></tr></table>
</BODY>
</HTML>
