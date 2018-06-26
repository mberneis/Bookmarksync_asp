<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<%
Dim email, pass, errmsg, refer, sendpassemail

if request.querystring("logout") <> "" then
	Response.Cookies("md5") = ""
	Response.Cookies("email") = ""
	Response.Cookies("pass") = ""
	Response.Cookies("ID") = ""
	Response.Cookies("name") = ""
	Response.Cookies("name").Expires = Date - 1000
	Response.Cookies("md5").Expires = Date - 1000
	Response.Cookies("email").Expires = Date - 1000
	Response.Cookies("pass").Expires = Date - 1000
	Response.Cookies("ID").Expires = Date - 1000
	Session.abandon
	Session("ID") = ""
	Session("email") = ""
	Session("name") = ""
	Session("pass") = ""
	Session("partner_id") = ""
	Session("partnerurl") = ""
	Session("partnerlogo") = ""
	Session("license") = ""
	ID = 0
	errmsg = "LOGGED OUT"
	email=""
	response.redirect "../default.asp?logout=true"
end if


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
	sql = "SELECT pass FROM person WHERE email='" & strip(sendpassemail) & "'"
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

	domail "mailer@syncit.com",sendpassemail,subj,body
End If

refer = Request("refer")
if refer <> "" and errmsg = "" then errmsg = "RESTRICTED ACCESS"

If email <> "" Then
	on error resume next
	Set rs = conn.execute ("SELECT personid, name, email, pass,partner_id,token FROM person WHERE email = '" & strip(email) & "'")
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	If rs.eof Then
		ID = 0
	ElseIf rs("pass") <> pass Then
		ID = 0
	Else
		ID = rs("personid")
	End If

	if refer = "" then refer = "../tree/view.asp"

	If ID > 0 Then
		Session("ID") = ID
		Session("name") = rs("name")
		Session("email") = rs("email")
		Response.Cookies("email") = email
		If request.form ("savepass") = "ON" then
			Response.Cookies("email").Expires = Date + 365
			Response.Cookies("name") = rs("name")
			Response.Cookies("name").Expires = Date + 365
			Response.Cookies("md5")=MD5Pwd(email, rs("pass"))
			Response.Cookies("md5").Expires = Date + 365
		End If
		Session ("partner_id") = rs("partner_id") 
		Session("token") = CLng("0" & rs("token"))
		Session("license") = 1
		conn.close
		Set conn=Nothing
		response.redirect refer
		response.Flush
		response.end
	else
		errmsg = "INCORRECT LOGIN"
	end if
end if
if email = "" then email = request.cookies("email") & ""
if email = "null" then email = ""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>SyncIT Login and Registration</title>
<link rel="StyleSheet" href="syncit.css" type="text/css">
</head>
<script language="JavaScript">
<!--
	function sendpass() {
		email = login_form.email.value;
		if (email.indexOf('.') == -1 || email.indexOf('@') == -1) {
			alert ("Please enter a valid E-mail address");
			login_form.email.focus();
		} else {
			location.href='login.asp?sendpass=' + email;
		}
	}
//-->
</script>
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
          <td colspan="2" width="503" height="37" bgcolor="#000066"><img src="../images/header_login_register.gif" width="503" height="37" alt="SyncIT Login &amp; Registration"></td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
<script language="JavaScript"><!--
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
//--></script><form method="POST" action="login.asp" onsubmit="return validate_login(this)" name="login_form">
          <input type="hidden" name="refer" value="<%= refer %>">
              <table width="491" border="0" cellspacing="0" cellpadding="2">
                <tr>
                <td rowspan="4" width="173"><img src="../images/registerillustration.gif" width="159" height="127" alt="Secure Login"></td>
                <td colspan="2" width="318"><img src="../images/lbl_registeredusers.gif" width="318" height="37" vspace="4" alt="Registered Users"></td>
              </tr>
              <tr>
                <td width="133">
                  <h4>
                  <% if errmsg <> "" then response.write "<font color=""#FF0000"">" & errmsg & "</font><br><br>" %>
                  <b>Please log in <br>
                      with your <br>
                      E-mail address and password.</b>
                      </h4>
                </td>
                <td width="185" ><div class="menub">E-MAIL:</div><br>
                    <input type="text" name="email" class="login" size="15" maxlength="50" value="<%= email %>" onkeypress="if (event.which==13) {document.login_form.pass.focus()}">
                  <br>
                  <div class="menub">PASSWORD:</div><br>
                    <input type="password" name="pass" class="login" size="15" maxlength="50" onkeypress="if (event.which == 13) {document.login_form.submit()}">
                </td>
              </tr>
              <tr>
                <td width="133" class="f11">Check &quot;Save Password&quot; to
                  enable automatic <br>
                  log in for the future.</td>
                <td width="185" class="menub" valign="middle">
                    <input type="image" border="0" name="login_submit" src="../images/btnlogin.gif" width="53" height="16" align="absmiddle" alt="LOG IN">
                  <input type="checkbox" name="savepass" value="ON">
                  SAVE PASSWORD</td>
              </tr>
              <tr>
                <td width="133" class="f11"><b>Forgot the password?</b></td>
                <td width="185" class="f11" valign="middle">
                   <a href="javascript:sendpass()"><img border="0" src="../images/btn_sendpass.gif" width="102" height="16" alt="Send Password"></a><br>
                      Have it sent to the e-mail above.</td>
              </tr>
            </table>
	    </form>
            <p> <a name="register"> <img src="../images/lbl_registration.gif" width="491" height="35" alt="New Users Register Free">
            </a>
            </p>
            <p>Please register so we can reserve your personal database space.<br>
              <i><b>It is important that you enter your correct e-mail address.
              </b></i>We will send you a special URL to complete your registration.
              Your e-mail address will not be given out to 3rd parties.</p>
<script language="JavaScript"><!--
function validate_register(theForm)
{

  if (theForm.name.value == "")
  {
    alert("Please enter a value for the 'name' field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.name.value.length > 50)
  {
    alert("Please enter at most 50 characters in the 'name' field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.email.value == "")
  {
    alert("Please enter a value for the 'email' field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.email.value.length > 50)
  {
    alert("Please enter at most 50 characters in the 'email' field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.pass.value == "")
  {
    alert("Please enter a value for the 'pass' field.");
    theForm.pass.focus();
    return (false);
  }

  if (theForm.pass.value.length > 50)
  {
    alert("Please enter at most 50 characters in the 'pass' field.");
    theForm.pass.focus();
    return (false);
  }

  if (theForm.pass2.value == "")
  {
    alert("Please enter a value for the 'pass2' field.");
    theForm.pass2.focus();
    return (false);
  }

  if (theForm.pass2.value.length > 50)
  {
    alert("Please enter at most 50 characters in the 'pass2' field.");
    theForm.pass2.focus();
    return (false);
  }
  return (true);
}
//--></script><form method="POST" action="register.asp" onsubmit="return validate_register(this)" name="register_form">
            <table width="491" border="0" cellspacing="0" cellpadding="0">
              <tr>
                  <td width="173" align="right" class="menub">NICK NAME/ALIAS:&nbsp;
                  </td>
                <td width="318">
                    <input type="text" name="name" maxlength="50" class="register" size="30" onkeypress="if (event.which == 13) {document.register_form.email.focus()}">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">E-MAIL:&nbsp; </td>
                <td width="318">
                    <input type="text" name="email" maxlength="50" class="register" size="30" onkeypress="if (event.which == 13) {document.register_form.pass.focus()}">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">PASSWORD:&nbsp; </td>
                <td width="318">
                    <input type="password" name="pass" maxlength="50" class="register" size="30"  onkeypress="if (event.which == 13) {document.register_form.pass2.focus()}">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">RETYPE PASSWORD:&nbsp;
                  </td>
                <td width="318">
                    <input type="password" name="pass2" maxlength="50" class="register" size="30" onkeypress="if (event.which == 13) {document.register_form.submit()}">
                </td>
              </tr>
              <tr>
                <td></td><td width="318" align="right" valign="bottom">
                    <input type="image" border="0" name="imageField2" src="../images/btnagree.gif" width="62" height="16" alt="I Agree">
                    &nbsp; </td>
              </tr>
            </table>
            <input type="hidden" name="url" value="../download/">
            <input type="hidden" name="partnerid" value="<%= Session("partner_id") %>">
            </form>
            </td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>
</html>
