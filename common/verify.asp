<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<!-- #include file = "../inc/welcome.inc" -->
<%
mainmenu = "REGISTER"
submenu  = ""
dim email
email = request.form("email")
if email = "" then
	email = request.cookies("email")
else
	restricted ("../inc/verify.asp")
	Session("email") = email
	Response.Cookies("email") = email
	Response.Cookies("email").Expires = Date + 365
	domail "support@bookmarksync.com",email,"Welcome to BookmarkSync",welcometxt (Session("name"),ncrypt(ID))
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>SyncIT - E-mail verification</title>
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
          <td colspan="2" width="503" height="37" bgcolor="#000066"><img src="../images/header_login_register.gif" width="503" height="37" alt="SyncIT Login &amp; Registration"></td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
<!-- * -->
            <br><h2>E-mail Verification</h2>
            <%
            	Dim url,vid,cnt,rs
            	if request.form("email") <> "" then
            		response.write "We have sent you a new welcome letter to your e-mail address at <b>" & email & "</b>."
            	else
            	   vid = request.querystring("id")
            	  if vid="" then
            %>
This service is only available to SyncIT users who have verified their e-mail
address. In your initial welcome e-mail we provided a link to our site to
complete the e-mail verification. If you have not received this letter you may
have mis-spelled your e-mail address.&nbsp;<br>
Below you will find the e-mail address we currently have in our records. Please check
it carefully and make any necessary corrections.&nbsp;<br>
Then press <b>Resend Verification</b> and we will re-send you our
welcome letter.<script Language="JavaScript"><!--
function validate(theForm)
{

  if (theForm.email.value == "")
  {
    alert("Please enter a value for the \"E-mail\" field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.email.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"E-mail\" field.");
    theForm.email.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒšœŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖØÙÚÛÜİŞßàáâãäåæçèéêëìíîïğñòóôõöøùúûüışÿ0123456789-.@-_";
  var checkStr = theForm.email.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter, digit and \".@-_\" characters in the \"E-mail\" field.");
    theForm.email.focus();
    return (false);
  }
  return (true);
}
//--></script><form method="POST" action="verify.asp" onsubmit="return validate(this)" name="form1">
  <p><input type="text" name="email" size="30" maxlength="50" value="<%= email %>"><input type="submit" value="Resend Verification" name="B1" class="pbutton"></p>
</form>
<p>Upon receipt please click on the verification URL provided in our e-mail.&nbsp;
<%
            		else
            	   		on error resume next
	            	   set rs = conn.execute("update person set lastverified=getdate() where personid=" & ndecrypt (vid),cnt)
							if err then dberror err.description & " [" & err.number & "]"
							on error goto 0	            	   
   		         	   if rs.state = 0 and cnt=1 then
				%> Thank you - We have validated your e-mail verification and
have provided you with full access to our site. <% else %> We are sorry - Your e-mail
verification has failed. Please <a href="../help/support.asp"><b>contact our
support</b></a> with full details so we can correct the problem. <%
					end if
				end if            	      				         	   
			end if
%>
</p>
<!-- * -->
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
