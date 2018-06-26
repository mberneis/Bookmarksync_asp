<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<!-- #include file = "../inc/welcome.inc" -->
<%
Dim email, name, pass, rs, sql
Dim url, cnt, i,pid,refer,partner_pub
Dim address1, address2, city, state, zip, country
Dim phone, fax, gender, age, description, optin,wherefrom

mainmenu = "LOG IN"
submenu  = "REGISTER"

url = request.form("url")
refer = "&refer=" & url
email = request.form("email")
if email = "" then myredirect "../common/login.asp?err=No+e-mail+address" & refer
ID = request.form("RID")
if  ID = "" then
	name = request.form("name")
	pass = request.form("pass")
	if name="" then myredirect "../common/login.asp?err=No+user+name" & refer
	if pass="" then myredirect "../common/login.asp?err=No+password"& refer
	if pass <> request.form("pass2") then myredirect "../common/login.asp?err=passwords+do+not+match"& refer
	if instr(email,"@") = 0 or Instr(email,".") = 0 then myredirect "../common/login.asp?err=Invalid+e-mail+address"& refer
	on error resume next
	set rs = conn.execute ("select personid from person where email='" & email & "'")
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	if not rs.eof then
		rs.close
		myredirect "../common/login.asp?err=E-Mail address already registered." & refer
	end if
	sql = "Insert into person (name,email,pass,age,gender,registered,token) values ('" & strip(name) & "','" & strip(email) & "','" & strip(pass) & "',-1,'U',getDate(),0)"	
	connexecute (sql)
	on error resume next
	set rs = conn.execute ("select personid from person where email='" & strip(email) & "'")
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	ID = rs("personid")
	Session("email") = email
	session("ID") = ID
	session("name") = name
	rs.close

	partner_id = Session("partner_id")
	partner_pub = session("partnerpub")
	if int("0" & partner_id) > 0 then
		connexecute "update person set partner_id=" & partner_id & " where personid=" & ID
		'Check for default publication
		if partner_pub <> "" then connexecute "INSERT INTO subscriptions(person_id,publish_id,created) Values(" & ID & "," & partner_pub & ",getDate())"
	end if

	'Preloaded subscriptions...
	connexecute "INSERT INTO subscriptions(person_id,publish_id,created) SELECT " & ID & ", publishid, getdate() FROM publish, person where email='support@bookmarksync.com' AND personid=user_id"
	'connexecute "insert into subscriptions (person_id,publish_id,created) Values (" & ID & "," & InitialCategoryID & ",'" & Now() & "')"
	url = request.form("url")
	i = instr(url,"addid=")
	if i > 0 then
		pid = ndecrypt(right(url,len(url)-i - 5))
		connexecute "insert into subscriptions (person_id,publish_id,created) Values (" & ID & "," & pid & ",getdate())"
		url = ""
	end if
		
	Response.Cookies("email") = email
	Response.Cookies("email").Expires = Date + 365
	Session("email") = email
	domail "mailer@syncit.com",email,"Welcome to BookmarkSync",welcometxt (name,ncrypt(ID))
	
	on error resume next
	set rs = conn.execute ("select count(*) as cnt from person")
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	cnt = CLng(rs("cnt"))
	rs.close
	'if cnt mod 100 = 0 then
	'	on error resume next
	'	set rs = conn.execute ("SELECT COUNT(*) AS bmcnt FROM bookmarks")
	'	if err then dberror err.description & " [" & err.number & "]"
	'	on error goto 0
	'	bmcnt = rs("bmcnt")
	'	rs.close
	'	domail "Admin@Bookmarksync.com","alerts@syncit.com","User Count",cnt & " Users registered with " & bmcnt & " unique Bookmarks at " & Now()
	'end if
else

	restricted ("../common/profile.asp")
	ID=Session("ID")
	address1 		= strip(request.form("address1"))
	address2 		= strip(request.form("address2"))
	city 			= strip(request.form("city"))
	state 			= strip(request.form("state"))
	zip 			= strip(request.form("zip"))
	country 		= strip(request.form("country"))
	wherefrom 		= strip(request.form("wherefrom"))
	phone 			= strip(request.form("phone"))
	fax				= strip(request.form("fax"))
	gender 		= request.form("gender")
	age 			= request.form("age")
	description	= strip(request.form("description"))
	sql = "update person set "
	if email <> Session("email") then
		sql = sql & "lastverified=Null,"
	end if
	if request.form("optin") ="" then optin = 0 else optin = 1
	sql = sql & "optin=" & optin & ","
	sql = sql & "address1='" & address1 & "',"
	sql = sql & "address2='" & address2 & "',"
	sql = sql & "city='" & city & "',"
	sql = sql & "state='" & state & "',"
	sql = sql & "zip='" & zip & "',"
	sql = sql & "country='" & country & "',"
	sql = sql & "phone ='" & phone & "',"
	sql = sql & "fax ='" & fax & "',"
	sql = sql & "gender ='" & gender & "',"
	sql = sql & "age ='" & age & "',"
	sql = sql & "wherefrom ='" & wherefrom & "',"
	sql = sql & "description='" & description& "' "
	sql = sql & "where personid = " & ID
	'response.write (sql)
	
	connexecute (sql)
	if email <> Session("email") then
		domail "mailer@syncit.com",email,"Welcome back to BookmarkSync",welcometxt (name,ncrypt(ID))
	end if
	' Add here preloaded subscriptions
	'if age = "20" then
	'	'MyBytes Subscription
	'	connexecute "insert into subscriptions (person_id,publish_id,created) Values (" & ID & ",709920,'" & Now() & "')"
	'end if
	if url = "" then url="../download/default.asp"
	myredirect(url)
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>SyncIT Login and Registration</title>
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
          <td colspan="2" width="503" height="37" bgcolor="#000066"><img src="../images/heading_registration.gif" width="503" height="37" alt="SyncIT Registration"></td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br> <img src="../images/lbl_registration.gif" width="491" height="35" alt="New Users Register Free">

            <p>Although the following information is not required we would
            appreciate it if you could give us a little bit more information
            about yourself.<br>
            <i><b>It is important that you enter your correct e-mail address.
              </b></i>We will send you a special URL to complete your registration.
              Your e-mail address will not be given out to 3rd parties.</p>
<script language="JavaScript"><!--
function validate(theForm)
{

  if (theForm.email.value == "")
  {
    alert("Please enter a value for the 'email' field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.email.value.length > 50)
  {
    alert("Please enter at most 50 characters in the 'E-mail' field.");
    theForm.email.focus();
    return (false);
  }
  return (true);
}
//--></script><form method="POST" action="register.asp" onsubmit="return validate(this)">
		<input type="hidden" name="RID" value="<%= ID %>">
            <table width="491" border="0" cellspacing="0" cellpadding="0">
              <tr>
                  <td width="173" align="right" class="menub">E-Mail:
                  </td>
                <td width="318">
                    <input type="text" name="email" maxlength="50" class="register" size="33" values="<%= email %>" value="<%= email %>">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">Address:&nbsp; </td>
                <td width="318">
                    <input type="text" name="address1" class="register" size="33">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">Address:&nbsp; </td>
                <td width="318">
                    <input type="text" name="address2" maxlength="50" class="register" size="25">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">City:&nbsp; </td>
                <td width="318">
                    <input type="text" name="city" maxlength="50" class="register" size="25">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">State:&nbsp; </td>
                <td width="318">
                    <input type="text" name="state" maxlength="20" class="register" size="25">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">Postal Code:&nbsp; </td>
                <td width="318">
                    <input type="text" name="zip" maxlength="10" class="register" size="25">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">Country </td>
                <td width="318">
                    <input type="text" name="country" maxlength="30" class="register" size="25">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">Phone: </td>
                <td width="318">
                    <input type="text" name="phone" maxlength="50" class="register" size="25">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">Fax: </td>
                <td width="318">
                    <input type="text" name="fax" maxlength="50" class="register" size="25">
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">Age: </td>
                <td width="318">
                  <select name="age" size="1" class="register">
                    <option value="10">13-18</option>
                    <option value="20">19-25</option>
                    <option value="30">26-35</option>
                    <option value="40">36-49</option>
                    <option value="50">50-59</option>
                    <option value="60">60-69</option>
                    <option value="70">70 and over</option>
                    <option selected value="-1">Not specified</option>
                  </select>
                </td>
              </tr>
			 <tr>
                    <td colspan=2 class="menub"> 
                      <div align="right"><font color="#CC0000">IF YOU ARE UNDER 
                        13 PLEASE LET A PARENT OR GUARDIAN SIGN IN!</font></div>
                    </td>
                  </tr>
              <tr>
                  <td width="173" align="right" class="menub">Gender: </td>
                <td width="318">
                    <select name="gender" size="1" class="register">
                    <option value="F">Female</option>
                    <option value="M">Male</option>
                    <option selected value="U">Not specified</option>
                  </select>
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">Where did
                    you&nbsp;<br>
                    hear about us: </td>
                <td width="318">
                    <textarea name="wherefrom" rows="3" class="register" cols="32"></textarea>
                </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub">What are&nbsp;<br>
                    your interests:&nbsp; </td>
                <td width="318">
                    <textarea name="description" rows="3" class="register" cols="32"></textarea>
                </td>
              </tr>
              <tr>
                  <td width="491" align="right" class="menub" colspan="2">
                  <img border="0" src="../images/psep.gif" width="491" height="10" alt="">
                  </td>
              </tr>
              <tr>
                  <td width="173" align="right" class="menub" valign="top"></td>
                <td width="318" rowspan="2" class="menub">
                &nbsp; &lt;- Press OK to continue...
                </td>
              </tr>
              <tr>
                <td width="173" align="right" valign="bottom">
                    <input type="image" border="0" name="imageField2" src="../images/btnok.gif" width="62" height="16" alt="OK">
                    &nbsp; </td>
              </tr>
            </table>
            <input type="hidden" name="url" value="<%= url %>">
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
