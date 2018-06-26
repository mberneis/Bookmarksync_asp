<!-- #include file = "../inc/asp.inc" -->
<% response.expires = -1 %>
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<!-- #include file = "../inc/welcome.inc" -->
<%
restricted ("../common/profile.asp")

Dim rs, name, address1, address2, city, state, zip, phone
Dim fax, email, description, wherefrom, pass, country
Dim age, gender, optin, mymsg, errmsg, sql
Dim pc_license,mac_license,linux_license

mainmenu = "PROFILE"
submenu  = ""

ID = Session("ID")
if request.form("email") <> "" then
	pass = strip(request.form("pass"))
	if pass <> "" then
		if pass <> strip(request.form("pass2")) then
			errmsg = "Passwords do not match."
		end if
	end if
	if errmsg = "" then
		email = strip(request.form("email"))
		name = strip(request.form("name"))
		if name = "" or email = "" then
			errmsg = "Name and E-mail are required fields"
		end if
	end if
	if errmsg = "" then
		if email <> Session("email") then
			set rs = conn.execute ("select  personid from person where email='" & email & "'")
			if not rs.eof then 
				errmsg = "New e-mail address &lt;<i>" & email & "</i>&gt; exists already!"
				email = Session("email")
			end if
			rs.close
		end if
	end if
	if errmsg = "" then
		address1 	= strip(request.form("address1"))
		address2 	= strip(request.form("address2"))
		city 		= strip(request.form("city"))
		state 		= strip(request.form("state"))
		zip 		= strip(request.form("zip"))
		country 	= strip(request.form("country"))
		phone 		= strip(request.form("phone"))
		fax		= strip(request.form("fax"))
		gender 		= request.form("gender")
		age 		= request.form("age")
		description	= strip(request.form("description"))
		sql = "UPDATE person SET "
		if pass <> "" then
			sql = sql & "pass='" & pass & "',"
		end if
		if email <> Session("email") then
			sql = sql & "lastverified=NULL,"
		end if
		if request.form("optin") ="" then optin = 0 else optin = 1
		sql = sql & "optin=" & optin & ","
		sql = sql & "email='" & email & "',"
		sql = sql & "name='" & name & "',"
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
		sql = sql & "description='" & description& "' "
		sql = sql & "where personid = " & ID
		connexecute (sql)
		Session("name") = name
		if email <> Session("email") then
			Response.Cookies("email") = email
			Response.Cookies("email").Expires = Date + 365
			Session("email") = email
			connexecute ("Update person set lastverified = Null where personid=" & ID)
			domail "support@bookmarksync.com",email,"Welcome to BookmarkSync",welcometxt (name,ncrypt(ID))
		end if
		mymsg = "is modified"
	end if
end if
on error resume next
set rs = conn.execute ("Select * from person where personid = " & ID)
if err then dberror err.description & " [" & err.number & "]"
on error goto 0
name = rs("name")
address1 = rs("address1")
address2 = rs("address2")
city = rs("city")
state = rs("state")
zip = rs("zip")
phone = rs("phone")
fax = rs("fax")
email = rs("email")
description = rs("description")
wherefrom = rs("wherefrom")
pass=rs("pass")
country = rs("country")
age= rs("age")
gender = rs("gender")
optin = rs("optin")
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>SyncIt Profile</title>
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
          <td colspan="2" width="503" height="37" bgcolor="#000066"><img src="../images/heading_profile.gif" width="503" height="37" alt="SyncIT Profile"></td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br><h2>Your Profile <font color="#FF0000"><%= mymsg %></font></h2>
            <p><font color="#FF0000"><b><%= errmsg %></b></font></p>
<script language="JavaScript"><!--
function validate(theForm)
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

  if (theForm.zip.value.length > 10)
  {
    alert("Please enter at most 10 characters in the 'zip' field.");
    theForm.zip.focus();
    return (false);
  }
  return (true);
}
//--></script><form method="POST" action="profile.asp" onsubmit="return validate(this)">
  <table border="0" width="100%">
    <tr>
      <td class="menub" align="right">Name</td>
      <td><input type="text" name="name" size="33" class="register" maxlength="50" value="<%= name %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">E-mail</td>
      <td><input type="text" name="email" size="33" class="register" maxlength="50" value="<%= email %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">Address</td>
      <td><input type="text" name="address1" size="33" class="register" maxlength="50" value="<%= address1 %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">Address</td>
      <td><input type="text" name="address2" size="33" class="register" maxlength="50" value="<%= address2 %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">City</td>
      <td><input type="text" name="city" size="33" class="register" maxlength="50" value="<%= city %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">State</td>
      <td><input type="text" name="state" size="33" class="register" maxlength="20" value="<%= state %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">Postal Code</td>
      <td><input type="text" name="zip" size="33" class="register" maxlength="10" value="<%= zip %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">Country</td>
      <td><input type="text" name="country" size="33" class="register" maxlength="30" value="<%= country %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">Phone</td>
      <td><input type="text" name="phone" size="33" class="register" maxlength="50" value="<%= phone %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">Fax</td>
      <td><input type="text" name="fax" size="33" class="register" maxlength="50" value="<%= fax %>"></td>
    </tr>
    <tr>
      <td class="menub" align="right">Gender</td>
      <td><select size="1" name="gender" class="register">
		<option value="U" <% if gender = "U" then response.write "selected" %>>Not specified</option>
       <option value="F" <% if gender = "F" then response.write "selected" %>>Female</option>
       <option value="M" <% if gender = "M" then response.write "selected" %>>Male</option>
	</select>
    </tr>
    <tr>
      <td class="menub" align="right">Age Group</td>
      <td><select size="1" name="age" class="register">
		<option value="10" <% if age = "10" then response.write "selected" %>>13-18</option>
       <option value="20" <% if age = "20" then response.write "selected" %>>19-25</option>
       <option value="30" <% if age = "30" then response.write "selected" %>>26-35</option>
       <option value="40" <% if age = "40" then response.write "selected" %>>36-49</option>
       <option value="50" <% if age = "50" then response.write "selected" %>>50-59</option>
       <option value="60" <% if age = "60" then response.write "selected" %>>60-69</option>
       <option value="70" <% if age = "70" then response.write "selected" %>>70 and over</option>
        </select></td>
    </tr>
			 <tr>
                    <td colspan=2 class="menub"> 
                      <div align="center"><font color="#CC0000">IF YOU ARE UNDER 
                        13 PLEASE LET A PARENT OR GUARDIAN SIGN IN!</font></div>
                    </td>
                  </tr>
              <tr>
    <tr>
      <td class="menub" align="right">About You</td>
      <td><textarea rows="4" name="description" cols="32" class="register"><%= description %></textarea></td>
    </tr>
    <tr>
      <td class="menub" align="right">New Password</td>
      <td><input type="password" name="pass" size="33" class="register" maxlength="50"></td>
    </tr>
    <tr>
      <td class="menub" align="right">Confirm<br>
        Password</td>
      <td><input type="password" name="pass2" size="33" class="register" maxlength="50"></td>
    </tr>
  </table>
  <p><input border="0" src="../images/btnmodify.gif" name="modify" value="modify" type="image" width="62" height="16" alt="Modify"></p>
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
