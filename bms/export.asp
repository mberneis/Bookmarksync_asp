<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%

	ID = GetID()
	If ID = 0 Then
		myredirect "../common/login.asp?refer=../bms/export.asp"
	End If

Dim rs, md5

mainmenu = "UTILITIES"
submenu  = "EXPORT"
on error resume next
set rs = conn.execute ("SELECT personid, name, email, pass FROM person WHERE personid=" & ID)
if err then dberror err.description & " [" & err.number & "]"
on error goto 0

If rs.eof Then
	rs.close
	set rs=Nothing
	myredir "../inc/login.asp?refer=../bms/export.asp"
End If

set md5 = Server.CreateObject("FishX.MD5")
md5.Init
md5.Update(rs("email"))
md5.Update(rs("pass"))
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync - Export Bookmarks</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#660066" vlink="#990000" alink="#CC0099">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="117" valign="top">
<!-- #include file = "../inc/ymenu.inc" -->
    </td>
    <td width="513" valign="top">
      <table width="513" border="0" cellspacing="0" cellpadding="0">
<!-- #include file = "../inc/yhead.inc" -->
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br>
            <h2>Export Bookmarks</h2>
            <p><br>
<form method="POST" action="/client/export.dll?">
  <h3>
  <input type="hidden" name="email" value="<%= rs("email") %>">
  <input type="hidden" name="pass" value="<%= rs("pass") %>">Select the format
  in which to export your bookmarks</h3>
  <table border="0">
    <tr>
      <td valign="top"><b>

  <input type="radio" name="format" value="netscape" checked></b></td>
      <td><b>Netscape<br>
        </b>This commonly used format is used by the <b>Netscape</b>
        Navigator Browser.&nbsp;<br>
        in a file called <i>bookmark.htm</i>.</td>
    </tr>
    <tr>
      <td valign="top"><b>
  <input type="radio" name="format" value="opera"></b></td>
      <td><b>Opera<br>
        </b>This format is used by the popular <a href="http://www.opera.com" target="_blank"><b>Opera
        </b></a>browser.</td>
    </tr>
    <tr>
      <td valign="top"><b>
  <input type="radio" name="format" value="xbel"></b></td>
      <td><b>XBEL<br>
        </b>This XML based format was designed by the <b><a href="http://www.python.org/topics/xml/xbel/" target="_blank">Python
        </a></b>group and extended by <b>SyncIT.com</b>.</td>
    </tr>
  </table>
  <p><b>&nbsp;</b><input border="0" src="../images/btnok.gif" name="I1" type="image" width="62" height="16" alt="OK"></p>
</form>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>

</html>
<%
rs.close
set rs=Nothing
set md5=Nothing
%>
