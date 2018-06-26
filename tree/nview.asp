<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
Dim menutitle, ref, pid
Dim referer

restricted ("../tree/nview.asp?" & server.urlencode(request.querystring()))
ID=Session("ID")
mainmenu= "SETUP"
submenu = "VIEW"
ref = request.querystring("ref")
if ref <> "" then menutitle = ref & "|" & menutitle
pid = ndecrypt(request.querystring("pid"))
if pid=0 then pid = Session("curcol")
if pid = 0 then myredirect "../collections/default.asp"
Session("curcol") = pid

referer = Request.ServerVariables("HTTP_REFERER")
If referer = "" Then referer = "../news/"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<!-- #include file = "../inc/tree.inc" -->
<html>
<head>
<title>View News</title>
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#660066" vlink="#990000" alink="#CC0099">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="117" valign="top">
<!-- #include file = "../inc/pmenu.inc" -->
    </td>
    <td width="513" valign="top">
      <table width="513" border="0" cellspacing="0" cellpadding="0">
<!-- #include file = "../inc/phead.inc" -->
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br>
            <%  tree request.querystring("pid"),false %>
            <br clear="all"><br>
			<a href="<%= referer %>" onclick="history.back();"><img border="0" src="../images/btnback.gif" alt="Back to previous page" width="62" height="16"></a>
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
