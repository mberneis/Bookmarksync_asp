<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
response.expires=-1
restricted ("../tree/view.asp")
ID=Session("ID")
mainmenu = "VIEW"
submenu  = ""
%>
<!-- #include file = "../inc/tree.inc" -->
<html>

<head>
<title>BookmarkSync - My Bookmarks</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#3300CC" vlink="#990000" alink="#CC0099">
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
            <h2>Welcome <%= Session("name") %> </h2>
            <h4>Your bookmarks and favorites are shown below - click on a link and go!</h4>
            <% tree 0,false %>
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
