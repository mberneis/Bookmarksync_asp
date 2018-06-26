<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<%
Dim invite, ref, pid,pubid
Dim referer

checkuser()
ID=Session("ID")
submenu = "VIEW"
invite = request.querystring("invite")
if invite <> "" then
	ref="BROWSE"
	pid = ndecrypt(invite)
else
	ref = request.querystring("ref")
	pubid = request.querystring("pid")
	pid = ndecrypt(pubid)
end if
if ref <> "" then mainmenu = ref
if pid=0 then pid = Session("curcol")
if pid = 0 then myRedirect (request.ServerVariables("HTTP_REFERER"))
Session("curcol") = pid

referer = Request.ServerVariables("HTTP_REFERER")
If referer = "" Then referer = "../collections/"
%>
<!-- #include file = "../inc/tree.inc" -->
<html>
<head>
<title>BookmarkSync - My Bookmarks</title>
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#3300CC" vlink="#990000" alink="#CC0099">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="117" valign="top">
<!-- #include file = "../inc/cmenu.inc" -->
    </td>
    <td width="513" valign="top">
      <table width="513" border="0" cellspacing="0" cellpadding="0">
<!-- #include file = "../inc/chead.inc" -->
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br>
            <% if invite<>"" then %>
            <h3><font color="#FF0000">You have been invited to <a href="../collections/subscription.asp?addid=<%= invite %>">subscribe</a> to this collection.</font></h3>
           <% 
            	tree ncrypt(pid),false 
           else 
            	tree ncrypt(pid),false 
           %>
            <h3><font color="#FF0000">Would you like to <a href="../collections/subscription.asp?addid=<%= pubid %>">subscribe</a> to this collection?</font></h3>
           <% end if
            	%>
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
