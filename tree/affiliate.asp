<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<%
Dim invite, ref, pid
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
	pid = ndecrypt(request.querystring("pid"))
end if
if ref <> "" then mainmenu = ref
if pid=0 then pid = Session("curcol")
if pid = 0 then myRedirect ("http://www.syncit.com")
Session("curcol") = pid
Session("target") = request.querystring("target")

referer = Request.ServerVariables("HTTP_REFERER")
If referer = "" Then referer = "../collections/"
%>
<!-- #include file = "../inc/tree.inc" -->
<html>
<head>
<title>BookmarkSync - General Publication</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#3300CC" vlink="#990000" alink="#CC0099">
<% tree ncrypt(pid),false 
if request("nologo") <> "true" then
%>
<hr size="1">
<a target=_top href="http://syncit.com/tree/cview.asp?invite=<%= request.querystring("pid") %>"><b>
<img border="0" src="../images/button1.gif" align="left" hspace="10" width="88" height="33" alt="BookmarkSync Free Download!">
 Get a live subscription&nbsp;<br>
 to this Collection.<br>
</b></a>
<% end if %>
</body>

</html>
