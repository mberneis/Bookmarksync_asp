<%
pid = request.querystring("pid")
if pid="" then response.end
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>BookmarkSync Affiliate Demo</title>
</head>
<frameset cols="200,*">
<frame name="tree" src="../tree/affiliate.asp?pid=<%= request.querystring("pid") %>&target=content">
<frame name="content" src="../collections/affiliate_demo.asp?pid=<%= pid %>">
</frameset>

</html>
