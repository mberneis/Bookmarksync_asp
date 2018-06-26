<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<% 
Dim bid,url,rs,access,path,uid
bid = request.querystring("b")
uid = request.querystring("u")
if bid = "" or uid="" then
	conn.close
	set conn=Nothing
	response.end
end if
set rs = conn.execute ("select url from bookmarks where bookid=" & bid)
if not rs.eof then url = rs("url")
rs.close
set rs = conn.execute ("select path,access from link where expiration is null and book_id=" & bid & " and person_id=" & uid)
if not rs.eof then 
	path=rs("path")
	access=rs("access")
end if
rs.close
%>
<html>

<head>
<title>BookmarkSync - Bookmark Detail</title>
<META name="HandheldFriendly" content="true"> 
</head>
<body bgcolor=#FFFFFF>
<center><img src="smalllogo.gif" width="56" height="26"></center>
<br><b><big>Bookmark Details:</big></b><p>
Path: <%= right(path,len(path)-1) %><br>
URL: <a href="<%= url %>"><%= url %></a><br>
Created: <%= access %><br>
<%
response.write "<hr>[ <a href=add.asp>Add URL</a> | <a href=default.asp>Home</a> | <a href=default.asp?logoff=true>Log off</a> ]"
conn.close
set conn=nothing
%>
</body>
</html>
