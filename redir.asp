<!-- #include file = "inc/asp.inc" -->
<!-- #include file = "inc/db.inc" -->
<%
Dim rs, sql
Dim u, b, p, url


u = request.querystring("u")	
b = request.querystring("b")	
p = request.querystring("p")
on error resume next	
set rs = conn.execute ("Select url from bookmarks where bookid=" & b)
if err then dberror err.description & " [" & err.number & "]"
on error goto 0
if not rs.eof then url = rs("url")
rs.close
if p = 709920 then ' log
	sql = "Insert into buttugly_redir (person_id,publish_id,book_id,access) Values (" & u & "," & p & "," & b & ",getDate())"
	conn.execute (sql)
end if
conn.close
set conn=Nothing
response.write "<head>"
response.write "<meta http-equiv=""REFRESH"" Content=""0; URL=" & url & """>"
response.write "</head>"
'response.redirect url
response.write "<a href=""" & url & """>Loading News Article...</a>"

%>
