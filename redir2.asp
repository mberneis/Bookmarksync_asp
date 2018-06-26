<%
option explicit

Dim conn, rs, sql
Dim u, b, p, url

set conn = Server.CreateObject("ADODB.Connection")
conn.open "bm", "websync", ""

u = request.querystring("u")	
b = request.querystring("b")	
p = request.querystring("p")
	
set rs = conn.execute ("Select url from bookmarks where bookid=" & b)
if not rs.eof then url = rs("url")
rs.close
if p = 709920 then ' log
	sql = "Insert into buttugly_redir (person_id,publish_id,book_id,access) Values (" & u & "," & p & "," & b & ",getDate())"
	conn.execute (sql)
end if
conn.close
set conn=Nothing
if p = 709920 then 
	response.write "<script>top.location.href='" + url + "';</script>"
	response.end
end if
response.redirect url
%>
