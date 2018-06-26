<%
	Dim bid,rs,conn,url
	bid = request.querystring("bid")
	if bid = "" then response.end
	on error resume next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "bm","websync",""
	set rs = conn.execute ("select url from bookmarks where bookid=" & bid)
	url = rs("url")
	rs.close
	conn.close
	response.redirect url
%>