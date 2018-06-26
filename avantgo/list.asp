<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<% 
Dim uid,rs,path,url,title,sql,license
uid = request.cookies("avantgoid")
if uid=""  then uid=request("uid")
if uid="" then 
	conn.close
	set conn=Nothing
	response.redirect "default.asp"
	'response.write "<a href=default.asp>AvantGo cookie not set</a>"
	response.end
end if
'response.cookies("avantgoid") = uid
'response.cookies("avantgoid").Expires = "1/1/2022"
'response.cookies("license") = license
'response.cookies("license").Expires = "1/1/2022"
%>
<html>

<head>
<title>BookmarkSync - List URLs</title>
<META name="HandheldFriendly" content="true"> 
</head>
<body bgcolor=#FFFFFF>
<center><img src="smalllogo.gif" width="56" height="26"></center>
<%
Dim spath,p,i,bmstr,olddir,dir,u
spath = request.querystring("p")
if spath = "" then 	spath = "\" 
if spath <> "\" then
	bmstr=left(spath,len(spath)-1)
	while len(bmstr) > 0 and right(bmstr,1) <> "\"
		bmstr=left(bmstr,len(bmstr)-1)
	wend
	response.write "<big><b><a href=""list.asp?p=" & bmstr & """>" & spath & "</a></b></big><hr>"
end if
bmstr=""
'sql = "{ call showbookidsX(" & uid & ") }"
sql = "{ call showbookmarks(" & uid & ") }"
on error resume next
set rs = conn.execute (sql)
if err then
	response.write "Database retrieval problem - Please try again later."
else
while not rs.eof
	p = rs("path")
	
	if lcase(left(p,len(spath))) = lcase(spath) then
		p = right(p,len(p)-len(spath))
		i = instr(p,"\")
		if i > 0 then
			dir = left(p,i-1)
			if lcase(dir) <> lcase(olddir) then
				olddir = dir
				response.write "<b>[<a href=""list.asp?uid=" & uid & "&p=" & Server.URLEncode(spath & dir) & "\"">" & dir & "</a>]</b><br>"
			end if
		end if
		i = instr(p,"\")
		u = rs("url")
		'if i = 0 then bmstr=bmstr & "<nobr><a href=# onCLick=""alert('" & u & "');"">" &  p & "</a><br>"
		'if i = 0 then bmstr=bmstr & "<nobr><a href=""disp.asp?u=" & uid & "&b=" & rs("book_id") & """>" &  p & "</a><br>"
		'if i=0 and u<>"" then bmstr=bmstr & "<nobr>" &  p & " [<a href=""" & u & """>" & u & "</a>]<br>"
		if i=0 and u<>"" then bmstr=bmstr & "<nobr><input type=""image"" src=""a.gif"" onClick=""alert('" & replace(replace(u,"'","\'"),"""","\'") & "')"">" &  p & "<br>"
		'if i=0 and u<>"" then bmstr=bmstr & "<nobr><input type=""image"" src=""a.gif"" onClick=""document.forms[0].submitNoResponse('" & replace(replace(u,"'","\'"),"""","\'") & "',true,false)"">" &  p & "<br>"
	end if
	rs.movenext
wend
end if
rs.close
response.write bmstr

	response.write "<hr><center>"
	response.write "<a href=""add.asp?uid=" & uid & """><img border=0 src=""add.gif""></a> &nbsp;&nbsp;<a href=""default.asp?uid=" & uid & """><img border=0 src=""home.gif""></a> &nbsp;&nbsp; <a href=""default.asp?logoff=true""><img src=""logoff.gif"" border=0></a></center>"

conn.close
set conn=nothing
%>
</body>
</html>
