<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
Dim bmcnt,lastchanged,name,msg,aid,license
dim email,pass,uid,rs,sql,dolist,errmsg

if request.querystring("logoff") <> "" then
	response.cookies("popup")=""
	response.cookies("avantgoid")=""
	response.cookies("avantgoid").Expires=Date-10
	uid=""
end if


email = request.form("email")
pass = request.form("pass")


if email <> "" and pass <> "" then
	sql = "select personid,name from person where email = '" & email & "' and pass='" & pass & "'"
	'response.write sql
	set rs = conn.execute (sql)
	if rs.eof then
		errmsg = "<hr>Invalid email/password combination<hr>"
		response.cookies("avantgoid") = ""
	else
		uid = rs("personid")
		response.cookies("avantgoid") =uid
		response.cookies("avantgoid").Expires = "1/1/2022"
		response.cookies("email") =email
		response.cookies("email").Expires = "1/1/2022"
		license = 1
	end if
	rs.close
	dolist = (request.form("dolist") <> "")
	if dolist then response.cookies("avantgolist") = "dolist" else response.cookies("avantgolist") =""	
	response.cookies("avantgolist").Expires = "1/1/2022"
end if
msg = request.cookies("msg")
response.cookies("msg") = ""
aid = request.cookies("initid")
if aid <> "" then
	response.cookies("initid") = ""
	uid = ndecrypt(aid)
	response.cookies("avantgoid") = uid
end if
if uid="" then uid=request.querystring("uid")
if uid <> "" then 
	response.cookies("avantgoid") = uid
	response.cookies("avantgoid").Expires = "1/1/2022"
end if
if len(uid)=0 then uid = request.cookies("avantgoid")
dolist = (request.cookies("avantgolist") <> "")
if CLng("0" & uid) > 0 then
	set rs = conn.execute ("SELECT token FROM person WHERE personid=" & uid)
End if

%>
<html>

<head>
<title>BookmarkSync - My Bookmarks</title>
<META name="HandheldFriendly" content="true"> 
</head>
<body bgcolor=#FFFFFF>
&nbsp;<br><center><img src="smalllogo.gif" width="56" height="26"></center>
<%			
'response.write "AvantGoid,License,,uid,aid: " & request.cookies("avantgoid") & request.cookies("license") & "|" & len(uid) & "|" & aid & "|<br>"
if uid = "" then

if errmsg<>"" then 
	response.write errmsg
	uid=""
	response.cookies("avantgoid") = ""
end if
%> 
<form method="post" action="default.asp">
  <b>Please enter your account info</b><br>
  E-mail: <input type="text" name="email" size=18 maxlength=64 value="<%= request.cookies("email") %>">
    <br>
    Password: <input type="password" name="pass" size=18 maxlength=64>
    <br>
    Download Bookmarks? <input type="checkbox" name="dolist" value="listall"> [<a href="help.htm">Explain</a>]<br>
    <input type="submit" name="Submit" value="Log in" 
    onClick="document.forms[0].submitNoResponse('Please synchronize now for full access to BookmarkSync',true,false)">  
</form>
<% 
	response.cookies("popup") = "done"

else 
  	if msg <> "" then response.write "<hr>" & msg 
  	response.write "<hr>"
  	on error resume next
  	set rs = conn.execute ("{ call countbookmarks(" & uid & ") }")
  	if err then dberror err.description & " [" & err.number & "]"
  	bmcnt = rs("cl")
  	rs.close
  	set rs = conn.execute ("Select name,lastchanged from person where personid = " & uid)
  	if err then dberror err.description & " [" & err.number & "]"
  	lastchanged = rs("lastchanged")
  	name = rs("name")
  	rs.close
	on error goto 0
	response.write "<b>Hi, " & name & "</b><br>You have " & bmcnt & " Bookmarks stored.<br>Last update: " & lastchanged & "<br>"

	'response.write "<hr>[ <a href=add.asp>Add URL</a> | <a href=list.asp>List Bookmarks</a> | <a href=default.asp?logoff=true>Log off</a>]"
	response.write "<hr><center>"
	response.write "<a href=""add.asp?uid=" & uid & """><img border=0 src=""add.gif""></a> &nbsp;&nbsp;"
	if dolist then response.write "<a href=""list.asp?uid=" & uid & """><img border=0 src=""list.gif""></a> &nbsp;"
	response.write "&nbsp; <a href=""default.asp?logoff=true""><img src=""logoff.gif"" border=0></a></center>"
 end if  
conn.close
set conn=nothing
%>
</body>
</html>
