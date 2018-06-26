<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<% 
Dim uid,rs,path,url,title,sql,p,dolist,license
uid = request.cookies("avantgoid")
dolist = request.cookies("avantgolist")
if uid = "" then uid = request("uid")
if dolist="" then dolist = request.form("dolist")
if uid = "" then 
	conn.close
	set conn=Nothing
	response.redirect "default.asp"
	response.end
end if

url = request.form("url")
if url <> "" and url <> "http://www.?.com" then
	title = request.form("title")
	if title="" then 
		title=url
		if left(title,7) = "http://" then title = right(title,len(title)-7)
		if left(title,4) = "www." then title = right(title,len(title)-4)
		if right(title,4) = ".com" then title = left(title,len(title)-4)
	end if
	path = request.form("folder")
	sql = "{ call addbookmark (" & uid & ",'" & replace(path & title,"'","''") & "','" & replace(url,"'","''") & "') }"
	connexecute (sql)
	connexecute ("{ call bumptoken(" & uid & ") }")

	response.cookies("msg") = request.cookies("msg") & "Added Bookmark " & title & "<br>"
	response.cookies("msg").Expires = "1/1/2022"
end if
'response.cookies("license") = license
'response.cookies("license").Expires = "1/1/2022"
'response.cookies("avantgoid") = uid
'response.cookies("avantgoid").Expires = "1/1/2022"
'response.cookies("avantgolist") = dolist
'response.cookies("avantgolist").Expires = "1/1/2022"
%>
<html>

<head>
<title>BookmarkSync - Add URL</title>
<META name="HandheldFriendly" content="true"> 
</head>
<body bgcolor=#FFFFFF>
<center><img border=0 src="smalllogo.gif" width="56" height="26"></center>
<form method="post" action="add.asp">
  <b>Add a Bookmark</b><br>
  Title:  <input type="text" name="title" size=18 maxlength=64>
    <br>
    URL: <input type="text" name="url" size=256 value="http://www.?.com">
    <br>
    Folder:    <select name="folder">
      <option value="\" selected>[Top-Level]</option> 
<%
Dim delim,level,foldernum,prefix,i
set rs = conn.execute ("{ call showbookidsX(" & uid & ") }")
delim = ""
level = 0
foldernum = 0
while not rs.eof
	p = rs("path")
	if delim = "" then
		delim = left(p,1)
		prefix = delim
	end if
	if delim <> left(p,1) then delim = left(p,1)
	while ucase(left(p,len(prefix))) <> ucase(prefix)
		level = level - 1
		do
			if len(prefix) = 0 then
				prefix = delim
				level = level + 1
			else
				prefix = left(prefix,len(prefix)-1)
			end if
			if right(prefix,1) = delim then exit do
		loop
	wend
	Do
		Dim f, spc, s, dbpath, csql, xs

	 	i = instr(len(prefix)+1,p,delim)
		if i = 0 then exit do
		' go forward...
		f = mid(p,len(prefix)+1,i-len(prefix)-1)
		spc = ""
		for s = 1 to level
			spc = spc & "&nbsp;-&nbsp; "
		next
		dbpath = prefix & f
		dbpath = right(dbpath,len(dbpath) -1)
		if len(prefix) > 1 then
			Dim froot
			froot = right(prefix,len(prefix)-1)
			froot = left(froot,len(froot)-1)
		end if
		if len(f) > 25 then f = left(f,23) & "..."	    
		response.write "<option value=""\" & dbpath & "\"">" & spc & f & "</option>"  & vbCrLf ' Directory
		level = level + 1
		prefix = left(p,i)
	Loop
	rs.movenext
WEnd
rs.close
	
%>      
    </select>
    <br>
    <input type="submit" name="Submit" value="Add" 
    onClick="document.forms[0].submitNoResponse('The Bookmark will be added during the next synchronization process',false,true)">
  <br>
  <input type=hidden name="license" value="<%= Request.Cookies ("license") %>">
  <input type=hidden name="uid" value="<%= uid %>">
  <input type=hidden name="dolist" value="<%= dolist %>">
</form>
<% 
'response.write "<hr>[ <a href=default.asp?uid=" & uid & ">Home</a> | <a href=list.asp?uid=" &  uid & ">List Bookmarks</a> | <a href=default.asp?logoff=true>Log off</a> ]"
'response.write "<a href=default.asp>Home</a>"

conn.close
set conn=nothing
%>
<center><a href="default.asp?uid=<%= uid %>"><img border=0 src="home.gif"></a></center>
</body>
</html>
