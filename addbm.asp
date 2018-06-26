<!-- #include file = "inc/asp.inc" -->
<!-- #include file = "inc/db.inc" -->
<%
	Dim AID,path,url,title,email,pass,rs,sql
	
	AID = request.form("AID")
	path = request.form("path")
	url = request.form("url")
	if url = "" then url = request.querystring("url")
	title = request.form("title")
	if title = "" then title = request.querystring("title")
	if title = "" then title = url
	email = strip(request.form("email"))
	pass = strip(request.form("pass"))
	if AID = "" then AID = Session("AID")
	if AID = "" and email <> "" and pass <> "" then
		on error resume next
		set rs = conn.execute ("Select personid from person where email = '" & email & "' and pass = '" & pass & "'")
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		if not rs.eof then
			AID = rs("personid")
			Session("AID") = AID
		end if
		rs.close
	end if
	

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<title>Add Page to BookmarkSync</title>
<link rel="StyleSheet" href="common/syncit.css" type="text/css">
</head>

<body text="#000000" link="#000066" vlink="#800000" alink="#FFFFFF" bgcolor=#FFFFFF>
<img border="0" src="images/logo1.gif" width="113" height="52">
<%
if AID = "" then ' log in
%>

<p><b>Log In</b></p>
<form method="POST" action="addbm.asp">
<input type="hidden" name="title" value="<%= title %>">
<input type="hidden" name="url" value="<%= url %>">

  <table border="0">
    <tr>
      <td align="right"><font face="Arial,Helvetica,Sans Serif" size="2">E-Mail:</font></td>
      <td><input type="text" name="email" size="20" style="width:300px;"></td>
    </tr>
    <tr>
      <td align="right"><font face="Arial,Helvetica,Sans Serif" size="2">Password:</font></td>
      <td><input type="password" name="pass" size="20" style="width:300px;"></td>
    </tr>
    <tr>
      <td align="right"></td>
      <td><input type="submit" value="Log in" name="B1" class="pbutton"></td>
    </tr>
  </table>
</form>
<% else 
if path  <> "" then 
	if path <> "\" then path = "\" & path & "\"
	sql = "{ call addbookmark (" & AID & ",'" & replace(path & title,"'","''") & "','" & replace(url,"'","''") & "') }"
	'response.write (sql)
	connexecute (sql)
	connexecute ("{ call bumptoken(" & AID & ") }")
%>
<p><b>The bookmark has been added.<br></b>
You can <a href="javascript:top.close();">close</a> the window now.</p>

<% else %>
<form method="POST" action="addbm.asp">
<input type="hidden" name="AID" value="<%= AID %>">
  <table border="0">
    <tr>
      <td align="right"><font face="Arial,Helvetica,Sans Serif" size="2">Title:</font></td>
      <td><input type="text" name="title" size="40" value="<%= title %>"></td>
    </tr>
    <tr>
      <td align="right"><font face="Arial,Helvetica,Sans Serif" size="2">URL:</font></td>
      <td><input type="text" name="url" size="40" value="<%= url %>"></td>
    </tr>
    <tr>
      <td align="right"><font face="Arial,Helvetica,Sans Serif" size="2">Folder:</font></td>
      <td><select size="1" name="path">
      <option value="\">TOPLEVEL</option>
      <%
      Dim p,delim,level,prefix,i,f,dbpath,xs,spc,s,froot
      
      ID = AID
      sql = "{ call showbookmarks(" & ID & ") }"
      on error resume next
		set rs = conn.execute (sql)
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		
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
			froot = right(prefix,len(prefix)-1)
			froot = left(froot,len(froot)-1)
		end if
		response.write "<option value=""" & dbpath & """>" & spc & f & "</option>" &vbCrLf ' Directory
		level = level + 1
		prefix = left(p,i)
	Loop
	rs.movenext			
		wend
      %>
        </select>&nbsp;</td>
    </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td>&nbsp;<input type="submit" value="Add Bookmark" name="B1" class="pbutton"></td>
    </tr>
  </table>
</form>
<% end if 
end if
on error resume next
conn.close
set conn=Nothing
%>
</body>

</html>
