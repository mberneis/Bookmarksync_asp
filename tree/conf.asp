<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
	Dim parg,farg,rs,argpath,mytitle,path,foldertitle,flist,addfolder
	Dim delim, publishit, level, foldernum, prefix,i,xpath,sql

response.expires=-1
restricted ("../tree/conf.asp")
ID=Session("ID")
mainmenu = "EDIT"
submenu  = ""
flist=""


addfolder = strip(request.form("addfolder"))
if addfolder<>"" then
	addfolder=trim(request.form("tree") & addfolder & "\")
	sql = "insert into link (person_id,path,access) values (" & ID & ",'" & (addfolder) & "',getdate())"
	on error resume next
	conn.execute (sql)
	'response.write sql & "<hr>"
	if err then
		on error goto 0
		sql = "update link set access = getdate(),expiration=Null where person_id=" & ID & " and path ='" & prep(addfolder) & "'"
		conn.execute (sql)
	end if
	on error goto 0
end if


parg = request.querystring("p")
if parg="" then parg = request.form("p")
farg = request.querystring("f")
if farg="" then farg=request.form("f")
if farg <> "" then 
	mytitle = "Folder" 
	path = left(farg,len(farg)-1)	
	foldertitle = "Below"	
	xpath=farg
else 
	mytitle = "Bookmark"
	path = parg
	foldertitle = "In Folder"
	xpath=parg
end if
'response.write path & "<hr>"
function makefolderlist () 
	Dim folderlist
	'folderlist = "<!-- " & argpath & " -->" & vbCrLf
	folderlist=""
	folderlist =  folderlist & "<option "
	if argpath = "\" then folderlist = folderlist &  " selected"
	folderlist = folderlist &  " value=""\"">[Root]</option>" & vbCrLf
	set rs = conn.execute ("{ call showbookmarks(" & ID & ") }")
	
	delim = ""
	publishit=true
	level = 0
	foldernum = 0
	
	while not rs.eof
		Dim p
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
			Dim f, spc, s, dbpath, xs
	
		 	i = instr(len(prefix)+1,p,delim)
			if i = 0 then exit do
			' go forward...
			f = mid(p,len(prefix)+1,i-len(prefix)-1)
			spc = ""
			for s = 1 to level
				spc = spc & "&nbsp;-&nbsp; "
			next
			dbpath = prefix & f
			dbpath = "\" & trim(right(dbpath,len(dbpath) -1)) & "\"
			if len(prefix) > 1 then
				Dim froot
	
				froot = right(prefix,len(prefix)-1)
				froot = left(froot,len(froot)-1)
			end if
			if parg<>"" or (instr(dbpath,farg)<>1 or dbpath=farg) then 
				folderlist = folderlist &  "<option "
				if trim(argpath)  = dbpath then folderlist = folderlist &  " selected "	
				if len(f) > 50 then f = left(f,48) & "..."	    
				folderlist = folderlist &  "value=""" & (dbpath) & """>" & spc & unprep(f) & "</option>"  & vbCrLf ' Directory
			end if
			flist = flist & "<option value=""" & dbpath & """>" & spc & f & "</option>"  & vbCrLf 
			level = level + 1
			prefix = left(p,i)
		Loop
		rs.movenext
	WEnd
	rs.close
	makefolderlist = folderlist
end function
%>
<!-- #include file = "../inc/tree.inc" -->
<html>

<head>
<title>BookmarkSync - Edit <%= mytitle %></title>

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>

<script language="javascript">
<!--
function delit(x) {
	if (x > 0) {
		if (confirm ('Deleting this folder also deletes all the folders and bookmarks contained within it.\nDo you want to proceed?')) {
			location.href='edit.asp?delpath=<%= server.urlencode(xpath) %>';
		} else {
			return false;
		}
	} else {
		location.href='edit.asp?delpath=<%= server.urlencode(xpath) %>';
	}
	return true;
}
//-->
</script>
<body bgcolor="#FFFFFF" text="#000000" link="#3300CC" vlink="#990000" alink="#CC0099">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="117" valign="top">
<!-- #include file = "../inc/ymenu.inc" -->
    </td>
    <td width="513" valign="top">
      <table width="513" border="0" cellspacing="0" cellpadding="0">
<!-- #include file = "../inc/yhead.inc" -->
<% if path <> "" then %>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br><%
            if addfolder <>"" then
            	response.write "<p><font color=#FF0000><b>Folder " &  addfolder & " added</b></font></p>"
            end if
            %>
            <h2>Edit <%= mytitle %></h2>
              <%
		i = len(path)
	    do
			if i = 0 then exit do
			if mid(path,i,1) = "\" then exit do
			i = i - 1
		loop
	    argpath = left(path,i)
	    path = right(path,len(path)-len(argpath))	    
      'response.write "argpath=" & argpath & "<br>path=" & path & "<hr>"
%> 
              <form method="post" action="edit.asp">
              	<input type="hidden" name="old" value="<%= argpath & path %>">
              	<input type="hidden" name="mytitle" value="<%= mytitle %>">
                <table border="0" cellpadding="2" cellspacing="0">
                  <tr>
                    <td width=80 height="19">Name</td>
                    <td height="19">
                      <input type="text" name="path" size=30 style="{width:400px;}" value="<%= unprep(path) %>">
                    </td>
                  </tr>
                  <tr>
                    <td><%= foldertitle %></td>
                    <td> 
                    <select style="{width:400px;}" name="tree" class="register">
<%= makefolderlist() %>
                    </select>
                    </td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td> 
                      <input type="button" name="btnBack" value="Back" class="pbutton" onClick="location.href='edit.asp';">
                      <input type="reset" name="btnReset" value="Reset" class="pbutton">
                      <input type="submit" name="btnSubmit" value="Modify" class="pbutton">
                      <input type="button" name="btnDelete" value="Delete" class="pbutton" onClick="return delit(<%= len(farg) %>);">
                    </td>
                  </tr>
                </table>
                <input type="hidden" name="arg" value="<%= mytitle %>">
              </form>
          </td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
        <td ><p>&nbsp;</p><hr size=2 noshade color="#990099"></td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
<% else
	makefolderlist
   end if
%>        
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br>
            <h3>Add Folder</h3>
              <form method="POST" action="conf.asp">
              	<input type="hidden" name="p" value="<%= parg %>">
              	<input type="hidden" name="f" value="<%= farg %>">
                <table border="0" cellpadding="2" cellspacing="0">
                  <tr>
                    <td height="19">Name</td>
                    <td height="19">
                      <input type="text" name="addfolder" size=30 style="{width:400px;}" value="">
                    </td>
                  </tr>
                  <tr>
                    <td width=80>Below</td>
                    <td> 
                    <select style="{width:400px;}" name="tree" class="register">
                    <option value="\" >[Root]</option>
					<%= flist %>
                    </select>
                    </td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td> 
                      <input type="button" name="btnBack" value="Back" class="pbutton" onClick="location.href='edit.asp';">
                      <input type="submit" name="btnSubmit" value="Add" class="pbutton">
                    </td>
                  </tr>
                </table>
                <input type="hidden" name="arg" value="<%= mytitle %>">
              </form>
          </td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br>
            <h3>Add Bookmark (Or use <a href="../bms/quickadd.asp">Quick Add</a>)</h3>
              <form method="POST" action="edit.asp">
              	<input type="hidden" name="p" value="<%= parg %>">
              	<input type="hidden" name="f" value="<%= farg %>">
                <table border="0" cellpadding="2" cellspacing="0">
                  <tr>
                    <td height="19">Title</td>
                    <td height="19">
                      <input type="text" name="newtitle" size=30 style="{width:400px;}" value="">
                    </td>
                  </tr>
                  <tr>
                    <td height="19">URL</td>
                    <td height="19">
                      <input type="text" name="newurl" size=30 style="{width:400px;}" value="">
                    </td>
                  </tr>
                  <tr>
                    <td width="80">In Folder</td>
                    <td> 
                    <select style="{width:400px;}" name="tree" class="register">
                    <option value="\" >[Root]</option>
					<%= flist %>
                    </select>
                    </td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td> 
                      <input type="button" name="btnBack" value="Back" class="pbutton" onClick="location.href='edit.asp';">
                      <input type="submit" name="btnSubmit" value="Add" class="pbutton">
                    </td>
                  </tr>
                </table>
                <input type="hidden" name="arg" value="<%= mytitle %>">
              </form>
          </td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>
</html>
