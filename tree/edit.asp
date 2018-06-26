<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
response.expires=-1
restricted ("../tree/edit.asp")
ID=Session("ID")
mainmenu = "EDIT"
submenu  = ""
%>
<!-- #include file = "../inc/tree.inc" -->
<html>

<head>
<title>BookmarkSync - My Bookmarks</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>

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
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br>
            <h2>Welcome <%= Session("name") %> </h2>
            <h4>Your bookmarks and favorites are shown below<br>Use the green button to edit or <a href="conf.asp">add a bookmark</a> now.</h4>
            <%
			Dim old,upd,mytitle,sql,delpath,newtitle,newurl
			
			delpath = strip(request.querystring("delpath"))
			if delpath<>"" then
				if right (delpath,1) = "\" then 
					sql = "update link set expiration=getdate() where path like '" & (delpath) & "%' and person_id=" & ID
				else
					sql = "update link set expiration=getdate() where path='" & (delpath) & "' and person_id=" & ID
				end if
				'response.write (sql)
				conn.execute (sql)
				connexecute ("{ call bumptoken(" & ID & ") }")				
			end if
			old = strip(request.form("old"))
			mytitle = left(request.form("mytitle") & " ",1)
			if old <> "" then
				upd = request.form("tree") & prep(request.form("path"))
'				'response.write request.form("mytitle") & "<hr>Old = " & old & "<br>new = " & upd & "<hr>"
			end if
			if mytitle="B" and old <> upd then ' edit bookmark
				sql = "Delete from Link where path='" & upd & "' and person_id = " & ID 
				'response.write (sql)
				connexecute (sql)
				'response.write "<hr>"
				sql = "Update Link set path='" & upd & "' where person_id = " & ID & " and path='" & old & "'"
				'response.write (sql)
				connexecute (sql)
'				response.write "Bookmark edit successful"
				connexecute ("{ call bumptoken(" & ID & ") }")				
			end if
			if mytitle = "F" and old <> upd then 'Folder Edit
				sql = "Delete from Link where person_id=" & ID & " and path in (select stuff(path,1," & len(old)+1 & ",'" & (upd) & "\') from link where person_id = " & ID & " and path like '" & old & "\%')"
				'response.write "<hr>" & sql & "<hr>"
				connexecute (sql)
				sql = "Update Link set path=stuff(path,1," & len(old)+1 & ",'" & (upd) & "\') where person_id = " & ID & " and path like '" & old & "\%'"
				'response.write "<hr>" & sql & "<hr>"
				connexecute (sql)
				connexecute ("{ call bumptoken(" & ID & ") }")				
			end if
			newurl = request.form("newurl")
			if newurl <> "" then ' add bookmark
				newtitle = request.form ("newtitle")
				if newtitle = "" then
					newtitle = newurl
					if lcase(right("xxxx" & newtitle,4)) = ".com" then newtitle = left(newtitle,len(newtitle)-4)
					if left(lcase(newtitle) & "xxxxxxx",7) = "http://" then newtitle = right(newtitle,len(newtitle) -7)
				end if
				sql = "{ call addbookmark (" & ID & ",'" & replace(request.form("tree") & newtitle,"'","''") & "','" & replace(newurl,"'","''") & "') }"
				connexecute (sql)
				connexecute ("{ call bumptoken(" & ID & ") }")				
			end if
			 tree 0,true %>
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
