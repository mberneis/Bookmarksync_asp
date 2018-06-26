<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
restricted ("../collections/publication.asp")

Dim action, mymsg, sql, delid
Dim pid, description, category, title, pcount
Dim rs

ID=Session("ID")
mainmenu = "PUBLISHED"
submenu  = ""
action = request.form("action")
if action <> ""  then ' Add the new publication

	description = nohack(strip(request.form("description")))
	category = request.form("category")
	title = nohack(strip(request.form("title")))
	'on error resume next
	if action = "add" then
		Dim tree, doadd, newpid

		tree = strip(request.form("tree"))
		on error resume next
		set rs = conn.execute ("select User_ID from publish where Path  ='" & tree & "' and User_ID = " & ID)
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		doadd = rs.eof
		rs.close
		if doadd then
			sql = "Insert into publish (User_ID,Path,Category_ID,Created,token,Anonymous,Title,Description) Values ("
			sql = sql & ID & ", "
			sql = sql & "'" & tree & "',"
			sql = sql & category & ", getdate(), 0, "
			if request.form("anonymous") = "" then sql = sql & "0," else sql = sql & "1,"
			sql = sql & "'" & title & "',"
			sql = sql & "'" & description & "')"
			connexecute (sql)
			on error resume next
			set rs = conn.execute  ("Select publishid from publish where path = '" & tree & "' and user_id = " & ID)
			if err then dberror err.description & " [" & err.number & "]"
			on error goto 0
			newpid = rs("publishid")
			rs.close
			myredirect ("invite.asp?pid=" & ncrypt(newpid) & "&addbm=true")
		else
			MyMsg = "The collection has already been added previously!<br><br>"
		end if
	end if
	if action = "mod" then
		sql =  "UPDATE publish SET category_id = " & category & ", title = '" & title & "', description = '" & description & "', anonymous="
		if request.form("anonymous") = "" then sql = sql & "0" else sql = sql & "1"
		sql = sql & ", token = token + 1" & " WHERE publishid = " & ndecrypt(request.form("pid")) 
		connexecute (sql)				

		MyMsg = "Your collection has been modified<br><br>"
	end if
	on error goto 0
end if

delid = request.querystring("delid")
if delid <> "" then
	delid = ndecrypt(delid)
	on error resume next
	set rs = conn.execute ("select user_id from publish where publishid=" & delid)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	
	if not rs.eof then
		if ID <> rs("user_id") then
			MyMsg = "You can only remove your own collections!<br><br>"
		else
			connexecute "Delete from subscriptions where publish_id=" & delid
			connexecute "Delete from publish where PublishID = " & delid & " and User_ID = " & ID
			MyMsg = "Your collection has been removed!<br><br>"
		end if
	end if
	rs.close
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync Published Collections</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#3333FF" vlink="#990099" alink="#CC0099">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="111" valign="top">
<!-- #include file = "../inc/cmenu.inc" -->
    </td>
    <td width="511" valign="top">
      <table width="511" border="0" cellspacing="0" cellpadding="0">
<!-- #include file = "../inc/chead.inc" -->
          <td width="10">&nbsp;</td>
          <td width="491" valign="top">
            <br>
            Here are your published collections.<br>
            Click on the green diamond to modify or remove any of them<br><b><font color="#FF0000"><%= MyMsg %></font>
            </b>
            <table width="491" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="59" rowspan="2"><img src="../images/subscribeicon.gif" width="59" height="43" alt=""></td>
                <td height="14" width="374" colspan="4"><img src="../images/tpix.gif" width="374" height="14" alt=""></td>
                <td width="58" rowspan="3"><img src="../images/inviteicon.gif" width="58" height="64" alt="Invite"></td>
              </tr>
              <tr>
                <td bgcolor="#FF9933" height="29" width="374" colspan="4"> &nbsp;<b>My
                  Published Collections</b></td>
              </tr>
              <tr>
                <td width="59"><img src="../images/lbl_fans.gif" width="59" height="21" alt="Fans"></td>
                <td bgcolor="#000066" height="21" width="51" align="center"><img src="../images/lbl_view.gif" width="36" height="21" alt="View"></td>
                <td bgcolor="#000066" width="55" align="center"><img src="../images/lbl_modify.gif" width="42" height="21" alt="Modify"></td>
                <td bgcolor="#000066" width="84" align="center"><img src="../images/lblb_title.gif" width="27" height="21" alt="Title"></td>
                <td bgcolor="#000066" width="184" align="center"><img src="../images/lblb_description.gif" width="93" height="21" alt="Description"></td>
              </tr>
<%
sql="SELECT publishid, title, (SELECT COUNT(*) FROM subscriptions WHERE publish_id=publishid) AS pcount, publish.description FROM publish WHERE publish.User_ID=" & ID
on error resume next
set rs = conn.execute (sql)
if err then dberror err.description & " [" & err.number & "]"
on error goto 0
while not rs.eof
	pid = ncrypt(rs("PublishID"))
	title = rs("title")
	pcount = rs("pcount") & ""
	description = rs("description")
	rs.movenext
	if pcount = "" then pcount = "0"
	response.write "<tr>" & vbCrLf
	response.write "<td width=""59"" align=""center"">"
	response.write pcount
	response.write "</td>" & vbCrLf
	response.write "<td width=""51"" align=""center""><a href=""../tree/cview.asp?ref=PUBLISHED&pid="
	response.write pid
	response.write """><img src=""../images/bsquare.gif"" width=""14"" height=""14"" border=""0"" alt=""View""></a></td>" & vbCrLf
	response.write "<td width=""55"" align=""center"" class=""f11""><a href=""editpub.asp?ref=PUBLISHED&pid="
	response.write pid
	response.write """><img src=""../images/gsquare.gif"" width=""14"" height=""14"" border=""0"" alt=""Modify""></a></td>" & vbCrLf
	response.write "<td width=""84"" align=""center"" class=""f11"" bgcolor=""#FFEEDD"">"
	response.write title
	response.write "</td>" & vbCrLf
	response.write "<td width=""184"" align=""center"" class=""f11"">"
	response.write description
	response.write "</td>" & vbCrLf
 	response.write "<td width=""58"" align=""center""><a href=""invite.asp?ref=PUBLISHED&pid="
	response.write pid
	response.write """><img src=""../images/ysquare.gif"" width=""14"" height=""14"" border=""0"" alt=""Invite""></a></td></tr>" & vbCrLf
	if not rs.eof then
		response.write "<tr><td width=""491"" align=""center"" colspan=""6""><img src=""../images/psep.gif"" width=""491"" height=""10"" alt=""""></td></tr>"
	end if
wend
rs.close
%>
            </table>
            <h3><b>Legend:</b></h3>
            <p><b> </b><img src="../images/bsquare.gif" width="14" height="14" alt="View"> View
              this Collection as a bookmark tree<br>
              <img src="../images/gsquare.gif" width="14" height="14" border="0" alt="Change">
              Make changes to your collection (modify title, description, access-rights)<br>
              <img src="../images/ysquare.gif" width="14" height="14" alt="Invite"> Send someone
              an e-mail invitation to subscribe to this collection
            </p>
            </td>
          <td width="10" valign="top">&nbsp;</td>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>

</html>
