<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
Dim  rs

mainmenu = "RECENT"
submenu  = ""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync Collections</title>
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
            <h2><br>
            Recent Collections</h2>
<%
			Dim token, sql

			sql = "SELECT top 20 publishid, person.name, publish.title, publish.User_id, publish.description,publish.anonymous FROM publish, person WHERE publish.Category_ID >0 AND publish.User_ID = person.personid order by publishid desc"
			on error resume next
			set rs = conn.execute (sql)
			if err then dberror err.description & " [" & err.number & "]"
			on error goto 0
			
			if rs.eof then
				response.write "No collections found"
			else
%>
            <table width="491" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="59" rowspan="2"><img src="../images/subscribeicon.gif" width="59" height="43" alt=""></td>
                <td height="14" width="374" colspan="4"><img src="../images/tpix.gif" width="374" height="14" alt=""></td>
                <td width="58" rowspan="3"><img src="../images/inviteicon.gif" width="58" height="64" alt="Invite"></td>
              </tr>
              <tr>
                <td bgcolor="#FF9933" height="29" width="374" colspan="4">
                <b>Most recent collections: </b>&nbsp;
                  </td>
              </tr>
              <tr>
                <td width="59"><img src="../images/lblb_subscribe.gif" width="59" height="21" alt="Subscribe"></td>
                <td bgcolor="#000066" height="21" width="51" align="center"><img src="../images/lbl_view.gif" width="36" height="21" alt="View"></td>
                <td bgcolor="#000066" width="55" align="center"><img src="../images/lblb_author.gif" width="42" height="21" alt="Author"></td>
                <td bgcolor="#000066" width="84" align="center"><img src="../images/lblb_title.gif" width="27" height="21" alt="Title"></td>
                <td bgcolor="#000066" width="184" align="center"><img src="../images/lblb_description.gif" width="93" height="21" alt="Description"></td>
              </tr>
<%
while not rs.eof
	Dim pid, uid, author, title, description

	pid = ncrypt(rs("PublishID"))
	uid = ncrypt(rs("User_ID"))
	author = rs("name")
	title = rs("title")
	description = rs("description")
	if rs("anonymous") then author="---"
%>
              <tr>
                <td width="59" align="center"><a href="subscription.asp?addid=<%= pid %>"><img src="../images/osquare.gif" width="13" height="13" border="0" alt="Subscribe"></a></td>
                <td width="51" align="center"><a href="../tree/cview.asp?ref=RECENT&amp;pid=<%= pid %>"><img src="../images/bsquare.gif" width="14" height="14" border="0" alt="View"></a></td>
                <td width="55" align="center" class="f11"><%= author %></td>
                <td width="84" align="center" class="f11" bgcolor="#FFEEDD"><%= title %></td>
                <td width="184" align="center" class="f11"><%= description %></td>
                <td width="58" align="center"><a href="invite.asp?ref=RECENT&amp;pid=<%= pid %>"><img src="../images/ysquare.gif" width="14" height="14" border="0" alt="Invite"></a></td>
              </tr>
<%
	rs.movenext
	if not rs.eof then
		response.write "<tr><td width=""491"" align=""center"" colspan=""6""><img src=""../images/psep.gif"" width=""491"" height=""10"" alt=""""></td></tr>"
	end if
wend
rs.close
%>
            </table>
            <h3><b>Legend:</b></h3>
            <p><b> </b> <a href="subscription.asp"><img src="../images/osquare.gif" width="13" height="13" border="0" alt="Subscribe"></a>
              Add this Collection (double-click your SyncIT icon to reflect changes)<b><br>
              </b><img src="../images/bsquare.gif" width="14" height="14" alt="View"> View this
              Collection as a bookmark tree<br>
              <img src="../images/ysquare.gif" width="14" height="14" alt="Invite"> Send someone
              an e-mail invitation to subscribe to this collection</p>
            <p>&nbsp; </p>
<% 
end if %>
            </td>
          <td width="10" valign="top">&nbsp;</td>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>

</html>
