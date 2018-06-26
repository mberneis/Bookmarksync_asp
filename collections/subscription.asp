<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
Dim myurl

myurl = ("../collections/subscription.asp")
if request.querystring<>"" then myurl = myurl & "?" & request.querystring()
restricted(myurl)

Dim xaddit, rs, sql, delid, MyMsg, addit

mainmenu = "SUBSCRIBED"
submenu  = ""

ID=Session("ID")
xaddit = request.querystring("addid")
if xaddit <> "" then 'Subscribe
	if xaddit = "-1" then xaddit = ncrypt(Session("curcol"))
	addit = ndecrypt(xaddit )
	on error resume next
	set rs = conn.execute ("select user_id from publish where publishid=" & addit)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	if not rs.eof then
		if rs("user_id") = ID then
			MyMsg = "You can not subscripe to your own collection!<br><br>"
		else
			sql = "Insert into subscriptions (person_id,publish_id,Created) Values (" & ID & "," & addit & ",'" & y2k(Now()) & "')"
			'response.write (sql)
			On Error resume next
			conn.execute (sql)
			if not err then MyMsg = "The collection has been added to your subscriptions<br><br>"
			On error goto 0
		end if
	end if
	rs.close
end if

delid = request.querystring("delid")
if delid <> "" then
	connexecute "Delete from subscriptions where subscriptionid = " & ndecrypt(delid)
	MyMsg = "The collection has been removed from your subscriptions<br><br>"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync Subscribed Collections</title>
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
            You currently subscribe to the following collections. Check out our <a href="recent.asp">recent</a>
            or most <a href="popular.asp">popular</a> collections from the
            navigation menu. You can also <a href="showcategory.asp">browse</a>
            or <a href="searchcol.asp">search</a> for specific topics.
            <p>Click on a red diamond to unsubscribe from any of the
            collections.
            <br>
            <font color="#FF0000"><b><%= MyMsg %></b></font>
            </p>
            <table width="491" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="59" rowspan="2"><img src="../images/subscribeicon.gif" width="59" height="43" alt=""></td>
                <td height="14" width="374" colspan="4"><img src="../images/tpix.gif" width="374" height="14" alt=""></td>
                <td width="58" rowspan="3"><img src="../images/inviteicon.gif" width="58" height="64" alt="Invite"></td>
              </tr>
              <tr>
                <td bgcolor="#FF9933" height="29" width="374" colspan="4"> &nbsp;<b>Subscribed
                  Collections</b></td>
              </tr>
              <tr>
                <td width="59"><img src="../images/lbl_remove.gif" width="59" height="21" alt="Remove"></td>
                <td bgcolor="#000066" height="21" width="51" align="center"><img src="../images/lbl_view.gif" width="36" height="21" alt="View"></td>
                <td bgcolor="#000066" width="55" align="center"><img src="../images/lblb_author.gif" width="42" height="21" alt="Author"></td>
                <td bgcolor="#000066" width="84" align="center"><img src="../images/lblb_title.gif" width="27" height="21" alt="Title"></td>
                <td bgcolor="#000066" width="184" align="center"><img src="../images/lblb_description.gif" width="93" height="21" alt="Description"></td>
              </tr>
<%
sql = "SELECT publish.PublishID, subscriptions.SubscriptionID, publish.Title, person.name as pname, publish.Description,publish.anonymous FROM person,subscriptions,publish where publish.PublishID = subscriptions.publish_id and person.personid = publish.User_ID and publish.user_id <> " & NewsID & " and (subscriptions.person_id)=" & ID
	on error resume next
	set rs = conn.execute (sql)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
while not rs.eof
	Dim pid, sid, title, author, description

	pid = ncrypt(rs("PublishID"))
	sid = ncrypt(rs("SubscriptionID"))
	title = rs("title")
	author = rs("pname")
	description = rs("description")
	if rs("anonymous") then author="---"

	rs.movenext
	if xaddit = pid then response.write "<tr bgcolor=""#FFEEDD"">" else response.write "<tr>"
%>

                <td width="59" align="center"><a href="subscription.asp?delid=<%= sid %>"><img border="0" src="../images/rsquare.gif" width="14" height="14" alt="Remove"></a></td>
                <td width="51" align="center"><a href="../tree/cview.asp?ref=SUBSCRIBED&pid=<%= pid %>"><img src="../images/bsquare.gif" width="14" height="14" border="0" alt="View"></a></td>
                <td width="55" align="center" class="f11"><%= author %></td>
                <td width="84" align="center" class="f11" bgcolor="#FFEEDD"><%= title %></td>
                <td width="184" align="center" class="f11"><%= description %></td>
                <td width="58" align="center"><a href="invite.asp?ref=SUBSCRIBED&pid=<%= pid %>"><img src="../images/ysquare.gif" width="14" height="14" border="0" alt="Invite"></a></td>
              </tr>
<%
if not rs.eof then
	response.write "<tr><td width=""491"" align=""center"" colspan=""6""><img src=""../images/psep.gif"" width=""491"" height=""10"" alt=""""></td></tr>"
end if
wend
rs.close
%>
            </table>
            <h3><b>Legend:</b></h3>
            <p><b> </b><img src="../images/rsquare.gif" width="14" height="14" alt="Remove"> Remove
              this Collection (double-click your SyncIT icon to reflect changes)<b><br>
              </b><img src="../images/bsquare.gif" width="14" height="14" alt="View"> View this
              Collection as a bookmark tree<br>
              <img src="../images/ysquare.gif" width="14" height="14" alt="Invite"> Send someone
              an e-mail invitation to subscribe to this collection</p>
            </td>
          <td width="10" valign="top">&nbsp;</td>
      </table>

    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>
</html>









