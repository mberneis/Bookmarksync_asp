<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
Dim deldate, realdate, sql,pid
Dim delbid, deluid, rs, p,p1,p2,undelcell,i,expiration,bg,delpid,doit 

restricted ("../bms/undelete.asp")
ID=Session("ID")
mainmenu = "UTILITIES"
submenu  = "UNDELETE"
deldate = request.form("deldate")
if deldate = "" then deldate = request.querystring("deldate")
if not IsDate(deldate)then deldate = Date - 1
realdate = dateadd("s",-1,deldate)
if request.form("undelete") <> "" then
	sql = "UPDATE Link SET Expiration = Null WHERE (link.person_id =" & ID & ") AND (convert(datetime,link.Expiration) >= convert(datetime,'" & realdate & "'))"
	connexecute (sql)
	connexecute "{ call bumptoken(" & ID & ") }"
end if
delbid = request.querystring("delbid")
deluid = request.querystring("deluid")
if delbid <> "" and deluid <> "" then
	pid = request.querystring("path")
	sql = "Update Link Set Expiration = Null where (link.person_id = " & deluid & " and link.book_id = " & delbid & " and link.path = '" &  replace(pid,"'","''") & "')"
	connexecute (sql)
	connexecute "{ call bumptoken(" & ID & ") }"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync: Undelete Bookmarks</title>
<meta http-equiv="Content-Type" content="text/html">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../common/syncit.css" type="text/css">
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
            <p>&nbsp;</p>
            <h2>Undelete Bookmarks </h2>
<%
	sql = "SELECT link.person_id, link.path, bookmarks.bookid as delbid, bookmarks.url, link.Expiration "
	sql = sql & "FROM bookmarks RIGHT JOIN link ON bookmarks.bookid = link.book_id "
	sql = sql & "WHERE ((link.person_id = " & ID & ") AND (convert(datetime,link.Expiration) >= convert(datetime, '" & realdate & "'))) "
	sql = sql & "ORDER by convert(datetime,link.Expiration) Desc"
	'response.write(sql)
	on error resume next
	set rs = conn.execute (sql)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	if rs.eof then
		if request.form("undelete") <> "" then
			response.write "<font color=""#FF0000""><b>Bookmarks undeleted. <br>Please double-click on your SyncIT Icon in the taskbar to re-synchronize</b></font>"
		else
%>
			<p><b>According to our database,
			you have not deleted any bookmarks or favorites since <%= deldate %>.</b></p>
<% 		end if %>
			<p>Questions? Please consult our handy list of Frequently Asked Questions
			(<a href="../help/faq.asp">FAQ</a>). If you do not find a matching response
			to your questions, please contact us through our
			<a href="../help/support.asp">support database</a>.</p>
            <form method="POST" action="undelete.asp">
              <span class="menub">Deleted Bookmarks since: (EST) </span><br>
              <input class="register" type="text" name="deldate" size="20" value="<%= deldate %>"> <input border="0" src="../images/btn_refresh.gif" name="I1" type="image" width="62" height="16" alt="Refresh">
            </form>
<% else %>
       <p>These Bookmarks have been deleted since <%= deldate %>.
       To restore deleted bookmarks, highlite the date you want SyncIT to go back to,
       and SyncIT will restore all bookmarks deleted since that date.
       To restore all shown deleted bookmarks, press <b>UNDELETE ALL</b>.
       To undelete individual bookmarks, click to the
       <img src="../images/gsquare.gif" border="0" alt="Instant Undelete" width="14" height="14">
       button next to each bookmark you want SyncIT to restore.</p>

			<p>Questions? Please consult our handy list of Frequently Asked Questions
			(<a href="../help/faq.asp">FAQ</a>). If you do not find a matching response
			to your questions, please contact us through our
			<a href="../help/support.asp">support database</a>.</p>

            <form method="POST" action="undelete.asp">
              <span class="menub">Deleted Bookmarks since: (EST) </span><br>
              <input class="register" type="text" name="deldate" size="20" value="<%= deldate %>"> <input border="0" src="../images/btn_refresh.gif" name="I1" type="image" width="62" height="16" alt="Refresh">
            </form>
		<form method="POST" action="undelete.asp">
		<table width="491" border="0" cellpadding="0" cellspacing="0">
			<tr bgcolor="#000066"><th align="left"><font color="#FFFFFF">&nbsp;Folder</font></th><th align="left"><font color="#FFFFFF">Bookmark</font></th><th align="right"><font color="#FFFFFF">Removed (EST)</font></th><th>&nbsp;&nbsp;</th></tr>

<%
		while not rs.eof
			p = rs("path")
			p = right(p,len(p)-1)
			p1 = ""
			p2 = p
			do
				i = instr(p2,"\")
				if i = 0 then exit do
				p2 = right(p2,len(p2)-i)
			loop
			expiration = dateadd("h",0,rs("expiration"))
			if len(p2) < len(p) then p1 = replace(left(p,len(p) - len(p2) -1),"\"," | ")
			if bg = "#EEEEEE" then bg = "#FFFFCC" else bg = "#EEEEEE"
			response.write "<tr bgcolor=" & bg & "><td>" & p1 & "</td><td>&nbsp;<a target=_blank href=""" & rs("url") & """>" & p2 & "</a></td>"
			response.write "<td align=right>"
			delbid = rs("delbid")
			deluid = rs("person_id")
			delpid = Server.URLEncode(rs("path"))
			if p2 <> "" then
				undelcell = "<td align=""right"" bgcolor=" & bg & ">&nbsp;<a href=""undelete.asp?delbid=" & delbid & "&deluid=" & deluid & "&deldate=" & Server.URLEncode(deldate) & "&path=" & delpid & """><img vspace=5 src=""../images/gsquare.gif"" border=0 alt=""Instant Undelete""></a></td>"
			else
				undelcell = "<td>&nbsp;</td>"
			end if
			rs.movenext
			if not rs.eof then
				doit = (dateadd("h",0,rs("expiration")) <> expiration )
			else
				doit = false
			end if
			if doit then
				response.write " <a href=""undelete.asp?deldate=" & server.URLEncode(expiration) & """>" &  expiration & "</a></td>"
			else
				response.write  expiration & "</td>"
			end if
			response.write undelcell & "</tr>" & vbCrLf
		wend
		%>
	      <tr><td colspan="3" align="right"><br>
          <input border="0" src="../images/btn_undeleteall.gif" name="I1" type="image" width="102" height="16" alt="Undelete All"></p>
         </td></tr>
<%	end if %>
		</table>
		<input type="hidden" name="deldate" value="<%= deldate %>">
		<input type="hidden" name="undelete" value="true">
		</form>
		</table>
          </td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
      </table>
<!-- #include file = "../inc/footer.inc" -->
</body>





