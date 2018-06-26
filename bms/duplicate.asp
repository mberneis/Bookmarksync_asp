<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
restricted ("../bms/duplicate.asp")

Dim sql, rs, token, upsql,strtemp,strtoken
Dim cnt, bid, firststr, delstr, lastb, bg, deldup

ID=Session("ID")
mainmenu = "UTILITIES"
submenu  = "DUPLICATES"
sql = "{ call showbookids(" & ID & ") }"
on error resume next
set rs = conn.execute ("SELECT token FROM person WHERE personid = " & ID)
if err then dberror err.description & " [" & err.number & "]"
on error goto 0
token = rs("token")
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync: Duplicate Links</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#660066" vlink="#990000" alink="#CC0099">
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
            <h2>Duplicate Links</h2>
            <form method="POST" action="duplicate.asp">
<%
	if request.form("token") <> "" then
		if CLng(token) = CLng(request.form("token")) then
			' delete duplicates
			deldup= " " & request.form("deldup") & ","
			on error resume next
			set rs = conn.execute (sql)
			if err then dberror err.description & " [" & err.number & "]"
			on error goto 0
			cnt = 0
			strtemp = ""
			while not rs.eof
				strtemp = strtemp & replace(rs("path"),"'","''") & chr(13)
				rs.movenext
			wend
			rs.close
			do			
				cnt = cnt + 1
				i = instr(strtemp,chr(13))
				if i = 0 then exit do
				strtoken = left(strtemp,i-1)
				strtemp = right(strtemp,len(strtemp)-i)
			
				if instr(deldup," " & cnt & ",") then
					'upsql = "UPDATE link SET expiration = getdate() WHERE person_id = " & ID & " and path = '" & replace(rs("path"),"'","''") & "'"
					upsql = "{call delbookmark(" & id & ",'" & strtoken & "')}"
					'response.write upsql
					connexecute upsql
				end if
			loop
			if cnt > 0 then
				connexecute ("{ call bumptoken(" & ID & ") }")
				response.write "<p><b>Selected duplicates removed. <br>Please double-click on the SyncIT icon in your taskbar to re-synchronize</b></p>"
			end if
		else
			response.write "<p><b>Your bookmarks have changed since your selection.<br>"
			response.write "Please reselect the duplicates to remove again.</b></p>"
		end if
	end if
	response.write "<input type=hidden name=""token"" value=""" & token & """>"
	on error resume next
	set rs = conn.execute (sql)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	if rs.eof then
		response.write "<p><b>You have no bookmarks</b></p>"
	else
%>		
	    These are all the duplicates we could find in your bookmarks.<br>
		You may want to remove some duplicate entries by checking the appropriate link.   <br><br>
		<table width="491" border="0" cellpadding="0" cellspacing = "0">
			<tr bgcolor="#000066"><th align="left"><font color="#FFFFFF">&nbsp;Folder</font></th><th align="left"><font color="#FFFFFF">Bookmark</font></th><th align="right"><font color="#FFFFFF">Delete ?&nbsp;</font></th></tr>
<%
		cnt = 0
		bid = 0
		while not rs.eof
			Dim p, p1, p2, chk, b

			cnt = cnt + 1
			p = rs("path")
			p = right(p,len(p)-1)
			p1 = ""
			p2 = p
			do
				dim i
				i = instr(p2,"\")
				if i = 0 then exit do
				p2 = right(p2,len(p2)-i)
			loop
			if len(p2) < len(p) then p1 = replace(left(p,len(p) - len(p2) -1),"\","|")
			if bid <> rs("book_id") then
				bid = rs("book_id")
				chk = ""
			else
				chk = "checked"
			end if
			b = rs("book_id")
			delstr = "<td>" & p1 & "</td><td><a target=_blank href=""showbm.asp?bid=" & b & """>" & p2 & "</a></td>"
			delstr = delstr &  "<td align=right><input name=deldup value = " & cnt & " type=checkbox " & chk &"></td></tr>"
			if b <> lastb then
				lastb = b
				firststr = delstr
			else
				if firststr <> "" then
					if bg = "#EEEEEE" then bg = "#FFFFCC" else bg = "#EEEEEE"
					response.write "<tr bgcolor=" & bg & ">" & firststr
					firststr = ""
				end if
				response.write "<tr bgcolor=" & bg & ">" & delstr
			end if
			rs.movenext
		wend
		rs.close
		response.write "<tr><td colspan=3 align=""right""><br><input type=""Submit"" name=""doit"" value="" Delete checked "" class=""pbutton""></td></tr>" & vbCrLf
		response.write "</table>"
	end if
%>
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

















