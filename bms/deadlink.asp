<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
Dim sql, rs, token,strtemp,strtoken
Dim bid, cnt, upsql

restricted("../bms/deadlink.asp")
ID=Session("ID")
mainmenu = "UTILITIES"
submenu  = "DEAD LINKS"
sql = "{ call showdeadlinks(" & ID & ") }"
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
<title>BookmarkSync: Dead Links</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>
<script language=javascript>
<!--
function CheckAll(checked) {
	var i,len = document.frmdead.elements.length;
	for( i=0; i<len; i++) {
		if (document.frmdead.elements[i].name=='deldup') {
			document.frmdead.elements[i].checked=checked;
		}
	}
	CheckSusp(checked);
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
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br>
            <h2>Dead Links</h2>
	<form name="frmdead" method="post" action="deadlink.asp">
<%
	if request.form("token") <> "" then
		if CLng(token) = CLng(request.form("token")) then
			Dim deldup, checkit

			' delete duplicates
			deldup= " " & request.form("deldup") & ","
			checkit= " " & request.form("checkit") & ","
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
					upsql= "{call delbookmark (" & ID & ",'" & strtoken & "')}"
					'response.write upsql
					connexecute upsql
				end if
			loop
			if cnt > 0 then
				connexecute ("{ call bumptoken(" & ID & ") }")
				token = token + 1
				response.write "<p><b>Selected dead links removed. <br>Please double-click on the SyncIT icon in your taskbar to re-synchronize</b></p>"
			end if
		else
			response.write "<p><b><font color=#FF6666>Your bookmarks have changed since your selection.<br>"
			response.write "Please re-select the dead links to be removed again.</font></b></p>"
		end if
	end if
	response.write "<input type=hidden name=""token"" value=""" & token & """>"
	on error resume next
	set rs = conn.execute (sql)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	if rs.eof then
		response.write "<p><b>You have no dead bookmark links</b></p>"
	else

%>  	The following list of bookmarks could not be accessed by our server.<br>Please review them carefully as
	they might be internal links or secure sites. <br>If you want to remove a specific bookmark check
	the Delete column.<br><br>

		<table width="491" border="0" cellpadding="0" cellspacing="0">
			<tr bgcolor="#000066"><th align="left"><font color="#FFFFFF">&nbsp;Folder</font></th><th align="left"><font color="#FFFFFF">Bookmark</font></th><th align="left"><font color="#FFFFFF">Status&nbsp;</font></th><th align="right"><font color="#FFFFFF">Delete ?&nbsp;</font></th></tr>
<%

		cnt = 0
		bid = 0
		while not rs.eof
			Dim deadurl, bg

			cnt = cnt + 1
			deadurl = rs("url")
			if left(deadurl,4) = "http" then
				Dim p, p1, p2

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
				if len(p2) < len(p) then p1 = replace(left(p,len(p) - len(p2) -1),"\"," | ")
				if bg = "#EEEEEE" then bg = "#FFFFCC" else bg = "#EEEEEE"
				response.write "<tr bgcolor=" & bg & "><td>" & p1 & "</td><td><a target=_blank href=""" & deadurl & """>" & p2 & "</a></td>"
				response.write "<td>" & rs("status") & "</td>"
				response.write "<td align=""center""><input name=deldup value = " & cnt & " type=checkbox></td>"
				response.write "</tr>"
			end if
			rs.movenext
		wend
		rs.close

%>  <tr><td colspan="4">
          <p align="right"><br><input type="image" border="0" src="../images/btn_delchecked.gif" name="I1" width="102" height="16" alt="Delete Checked">
          </p><p align="left"><input type=button class="pbutton" onClick="CheckAll(true)" value="Check All"> <input type=button onClick="CheckAll(false)" value="Clear All" class="pbutton"> </p>
      </td></tr>
		</table>
</form>
<% end if %>
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
