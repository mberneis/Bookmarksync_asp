<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" --><%
Server.ScriptTimeout = 1000
Dim searchstr, sherlock, cntmax, showhelp,rs,spath,idx,oidx,mytitle,firstid,bpic,expiration

mainmenu = "SEARCH"
submenu  = ""
restricted("../bms/searchbm.asp?" & Request.querystring())
searchstr = trim(request("searchbm"))
sherlock = request("sherlock")
if sherlock <> "" then searchstr = sherlock
cntmax = request.form("cnt") +0
if cntmax = 0 then cntmax = request.querystring("cnt") +0
if cntmax = 0 then cntmax = 20 
showhelp = "No Records found"
idx = request.querystring("idx") 
oidx = idx
ID = Session("ID")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync - Search</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../common/syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#330099" vlink="#990000" alink="#CC0099" onLoad="document.searchfrm.searchbm.focus();document.searchfrm.searchbm.select()">
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
          <td width="491" valign="top"> <!-- * -->
            <img src="../images/icon_searchbm1.gif" width="104" height="56" align="right" vspace="10" alt="Search for Bookmarks"><br>
            <h2>Search Your Bookmarks</h2>
            <p class="menul">[<a href="../collections/searchcol.asp">Want to
            search inside publications? Click here!</a>]
            </p>
            <p>You can search for any web-page that you have bookmarked.
            
            </p>
            <form method="get" action="searchbm.asp" name="searchfrm">
              <span class="menub"> &nbsp;ENTER SEARCH TERM</span><br>
              <input type="text" name="searchbm" class="register" size="30" value="<%= searchstr %>">
              <input type="image" border="0" name="imageField" src="../images/btnok.gif" width="62" height="16" align="absmiddle" alt="OK">
              <br><input type=checkbox name="checkold" value="X" <% if request("checkold") <> "" then response.write "CHECKED" %>><span class="menub"> &nbsp;Search also previously deleted bookmarks</span>
            </form>
<%

	if searchstr <> "" then
		Dim sql, cnt, lasturl, surl

		'if idx = "" then 
		'	sql = "SELECT top " & (cntmax *2 ) & " url, title, bookid FROM bookmarks WHERE FREETEXT(*, '" & searchstr & "') AND (status >= 200 AND status < 300) ORDER by bookid DESC"
		'else
		'	sql = "SELECT top " & (cntmax * 2) & " url, title, bookid FROM bookmarks WHERE FREETEXT(*, '" & searchstr & "') AND (status >= 200 AND status < 300) and bookid < " & idx & " ORDER by bookid DESC"
		'end if
		'sql = "SELECT top 32 url, title FROM bookmarks WHERE (title like '%" & searchstr & "%' or url like '%" & searchstr & "%') AND (status between 200 AND 299)"
		'sql = "{call searchbm('" & searchstr & "')}"
		
		sql = "select bookid,path as title,url,expiration from bookmarks,link where person_id = " & ID & " and bookid = book_id"
		if idx<>"" then sql = sql & " and bookid < " & idx
		if request("checkold") ="" then sql = sql & " and expiration is null"
		sql = sql & " and (link.path like '%" & searchstr & "%' or bookmarks.url like '%" & searchstr & "%') ORDER by bookid DESC"
		'if idx="" then
		'	sql = "select bookid,path as title,url from bookmarks,link where person_id = " & ID & " and expiration is null and bookid = book_id and (link.path like '%" & searchstr & "%' or bookmarks.url like '%" & searchstr & "%') ORDER by bookid DESC"
		'else
		'	sql = "select bookid,path as title,url from bookmarks,link where person_id = " & ID & " and expiration is null and bookid = book_id and bookid < " & idx & " and (link.path like '%" & searchstr & "%' or bookmarks.url like '%" & searchstr & "%') ORDER by bookid DESC"
		'end if

		'response.write sql & "<hr>"
		on error resume next
		set rs = conn.execute (sql)
		if err then
			response.write "<b><font color=""#FF0000"">Database unavailable...</font></b>"
			response.write "<br><h6>" & err.description & "</h6>"
		else
			if not rs.eof then
				cnt = 0
				lasturl = ""
				firstid = rs("bookid")
				while not rs.eof and cnt <= cntmax
					surl = rs("url")
					if right(surl,1) = "/" then surl = left(surl,len(surl)-1)
					spath = trim(rs("title") & "")
					'response.write spath & "<hr>"
					'if trim(spath) = "" then spath = "[No Title]"
					'if trim(spath) <> "" and surl <> lasturl and left(surl+"xxxxx",5) <> "file:"  then
					if left(surl+"xxxxx",5) <> "file:"  and surl <> lasturl then
						if len(spath) > 62 then spath = left(spath,60) & "..."
						spath = replace (spath,"<","&lt;")
						spath = replace (spath,">","&gt;")
						showhelp = ""
						if ("" = ("" & rs("expiration")))  then bpic="a" else bpic="r"
						if sherlock <> "" then
					    	response.write "<!-- item_start --><a href=" & surl & ">" & spath & "</a><!-- item_end --><br>" & vbCrLf
						else
							expiration = trim("" & rs("expiration"))
							if expiration="" then
						    	response.write "<img src=""../tree/a.gif"" alt=""""><span class=""tlist""><a target=_blank href=" & surl & ">" & spath & "</a><br></span>" & vbCrLf
						    else
						    	response.write "<img src=""../tree/r.gif"" alt=""""><span class=""tlist""><a target=_blank href=" & surl & ">" & spath & "</a> <small>[" & expiration & "]</small><br></span>" & vbCrLf
						    end if
				    	end if
						cnt = cnt + 1
						lasturl = surl
					end if
					idx = rs("bookid")
					rs.movenext

				wend
			else
				response.write "<b><font color=""#FF0000"">Nothing found.</font></b>"
			end if
		end if
		on error goto 0
	end if
%>
          </td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
      </table>
      <p class="menul">
<%
	if oidx <> "" then
		response.write "<a href=""javascript:history.back()"">"
		response.write "&lt;&lt;&lt; Prev</a>"
	end if
	if oidx <> "" and cnt+0 > cntmax+0 then response.write " || "
	if cnt+0 > cntmax+0 then
		response.write "<a href=""searchbm.asp?"
		response.write "searchbm=" & Server.URLEncode(searchstr) & "&cnt=" & cntmax & "&idx=" & idx & "&checkold=" 	& request("checkold")
		response.write """>Next &gt;&gt;&gt;</a>"
	end if	
	'response.write "<hr>" & cnt & "<hr>" & cntmax 
%>      
      </p>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>
</html>
