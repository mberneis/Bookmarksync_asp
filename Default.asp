<% name=Session("name")
if instr(request.servervariables("HTTP_ACCEPT"),"text/vnd.wap.wml") > 0 then 
if instr(request.servervariables("HTTP_ACCEPT"),"text/x-hdml") > 0 then 
Response.ContentType = "text/x-hdml" 
%>
<HDML VERSION="3.0" PUBLIC="TRUE" MARKABLE="TRUE" Title="SyncIT.com Mobile">
<ACTION TYPE=ACCEPT TASK=GO LABEL=Enter DEST="/wap/mainxml.asp">
    <DISPLAY>
		<br><CENTER><IMG SRC=/wap/syncit.bmp ALT="SyncIT.com Inc."><BR><br>
		<CENTER>www.SyncIT.com<br>    
		<% if name<>"" then response.write "<center>Welcome " & name %>
     </DISPLAY>
</HDML>
<%
else
Response.ContentType = "text/vnd.wap.wml" 
%>
<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
  <meta http-equiv="Cache-Control" content="max-age=0"/>
</head>

<card id="intro" ontimer="/wap/mainxml.asp"  newcontext="true">
<timer name="key" value="25"/>
<p align="center">
<img src="/wap/syncit.wbmp" alt="SyncIT.com" />
<% if name<>"" then response.write "<br/>Welcome " & name & "<br/>" %>
</p>
</card>
</wml>
<% 
end if
else 
Dim token, sql,rs,pid,title,description,bmcnt,lastchanged

function crypt (x) ' poor mans encryption ;-)
	Dim i, s

	for i = 1 to len(x)
		s = s & hex(asc(mid(x,i,1))-30+i)
	next
	crypt = s
end function


function ncrypt (x)
	ncrypt = crypt(CStr(Clng (x) * 18))
end function

Dim partner_id

partner_id = request.querystring("sp")	
if partner_id = "" then 
	partner_id = Session("partner_id")
else
	Session("partner_id") = partner_id
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<!-- Welcome to our bookmark management software site -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<title>bookmark management software - A SyncIT Service</title>
<meta http-equiv="Content-Language" content="en-us">
<meta NAME="author" CONTENT="bookmark management software - Michael Berneis, Terence Way">
<meta NAME="publisher" CONTENT="SyncIT.com Inc">
<meta NAME="copyright" CONTENT="(c) 1998-2000 SyncIt.com  Inc, New York, support@bookmarksync.com">
<meta NAME="keywords" CONTENT="bookmark management software">
<meta NAME="description" CONTENT="bookmark management software: BookmarkSync is software that keeps your browser bookmarks synchronized across multiple computers and browsers.  A download tool for explorer and netscape browsers running on windows95, windows98, windowsnt, mac on OS8.5, apple. This tool lets you upgrade from oneview or extend visto briefcase">
<meta NAME="page-topic" CONTENT="bookmark management software">
<meta NAME="page-type" CONTENT="Software Download">
<meta NAME="audience" CONTENT="All">
<meta NAME="MS.LOCALE" CONTENT="US">
<meta NAME="Content-Language" CONTENT="english">
<meta NAME="ABSTRACT" content="bookmark management software : A download tool for IE Explorer and netscape browsers running on windows95, windows98, windowsnt, mac on OS8.5, apple.  A SyncIT product: BookmarkSync is software that keeps your browser bookmarks synchronized across multiple computers and browsers.  This tool lets you upgrade from oneview or extend visto briefcase">
<meta NAME="Classification" content="bookmark management software : This tool lets you upgrade from oneview or extend visto briefcase.  A SyncIT product: BookmarkSync at http://www.bookmarksync.com is software that keeps your browser bookmarks synchronized across multiple computers and browsers.  A download tool for explorer and netscape browsers running on windows95, windows98, windowsnt, mac on OS8.5, apple. ">
<meta http-equiv="Publication_Date" CONTENT="may 1999">
<meta http-equiv="Custodian" content="SyncIt.com, USA">
<meta http-equiv="Custodian Contact" content="Mike Berneis, 917-4060243, webmaster@bookmarksync.com">
<meta http-equiv="Custodian Contact Position" content="CEO, SyncIt.com">
<meta NAME="ROBOTS" CONTENT="INDEX,FOLLOW">
<meta NAME="DC.Title" CONTENT="bookmark management software : BookmarkSync is software that keeps your browser bookmarks synchronized across multiple computers and browsers.  A download tool for explorer and netscape browsers running on windows95, windows98, windowsnt, mac on OS8.5, apple.  This tool lets you upgrade from oneview or extend visto briefcase">
<meta NAME="DC.Creator" CONTENT="Michael Berneis, Terence Way">
<meta NAME="DC.Subject" CONTENT="bookmark management software : Synchronization of bookmarks across the Internet">
<meta NAME="DC.Description" CONTENT="bookmark management software : A download tool for explorer and netscape browsers running on windows95, windows98, windowsnt, mac on OS8.5, apple.  A SyncIT product: BookmarkSync is software that keeps your browser bookmarks synchronized across multiple computers and browsers.  This tool lets you upgrade from oneview or extend visto briefcase">
<meta NAME="DC.Publisher" CONTENT="SyncIT.com">
<meta NAME="rating" CONTENT="General">
<link rel="Start" href="default.asp">
<link REV="made" href="mailto:support@bookmarksync.com">
<link rel="StyleSheet" href="common/syncit.css" type="text/css">
<style type="text/css"></style>
</head>

<body text="#000000" link="#3333FF" vlink="#3333FF" bgcolor="#FFFFFF">
<!-- This page is about bookmark management software --> 

  <table width="628"  align="center"border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td colspan="5"><img src="images/introbanner.gif" width="627" height="55" usemap="#Map" border="0" alt="BookmarkSync"><map name="Map"><area shape="rect" coords="117,3,237,16" alt="BookmarkSync" href="bms/default.asp"><area shape="rect" coords="239,2,384,16" Alt="Collections" href="collections/default.asp"><area Alt="News" shape="rect" coords="385,2,508,17" href="news/default.asp"><area shape="rect" Alt="Express" coords="509,3,622,17" href="express/default.asp"></map></td>
    </tr>
<% if ID = 0 then %>
  </table><table width="628"  align="center"border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td colspan="4">&nbsp;</td>
      <td colspan="3" align="right"><a href="common/login.asp#register"><img src="images/introregister.gif" width="120" height="25" vspace="5" border="0" alt="Register Now!"></a></td>
    </tr>
<% else %>
    <tr> 
      <td colspan="7">&nbsp;</td>
    </tr>
<% end if %>
  </table><table width="628"  align="center"border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td colspan="7"><img src="images/tblbanner1.gif" width="628" height="18" alt="Search all Bookmarks"></td>
    </tr>
    <% 
    ID = Session("ID")

    Dim cookiename,cookievalid
    cookiename=request.cookies("name")
    cookievalid = (request.cookies("email") <> "" and request.cookies("md5") <> "" and cookiename<>"")
    if request.querystring("logout") <> "" then cookievalid=false
    if (ID=0 or ID = "") and not cookievalid then 
    %>
    <tr valign="top"> 
      <td width="2" background="images/tblleft.gif" bordercolor="#000000"><img src="images/tblleft.gif" width="2" height="2"></td>
      <form method="POST" action="collections/searchcol.asp">
        <td width="194" bgcolor="#000066" align="center"> <img src="images/tpix.gif" width="189" height="1"> 
          <input type="text" name="searchcol" size="18" class="footer" xstyle="{width:180px;}">
        </td>
        <td width="40" bgcolor="#000066" valign="bottom"> <img src="images/tpix.gif" width="40" height="1"> 
          <input type="image" border="0" name="imageField" src="images/introok.gif" width="38" height="18" alt="OK">
        </td>
      </form>
      <form method="POST" action="common/login.asp">
        <td width="1" background="images/tblcenter2.gif" bgcolor="#330033"><img src="images/tblcenter.gif" width="1" height="2"></td>
        <td width="348" bgcolor="#000066"> <img src="images/tpix.gif" width="343" height="1"> 
          <span class="menuw"> &nbsp;&nbsp;EMAIL</span> 
          <input type="text" style="{width:100px;}" name="email" size="14" class="footer">
          <span class="menuw">&nbsp;&nbsp;PASSWORD</span> 
          <input type="password" style="{width:100px;}" name="pass" size="8" class="footer">
        </td>
        <td width="40" bgcolor="#000066" valign="bottom"> <img src="images/tpix.gif" width="40" height="1"> 
          <input type="image" border="0" name="imageField2" src="images/introok.gif" width="38" height="18" alt="OK">
        </td>
        <td width="3" background="images/tblright.gif" bgcolor="#000000"><img src="images/tblright.gif" width="3" height="2"></td>
	  <input type="hidden" name="url" value="/default.asp">
      </form>
    </tr>
<% else 
if ID > 0 then 
%>
<!-- #include file = "inc/db.inc" -->
	<%   
      	on error resume next
      	set rs = conn.execute ("{ call countbookmarks(" & ID & ") }")
      	if err then dberror err.description & " [" & err.number & "]"
      	bmcnt = "<br>&nbsp;&nbsp;YOU HAVE " & rs("cl") & " BOOKMARKS STORED."
      	rs.close
	on error goto 0
else
	Session("name") = cookiename
	bmcnt="<br>&nbsp;&nbsp;QUICK LOGIN TO YOUR BOOKMARKS"
end if
    %> 
    <tr valign="top"> 
      <td width="2" background="images/tblleft.gif" bordercolor="#000000"><img src="images/tblleft.gif" width="2" height="2"></td>
      <form method="POST" action="collections/searchcol.asp">
        <td width="194" bgcolor="#000066" align="center"> <img src="images/tpix.gif" width="189" height="1"> 
          <input type="text" name="searchcol" size="20" class="footer" style="{width:180px;}">
        </td>
        <td width="40" bgcolor="#000066" valign="bottom"> <img src="images/tpix.gif" width="40" height="1"> 
          <input type="image" border="0" name="imageField" src="images/introok.gif" width="38" height="18" alt="OK">
        </td>
      </form>
        <td width="1" background="images/tblcenter2.gif" bgcolor="#330033"><img src="images/tblcenter.gif" width="1" height="2"></td>
        
      <td width="318" bgcolor="#000066"> <img src="images/tpix.gif" width="315" height="1"> 
        <span class="menuw">&nbsp;&nbsp;<font color="#FFCC00">WELCOME <%= Session("name") %></font><%= bmcnt %></span> <a href="/tree/view.asp"><img src="images/introviewnow.gif" width="100" height="18" border="0" alt="View Bookmarks Now"></a></td>
        
      <td width="70" bgcolor="#000066" valign="bottom"> <img src="images/tpix.gif" width="70" height="1"> 
        <a href="common/login.asp?logout=true"><img src="images/intrologoff.gif" width="65" height="18" border="0" alt="Log Out"></a> 
      </td>
        <td width="3" background="images/tblright.gif" bgcolor="#000000"><img src="images/tblright.gif" width="3" height="2"></td>
    </tr>
    <% end if %> 
    <tr> 
      <td colspan="7"><img src="images/tblbottom.gif" width="628" height="8"></td>
    </tr>
    <tr> 
      <td colspan="7">&nbsp;</td>
    </tr>
  </table>

<table width="628" align="center" border="0" cellspacing="0" cellpadding="0">
<tr><td>BookmarkSync is now being released under the GPL license and will be available
on sourceforge very soon.
</td>
</tr>
</table>

<table width="628"  align="center"border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
      <td colspan="3" class="footer" align="center">[ <a href="bms/">BOOKMARKSYNC</a> 
        | <a href="collections/">MYSYNC COLLECTIONS</a> | <a href="download/">DOWNLOAD</a> | <a href="common/profile.asp">PROFILE</a> ] </td>
  </tr>
</table>

<p>&nbsp;</p><p><A HREF="http://search.yahoo.com/bin/search?p=bookmark+management+software"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/sitemap.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/bookmark_management/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/management_software/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/manager_software/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/bookmark_storage/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/unite_bookmarksets/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/organize_bookmarks/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/bookmark_manger/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/online_storage/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/access_bookmarks/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/convert_bookmarks/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/merge_bookmarks/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/synchronize_bookmarks/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/bookmarks/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/share_bookmarks/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/transfer_bookmarks/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.bookmark--manager.com/access/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.favorites-manager.com/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.favorites-manager.com/transfer_favorite_places/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.favorites-manager.com/convert_favorite/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.favorites-manager.com/unite_favorites/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.favorites-manager.com/organize_favorites/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
<A HREF="http://www.favorites-manager.com/merge_favorites/index.html"><IMG SRC="http://www.bookmark--manager.comimages/dotclear.gif" WIDTH="1" HEIGHT="1" BORDER="0" alt="bookmark management software"></A> 
</p></body>
<!-- Thank you for visiting our bookmark management software site -->
</html><%
end if 
on error resume next
conn.close
set conn=Nothing
if Session("popup") = "" then 
%>
<script>
self.name='syncitmain';
if (document.cookie == '') alert('You must at least enable temporary cookies to use this site!');
</script>
<% end if %>