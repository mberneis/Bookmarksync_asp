<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
mainmenu = "HOME"
submenu  = ""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
<title>SyncIT BookmarkSync syncs favorites across computers for easy access</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="117" valign="top">
<!-- #include file = "../inc/ymenu.inc" --><a href="#start"><img src="../images/arrowmove.gif" border="0" WIDTH="117" HEIGHT="97">
</a>
    </td>
    <td width="513" valign="top">
      <table width="513" border="0" cellspacing="0" cellpadding="0">
<!-- #include file = "../inc/yhead.inc" -->
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <table width="491" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="330">
                  <a href="../tree/view.asp"><img border="0" alt="View your Bookmarks now" src="../images/bmillustration.gif" WIDTH="330" HEIGHT="370"></a>
                  </td>
                <td width="161" valign="top" align="right">
				<br><form method="GET" action="searchbm.asp">
              <span class="menub"> &nbsp;Search your Bookmarks</span>
              <input type="text" name="searchbm" size="10" style="width:150px" value></form>
	<% if ID + 0 = 0 then %>
                  <h4>Your bookmarks, favorites and favorite places at your fingertips from anywhere in the world.&nbsp;</h4>
                  <h4>Great for home computer users, travelers, telecommuters, teachers &amp; students, Internet lovers everywhere!</h4>
                  <h4>Always up to date - no importing or exporting files.&nbsp; Let us do the work while you enjoy the Internet!</h4>
	<% else 
              	Dim bmcnt,lastchanged,rs,license
              	on error resume next
              	set rs = conn.execute ("{ call countbookmarks(" & ID & ") }")
              	if err then dberror err.description & " [" & err.number & "]"
              	bmcnt = rs("cl")
              	rs.close
              	set rs = conn.execute ("Select lastchanged from person where personid = " & ID)
              	if err then dberror err.description & " [" & err.number & "]"
              	license = 1
              	lastchanged = rs("lastchanged")
              	rs.close

		on error goto 0

		response.write "<h4>Welcome "
		response.write Session("name")
		response.write "</h4>" & vbCrLf
		response.write "<p>You currently have <b>"
		response.write bmcnt
		response.write "</b> bookmark entries</p>" & vbCrLf
		if lastchanged <> "" then
			response.write "<p>The last change was at <b>"
			response.write lastchanged
			response.write "</b></p>" & vbCrLf
		end if

		'response.write "<p align='center'><a href='../inc/partner.asp'><img border='0' src='../images/btnpartner.gif'></a></p>" & vbCrLf
	end if %>
                </td>
              </tr>
            </table>
            <h3><a href="../download/" name="start"><img src="../images/download.gif" border="0" align="right" alt="Download" WIDTH="180" HEIGHT="65">It's
              Easy to Use BookmarkSync!</a></h3>
            <p><b>Getting Started</b>
            <p>Download and install BookmarkSync software,&nbsp;<br>
 a <img border="0" src="../images/trayicon.gif" WIDTH="16" HEIGHT="12"> will appear in the
				lower right hand corner of your screen.</p>
            <p>Click <img border="0" src="../images/trayicon.gif" WIDTH="16" HEIGHT="12"> to view your bookmarks and favorites.&nbsp;<br>
Click your bookmarks &amp; connect to your favorite web sites.</p>
<p>Adding or removing bookmarks is a snap - just add or delete them on
your browser. No files to import or export - we'll keep track of changes for you.</p>
<p><i>When you're traveling or using a new computer or browser</i>, you can visit
<a href="http://www.bookmarksync.com/">www.bookmarksync.com</a>, and log in with
your name and password to be connected to your bookmarks and favorites.</p>
            <p>It's that easy!
            </p>
          </td>
          <td width="10" valign="top">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
<!-- #include file = "../inc/footer.inc" -->
</body>

</html>













