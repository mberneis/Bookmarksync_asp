<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
mainmenu = "UTILITIES"
submenu  = "AVANTGO"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<%
Dim extra_cookie
if ID > 0 then extra_cookie = "&amp;set_cookie=initid%3d" & ncrypt(ID) & "%3B"
%>
<head>
<title>BookmarkSync</title>
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
            <td width="491" valign="top"> <!-- ******************************************************************************************************************************************** --> 
              <br> 
              <h2>AvantGO Channel </h2>
              <p>Add links to your Bookmarkset on the Go.</p>
              <p>Through SyncIT's partnership with <a href="http://www.avantgo.com" target="_blank">AvantGo</a> 
                you are able to enter bookmarks while on the road. Once you synchronize 
                your palmpilot these bookmarks will be added to all your computers 
                where syncit is installed.</p>
              <p align="center"><b> <a href="https://avantgo.com/channels/detail.html?cha_id=1840&amp;cat_id=&amp;type=search_result&amp;data=syncit<%= extra_cookie %>" target="_blank"><img src="../images/avantgosubscribe.gif" width="110" height="31" border="0"></a> 
                </b></p>
              <p><!-- ******************************************************************************************************************************************** --> 
              </p>
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
