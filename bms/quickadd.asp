<!-- #include file = "../inc/asp.inc" -->
<%
mainmenu = "UTILITIES"
submenu  = "QUICK ADD"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>Quick Add</title>
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
<!-- * -->
            <br>
            <h2>QUICK ADD</h2>
            <p>Need to add bookmarks to your account from a foreign computer? It's easy! </p>
<% 
Dim agent

agent = Request.ServerVariables("HTTP_USER_AGENT")

Dim b_type
' 1 - Netscape
' 2 - MSIE
' 3 - Other?

If InStr(agent, "compatible") = 0 Then
   b_type = 1
ElseIf InStr(agent, "MSIE") > 0 Then
   b_type = 2
Else
   b_type = 3
End If

Dim b_platform
' 1 - Win
' 2 - Mac
' 3 - Other

If InStr(agent, "Win") > 0 Then
   b_platform = 1
ElseIf InStr(agent, "Mac") > 0 Then
   b_platform = 2
Else
   b_platform = 3
End If

If b_type = 1 And b_platform = 1 Then %>
 <p><b>Netscape on Windows</b>
     <ol><li>Right-click on the link below</li>
         <li>Select &quot;Add Bookmark&quot;</li></ol></p>
<% ElseIf b_type = 2 And b_platform = 1 Then %>
 <p><b>Internet Explorer on Windows</b>
     <ol><li>Right-click on the link below</li>
         <li>Select &quot;Add To Favorites...&quot;</li></ol></p>
<% ElseIf b_type = 2 And b_platform = 2 Then %>
 <p><b>Internet Explorer on MacOS</b>
    <ol><li>Hold down the control key</li>
        <li>Click on the link below</li>
        <li>Select &quot;Add Link to Favorites&quot;</li></ol></p>
<% ElseIf b_type = 1 and b_platform = 2 Then %>
 <p><b>Netscape on MacOS</b>
    <ol><li>Hold down the control key</li>
        <li>Click on the link below</li>
        <li>Select &quot;Add Bookmark for this Link&quot;</li></ol></p>
<% End If %>

<p align="center">
<A href="javascript:void(open('http://www.syncit.com/addbm.asp?title='+escape(document.title)+'&url='+escape(location.href),'SyncIT','height=350,width=400,location=no,scrollbars=no,menubar=no,toolbar=no,directories=no,resizable=yes'));"><big><b>Add to SyncIT</b></big></a>
</p>
<table border='0'>
 <tr><td valign="top" align="right"><b>Netscape on Windows</b></td>
     <td><ol><li>Right-click on the link above</li>
             <li>Select &quot;Add Bookmark&quot;</li></ol></td></tr>
 <tr><td valign="top" align="right"><b>Internet Explorer on Windows</b></td>
     <td><ol><li>Right-click on the link above</li>
             <li>Select &quot;Add To Favorites...&quot;</li></ol></td></tr>
 <tr><td valign="top" align="right"><b>Internet Explorer on MacOS</b></td>
     <td><ol><li>Hold down the control key</li>
             <li>Click on the link above</li>
             <li>Select &quot;Add Link to Favorites&quot;</li></ol></td></tr>
 <tr><td valign="top" align="right"><b>Netscape on MacOS</b></td>
     <td><ol><li>Hold down the control key</li>
             <li>Click on the link above</li>
             <li>Select &quot;Add Bookmark for this Link&quot;</li></ol></td></tr>
</table>
<p>
For convenient access you might want to move the bookmark to your Netscape personal toolbar folder, or store it directly in your IE Links folder.
</p>
<p>
Browse to a site you would like to bookmark and select the link "Add to SyncIT". A popup window will appear (you will need to login the first time it pops up) where you can select the folder for the new bookmark.
</p>
<p>
When you return to your own computer all the new sites will already be in your bookmark list.</p>
</p>
<!-- * -->
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
