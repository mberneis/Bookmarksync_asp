<!-- #include file = "../inc/asp.inc" -->
<%
'restricted ("../download/default.asp")
'ID=Session("ID")
mainmenu = "DOWNLOAD"
submenu  = ""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>Download BookmarkSync Client</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../common/syncit.css" type="text/css">
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
<%
Dim agent 
agent = Request.ServerVariables("HTTP_USER_AGENT")

DIm b_platform

If InStr(agent, "Win") > 0 Then
   b_platform = 1
Elseif InStr(agent, "Mac") > 0 Then
   b_platform = 2
Else
   b_platform =3
End If
%>
            <br><h2>Download</h2>
            <h4>Click the download link for <%If b_platform = 1 Then %>Windows or <% end if %>Macintosh<% if b_platform = 2 Then %> or Windows<% End If %>.</h4>
            <p>After downloading, the file will be saved in a program folder on
            your hard drive. To execute the program and install SyncIT, click on
            the SyncIT installation file and follow the install instructions.<br>
            You need to <a href="../common/login.asp#register"><b>register</b></a> at our
            site to use the client.<br>
</p>
<% If b_platform = 1 Then %>

            <h4>Windows Client</h4>
<p> <b>OS</b>: 95, 98, NT4, 2000, ME, XP<br>
            <b>Client Version</b>: 1.3.1<br>
<a href="../clienthelp/release13.html" target="_blank">Release Notes</a></p>
<p><b>Download Windows Version 1.3.1</b>
<ul>
 <li><a href="syncit131setup.exe">syncit131setup.exe</a> Windows 32-bit (384KB)
</ul></p>
            <h4>&nbsp;</h4>
<% End if %>
            <h4>Mac Client</h4>
            <p><b>Hardware:</b> Any Power Macintosh (including iMac and
            Powerbook)<b><br>
            OS</b>: MacOS 8 or higher recommended, or System 7.6.1 with Appearance Manager<br>
            <b>Client Version</b>: 1.0b27</p>
            <p><b>Download Mac Version 1.0b27</b>
<ul>
 <li><a href="SyncIT_b27_Install.hqx">SyncIT_b27_Install.hqx</a> HQX (222KB)
</ul></p>
<% If b_platform = 2 Then %>

            <h4>Windows Client</h4>
<p> <b>OS</b>: 95, 98, NT4, 2000, ME<br>
            <b>Client Version</b>: 1.3<br>
<a href="../clienthelp/release13.html" target="_blank">Release Notes</a></p>
<p><b>Download Windows Version 1.3</b>
<ul>
 <li><a href="syncit13setup.exe">syncit13setup.exe</a> Windows 32-bit (365KB)
</ul></p>
            <h4>&nbsp;</h4>
<% End If %>
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
