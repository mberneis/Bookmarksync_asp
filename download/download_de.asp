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
            <h4>Klicken Sie auf den Download link f&uuml;r <%If b_platform = 1 Then %>Windows oder <% end if %>Macintosh<% if b_platform = 2 Then %> oder Windows<% End If %>.</h4>
            <p>Speichern Sie das Programm auf Ihrer Festplatte und f&uuml;hren Sie es danach aus.
            Das Installationsprogramm f&uuml;hrt Sie dann durch die Installation.<br>
            Sie m&uuml;ssen <a href="../inc/login.asp#register"><b>registriert</b></a> sein um das Programm zu benutzen.
            <br>
</p>
<% If b_platform = 1 Then %>

            <h4>Windows Client</h4>
<p> <b>OS</b>: 95, 98, NT4, 2000, ME<br>
            <b>Client Version</b>: 1.3<br>
<a href="../clienthelp/release13.html" target="_blank">Release Notes</a></p>
<p><b>Windows Version 1.3</b>
<ul>
 <li><a href="mp13setup.exe">mp13setup.exe</a> Windows 32-bit (365KB)
</ul></p>
            <h4>&nbsp;</h4>
<% End if %>
            <h4>Mac Client</h4>
            <p><b>Hardware:</b> Alle Power Macintosh (inclusive iMac and
            Powerbook)<b><br>
            OS</b>: System 7.6.1 or higher, MacOS 8 oder h&ouml;her bevorzugt.<br>
            <b>Client Version</b>: 1.0b22</p>
            <p><b>Mac Version 1.0b22 for OS 7.6.1</b><br>
<ul>
 <li><a href="SyncIT_b22_Installer_B.bin">SyncIT_b22_Installer_B.bin</a> BIN (680KB)
 <li><a href="SyncIT_b22_Installer_B.hqx">SyncIT_b22_Installer_B.hqx</a> HQX (933KB)
</ul>
            <p><b>Mac Version 1.0b22 f&uuml;r MacOS 8 oder h&ouml;her</b>
<ul>
 <li><a href="SyncIT_b22_Installer.bin">SyncIT_b22_Installer.bin</a> BIN (240KB)
 <li><a href="SyncIT_b22_Installer.hqx">SyncIT_b22_Installer.hqx</a> HQX (325KB)
</ul></p>
<% If b_platform = 2 Then %>

            <h4>Windows Client</h4>
<p> <b>OS</b>: 95, 98, NT4, 2000, ME<br>
            <b>Client Version</b>: 1.3<br>
<a href="../clienthelp/release13.html" target="_blank">Release Notes</a></p>
<p><b>Windows Version 1.3</b>
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
