<!-- #include file = "../inc/asp.inc" -->
<% response.expires = -1 %>
<!-- #include file = "../inc/db.inc" -->
<%
Dim rs,name,pc_license,mac_license,temp
restricted ("../common/receipt.asp")
set rs = conn.execute ("Select name,address1,address2,city,state,zip,country,pc_license,mac_license from person where personid=" & Session("ID"))
name = rs("name")
if not isNull(rs("address1")) then name = name & "<br>" & rs("address1")
if not isNull(rs("address2")) then name = name & "<br>" & rs("address2")
if not isNull(rs("city")) then name = name & "<br>" & rs("city")
temp = rs("state") & " " & rs("zip")
if temp<>" " then name = name & "<br>" & temp
if not isNull(rs("country")) then name = name & "<br>" & rs("country")

pc_license = rs("pc_license")
mac_license= rs("mac_license")
rs.close
conn.close
set conn=Nothing
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>SyncIt Invoice/Receipt</title>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<basefont face="Arial,Helvetica">
<p align="right"><b>SyncIt.com<br>
</b>315 7th Avenue<br>
Suite 21-D<br>
New York, NY 10001</p>
<h1>Invoice/Receipt</h1>
<p><b>To</b>:<br> <%= name %></p>
<p><b>Ordered services for User #<%= ncrypt(ID) %>:</b></p>
<table border="1" width="100%" cellspacing="0" bordercolorlight="#000000" bordercolordark="#000000" bordercolor="#000000" cellpadding="3">
<tr><td><b>License</b></td><td><b>Amount</b></td><td><b>Purchased</b></td></tr>
<% 
Dim pc_price,mac_price
'if datediff ("d",pc_license,"11/5/2001") >= 0 then pc_price=50 else pc_price=40
'if datediff ("d",mac_license,"11/5/2001") >= 0 then mac_price=50 else mac_price=40
pc_price=50
mac_price=50

if isNull(mac_license) and not isNull(pc_license) then %>
  <tr>
    <td width="33%" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#000000">PC
      License&nbsp;&nbsp;&nbsp; </td>
    <td width="33%">US$&nbsp; <%= pc_price %></td>
    <td width="34%"><%= pc_license %></td>
  </tr>
<%
end if
if not isNull(mac_license) and isNull(pc_license) then
%>
  <tr>
    <td width="33%" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#000000">MAC
      License&nbsp;&nbsp;&nbsp; </td>
    <td width="33%">US$&nbsp; <%= mac_price %></td>
    <td width="34%"><%= mac_license %></td>
  </tr>
<%
end if
if not isNull(mac_license) and not isNull(pc_license) then
%>
  <tr>
    <td width="33%" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#000000">PC/MAC
      License&nbsp;&nbsp;&nbsp; </td>
    <td width="33%">US$&nbsp; <%= pc_price %></td>
    <td width="34%"><%= pc_license %></td>
  </tr>
<% end if %>
</table>
<p><b>Delivery of service</b>: online</p>
<p><b>Payment terms</b>: Payment received in full.</p>
<hr>
<p>Thank you for your business.</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<i>SyncIt.com Customer Service</i></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>

</html>
