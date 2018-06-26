<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
restricted ("../collections/invite.asp?" & request.querystring())

Dim ref, pubid, pid, menutitle, rs,cid,anonymous
Dim ps, user_id, tree, catid, title, description
Dim email,delsub,sql

mainmenu = "MODIFY"
submenu  = ""
response.expires=0

delsub = request("delsub")
if delsub <> "" then
	sql = "Delete from subscriptions where subscriptionid=" & replace(delsub,","," or subscriptionid=")
	conn.execute sql
end if

ref = request.querystring("ref")
if ref <> "" then mainmenu = ref
ID=Session("ID")
pubid = request.querystring("pid")
if pubid="" then pid = request.form("pid")
if pubid = "" then pubid = ncrypt(Session("curcol"))
if pubid = "" then myredirect ("../collections/publication.asp")
pid = ndecrypt(pubid)
if pid = 0 then
	myRedirect (request.ServerVariables("HTTP_REFERER"))
end if
Session("curcol") = pid

on error resume next
set ps = conn.execute ("Select * from publish where publishid = " & pid)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0

user_id = ps("user_id")
tree = ps("path")
catid = ps("category_id")
title = ps("title")
description = ps("description")
if ps("anonymous") then anonymous="checked"
ps.close
if user_id <> ID then myredirect ("default.asp")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync Modify Collection</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>
<script language="JavaScript">
<!--
function delit(pid) {
	if (confirm('Remove this collection?')) {
		location.href = 'publication.asp?delid='+pid;
	}
}
//-->
</script>
<body bgcolor="#FFFFFF" text="#000000" link="#3333FF" vlink="#990099" alink="#CC0099">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="111" valign="top">
<!-- #include file = "../inc/cmenu.inc" -->
    </td>
    <td width="511" valign="top">
      <table width="511" border="0" cellspacing="0" cellpadding="0">
<!-- #include file = "../inc/chead.inc" -->
          <td width="10">&nbsp;</td>
          <td width="491" valign="top">
              <img src="../images/iconpublish.gif" width="79" height="82" align="right" hspace="5" vspace="5" alt="Publish">
              <br><h2>Modify your collection.</h2>
		<br clear="all"><br><script language="JavaScript"><!--
function validate(theForm)
{

  if (theForm.title.value == "")
  {
    alert("Please enter a value for the 'title' field.");
    theForm.title.focus();
    return (false);
  }

  if (theForm.title.value.length > 50)
  {
    alert("Please enter at most 50 characters in the 'title' field.");
    theForm.title.focus();
    return (false);
  }

  if (theForm.description.value == "")
  {
    alert("Please enter a value for the 'description' field.");
    theForm.description.focus();
    return (false);
  }
  return (true);
}
//--></script><form method="POST" action="publication.asp" onsubmit="return validate(this)" name="edit_form">
				<input type="hidden" name="action" value="mod">
				<input type="hidden" name="pid" value="<%= pubid %>">
                <table border="0" cellpadding="2" cellspacing="2">
                  <tr> 
                    <td align="right" class="menub">FOLDER: </td>
                    <td colspan="3"><b><%= tree %></b></td>
                  </tr>
                  <tr> 
                    <td align="right" class="menub">TITLE: </td>
                    <td colspan="3"> 
                      <input type="text" name="title" class="register" value="<%= title %>" size="20" maxlength="50">
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" class="menub"> &nbsp;&nbsp;DESCRIPTION:&nbsp;</td>
                    <td colspan="3"> 
                      <textarea name="description" rows="6" class="register" cols="20"><%= description %></textarea>
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" class="menub">CATEGORY: </td>
                    <td colspan="3"> 
                      <select name="category" class="register">
                        <%
	on error resume next
	Set rs = conn.execute ("Select * from Category order by name")
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0

	while not rs.eof
		cid = rs("CategoryID")
		if catid = cid then
			response.write "<option selected value = " & cid & ">" & rs("name") & "</option>"
		else
			response.write "<option value = " & cid & ">" & rs("name") & "</option>"
		end if
		rs.movenext
	wend
	rs.close
%> 
                      </select>
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" class="menub">ANONYMOUS: </td>
                    <td align="right"> 
                      <div align="left"> 
                        <input type="checkbox" name="anonymous" value="checkbox" <%= anonymous %>>
                      </div>
                    </td>
                    <td align="right" valign="bottom"> 
                      <div align="center"><a href="javascript:delit('<%= pubid %>');"><img border="0" src="../images/btn_remove.gif" width="62" height="16" alt="Remove"></a><a href="publication.asp"><img border="0" src="../images/btncancel.gif" width="62" height="16" alt="Cancel"></a> 
                      </div>
                    </td>
                    <td align="right" valign="bottom">
                      <input type="image" border="0" name="imageField" src="../images/btnmodify.gif" width="62" height="16" alt="Modify">
                    </td>
                  </tr>
                </table>
            </form>
<% if catid = 0 then %>            
            <form method="post" action="editpub.asp">
            <table border=0 width="100%">
            <tr bgcolor=#000066><td colspan=3 class="menuw">SUBSCRIBERS</td></tr>
    
<%
			set rs = conn.execute ("select subscriptionid,email,name from subscriptions,person where publish_id = " & pid & " and subscriptions.person_id=person.personid")
			while not rs.eof
				response.write "<tr bgcolor=#EEEEEE><td><input type=""checkbox"" name=""delsub"" value=" & rs("subscriptionid") & "></td>"
				email = rs("email")
				response.write "<td>" & rs("name") & "</td><td><a href=""mailto:" & email & """>" & email & "</a></td></tr>"
				rs.movenext				
			wend
			rs.close
%>            
		</table>
		<input type=submit value="Remove checked Subscriber" class="pbutton">
            </form>
<% 
end if %>            
            </td>
          <td width="10" valign="top">&nbsp;</td>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>
</html>
