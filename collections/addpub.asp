<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
restricted ("../collections/addpub.asp")

ID=Session("ID")
response.expires=0

mainmenu = "ADD NEW"
submenu  = ""

Dim i
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync Add New Collection</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>
<script language="JavaScript">
<!--
function fillfrm() {
	   lbl = document.publish_form.tree.options[document.publish_form.tree.selectedIndex].value;
	   if (document.publish_form.title.value == '') document.publish_form.title.value = lbl;
	   if (document.publish_form.description.value == '') document.publish_form.description.value = lbl;
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
            <br><h2>Add your own collection.</h2>
            <dir>
              <p>The folders into which you have saved your bookmarks and
              favorites appear below. Select the folder which you would like to
              publish. Give your collection a Title and Description, and select
              a category which best describes your bookmarks.&nbsp;<br>
              If you choose <b>* Private *</b>, your publication will not be
              available to anyone but you and the people you invite to share.</p>
              <p>You can make additions to your publication at any time from
              your browser by saving bookmarks into the corresponding folder. If
              you wish to remove or modify the title or description of an
              existing publication, click <a href="publication.asp">here</a>.&nbsp;</p>
            </dir>
<%
Dim rs
on error resume next
set rs = conn.execute ("{ call showbookmarks(" & ID & ") }")
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
if rs.eof then
	response.write "No bookmarks to publish yet."
else
%>
<script language="JavaScript"><!--
function validate(theForm)
{

  if (theForm.tree.selectedIndex <= 0)
  {
    alert("Please select a Folder to publish.");
    theForm.tree.focus();
    return (false);
  }

  if (theForm.category.selectedIndex <= 0)
  {
    alert("Please select a Category for your collection.");
    theForm.category.focus();
    return (false);
  }

  
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
//--></script><form action="publication.asp" method="POST" onsubmit="return validate(this);" name="publish_form">
			<input type="hidden" name="action" value="add">
                <table border="0" cellpadding="2" cellspacing="2">
                  <tr> 
                    <td align="right" class="menub">FOLDER: </td>
                    <td colspan="2"> 
                      <select name="tree" class="register" onclick="fillfrm();">
                        <option>[Pick Folder to Publish]</option>
                        <%
Dim delim, showit, publishit, level, foldernum, prefix

delim = ""
showit = true
publishit=true
level = 0
foldernum = 0

while not rs.eof
	Dim p
	p = rs("path")
	if delim = "" then
		delim = left(p,1)
		prefix = delim
	end if
	if delim <> left(p,1) then delim = left(p,1)
	while ucase(left(p,len(prefix))) <> ucase(prefix)
		if showit then
			level = level - 1
		end if
		do
			if len(prefix) = 0 then
				prefix = delim
				level = level + 1
			else
				prefix = left(prefix,len(prefix)-1)
			end if
			if right(prefix,1) = delim then exit do
		loop
	wend
	Do
		Dim f, spc, s, dbpath, csql, xs

	 	i = instr(len(prefix)+1,p,delim)
		if i = 0 then exit do
		' go forward...
		f = mid(p,len(prefix)+1,i-len(prefix)-1)
		spc = ""
		for s = 1 to level
			spc = spc & "&nbsp;-&nbsp; "
		next
		dbpath = prefix & f
		dbpath = right(dbpath,len(dbpath) -1)
		if len(prefix) > 1 then
			Dim froot

			froot = right(prefix,len(prefix)-1)
			froot = left(froot,len(froot)-1)
			csql = "select path from publish where (path = '" & replace(froot,"'","''") & "' or path like '" & replace(dbpath,"'","''") & "%') and User_ID =" & ID
		else
			csql = "select path from publish where path = '" & replace(f,"'","''") & "' and User_ID =" & ID
		end if
		on error resume next		
		set xs = conn.execute (csql)
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		if xs.eof then	response.write "<option value=""" & dbpath & """>" & spc & f & "</option>"  ' Directory
		xs.close
		level = level + 1
		prefix = left(p,i)
	Loop
	rs.movenext
WEnd
rs.close
%> 
                      </select>
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" class="menub">TITLE: </td>
                    <td colspan="2"> 
                      <input type="text" name="title" class="register" size="20" maxlength="50">
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" class="menub"> &nbsp;&nbsp;DESCRIPTION:&nbsp;</td>
                    <td colspan="2"> 
                      <textarea name="description" rows="6" class="register" cols="20"></textarea>
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" class="menub">CATEGORY: </td>
                    <td colspan="2"> 
                      <select name="category" class="register">
                      <option>[Please Pick Category]</option>
                        <%
	Set rs = conn.execute ("Select * from Category order by name")
	while not rs.eof
		response.write "<option value = " & rs("CategoryID") & ">" & rs("name") & "</option>"
		rs.movenext
	wend
	rs.close
	conn.close
	set conn=nothing
%> 
                      </select>
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" class="menub">ANONYMOUS: </td>
                    <td align="left">
                      <input type="checkbox" name="anonymous" value="checkbox">
                    </td>
                    <td align="right">
                      <input type="image" border="0" name="imageField" src="../images/btnadd.gif" width="62" height="16" alt="Add">
                    </td>
                  </tr>
                </table>
            </form>
<% 		end if
%>
            <h3>&nbsp;</h3>
            </td>
          <td width="10" valign="top">&nbsp;</td>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>
</html>
