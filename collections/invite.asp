<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<%
restricted ("../collections/invite.asp?" & request.querystring())

Dim ref, pid, pubid,I,allemail
Dim toemail, fromemail, fromname, pubtitle, body
Dim rs, token, MyMsg

submenu  = "INVITE"

function goodemail (e)
	Dim eok
	eok = (instr(e,".")>0 and instr(e,"@") > 0)
	'if not eok then response.write (e & " is bad<br>") else response.write (e & " is good<br>")
	goodemail = eok
end function

ref = request.querystring("ref")
if ref <> "" then mainmenu = ref
ID=Session("ID")
pubid = request.querystring("pid")
if pubid="" then pid = request.form("pid")
if pubid = "" then pubid = ncrypt(Session("curcol"))
if pubid = "" then myredirect "../collections/default.asp"
pid = ndecrypt(pubid)
if pid = 0 then
	myRedirect (request.ServerVariables("HTTP_REFERER"))
end if
Session("curcol") = pid

allemail = request.form("email")
fromemail = Session("email")
fromname = Session("name")
if allemail <> "" then
	Dim subject, comments, toname, sql

	comments = request.form("comments")
	pubtitle = request.form("pubtitle")


	token = request.form("pid")
'		012345678901234567890123456789012345678901234567890123456789012345678901
	body = "It is my pleasure to inform you that your friend " & fromname
 body = body & " thought you might like this collection of web sites:"
	body = makereply(body, "")
 body = body & vbCrLf
 body = body & "Collection: " & pubtitle					& vbCrLf & vbCrLf
	If comments <> "" Then
		body = body & makereply(comments, "> ") & vbCrLf
		body = body & vbCrLf
	End If
 body = body & "To preview and subscribe to this free invitation visit"		& vbCrLf
 body = body & "      http://syncit.com/tree/cview.asp?invite=" & token		& vbCrLf
 body = body & vbCrLf
 body = body & "BookmarkSync can keep your browser bookmarks synchronized too,"	& vbCrLf
 body = body & "and can make sure you always have an up-to-date copy of the"	& vbCrLf
 body = body & "bookmark publication " & pubtitle & "."				& vbCrLf

	subject = "BookmarkSync Invite from " & fromname

		Do
			I = instr(allemail,vbCrLf)
			if I = 0 then exit do
			toemail = left(allemail,I-1)
			allemail = right(allemail,len(allemail)-I)
			if goodemail(toemail) then domail fromemail, toemail, subject, body			
			sql = "Insert into invitations (personid,email,publishid,invited) Values ("
			sql = sql & ID & ",'" & toemail & "'," & pid & ",'" & Now() & "')"
			connexecute (sql)
		Loop
		toemail = allemail
		if goodemail(allemail) then domail fromemail, toemail, subject, body
		sql = "Insert into invitations (personid,email,publishid,invited) Values ("
		sql = sql & ID & ",'" & toemail & "'," & pid & ",'" & Now() & "')"
		connexecute (sql)

	
	MyMsg = "The invitation has been sent to <hr align=left size=1 width=""50%""><font color=""#770000"">" & replace(request.form("email"),vbCrLf,"<br>") & "</font><hr align=left size=1 width=""50%"" ><br>"
else
	on error resume next
	set rs = conn.execute ("Select person.name,publish.title from person,publish where person.personid = publish.user_id and publishid = " & pid)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	
	pubtitle = "'" & rs("title") & "' by " & rs ("name")
	rs.close
end if
if request.querystring("addbm") <> "" then
	MyMsg = "The collection has been added to your publications<br>Now it's time to invite someone.<br>"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<title>BookmarkSync Invite to Collection</title>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<link rel="StyleSheet" href="../inc/syncit.css" type="text/css">
</head>
<script language="JavaScript">
<!--
function checkit (f) {
	if ((f.email.value == "")) {
		alert ("Please enter one or more E-mail addresses on seperate lines");
		return (false);
	}
	else {
		return true;
	}
}
-->
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
            <br><img src="../images/iconinvite.gif" width="87" height="69" align="right" hspace="5" vspace="5" alt="Invite">
           <font color="#FF0000"><b><%= MyMsg %></b></font>
              <h2>Invite someone</h2>
		Share your collection with your friends. If you set the category
            of your collection as *Private* it will not appear in the public lists.<br>
            You can enter as many e-mail addresses as you like, one per line.<br>
            <br clear="all">
            <form method="POST" action="invite.asp" onsubmit="return checkit(this)">
            	<input type="hidden" name="pid" value="<%= pubid %>">
            	<input type="hidden" name="pubtitle" value="<%= pubtitle %>">
              <table border="0" cellpadding="2" cellspacing="2">
                <tr><td align="right" class="menub">Collection: </td><td><%= Pubtitle %></td></tr>
                <tr><td align="right" class="menub">From: </td><td><%= fromname %> [<i><%= fromemail %></i>]</td></tr>
                <tr>
                  <td align="right" class="menub">E-MAIL: </td>
                  <td>
                    <textarea name="email" rows="6" class="register" cols="30"></textarea>
                  </td>
                </tr>
                <tr>
                  <td align="right" class="menub"> &nbsp;&nbsp;COMMENTS:&nbsp;</td>
                  <td>
                    <textarea name="comments" rows="6" class="register" cols="30"></textarea>
                  </td>
                </tr>
                <tr>
                  <td align="right" class="menub">&nbsp;</td>
                  <td align="right">
                    <input type="image" border="0" name="imageField" src="../images/btninvite.gif" width="62" height="16" alt="Invite">
                  </td>
                </tr>
              </table>
            </form>
            <img src="../images/psep.gif" width="491" height="4" vspace="5" alt><br>
              <img src="../images/iconwordout.gif" width="149" height="81" align="right" alt="Get the word out!">
            <h2>Put your collection on the web.</h2>
            Any visitor to your web site can subscribe to your published collection.
            <br clear="all">
            <form>
            <div class="menub"> &nbsp;Copy the code below to your Web Page </div>
              <textarea name="textfield4" class="longbox" rows="4" cols="30"><a href="http://syncit.com/tree/cview.asp?invite=<%= pubid %>">
<img src="http://syncit.com/images/button4.gif" alt="MySync Bookmark Collection">
</a>
              </textarea>
            </form>
            <a target="_blank" href="/tree/cview.asp?invite=<%= pubid %>">
            <img vspace="5" border="0" src="/images/button4.gif" align="absmiddle" hspace="3" alt="MySync Bookmark Collection" width="88" height="33"></a>&nbsp;&nbsp;
            <b>&lt;--</b> It will appear <a target="_blank" href="/tree/cview.asp?invite=<%= pubid %>">
like this</a>
          </td>
          <td width="10" valign="top">&nbsp;</td>
      </table>
    </td>
  </tr>
</table>
<!-- #include file = "../inc/footer.inc" -->
</body>
</html>
