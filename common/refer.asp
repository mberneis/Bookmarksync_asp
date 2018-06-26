<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<%
mainmenu = "TELL A FRIEND"
submenu  = ""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>Refer a Friend</title>
<link rel="StyleSheet" href="syncit.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#660066" vlink="#990000" alink="#CC0099">
<!-- #include file = "../inc/header.inc" -->
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="117" valign="top">
<!-- #include file = "../inc/bmenu.inc" -->
    </td>
    <td width="513" valign="top">
      <table width="513" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="2" bgcolor="#6633FF" width="503"><a href="../bms/"><img src="../images/horbtn_bmup.gif" width="120" height="18" border="0" alt="BookmarkSync"></a><a href="../collections/"><img src="../images/horbtn_collup.gif" width="147" height="18" border="0" alt="MySync Collections"></a><a href="../news/"><img src="../images/horbtn_newsup.gif" width="124" height="18" border="0" alt="QuickSync News"></a><a href="../express/"><img border="0" src="../images/horbtn_exup.gif" width="112" height="18" alt="SyncIT Express"></a></td>
          <td rowspan="2" width="10" valign="top"><img src="../images/rightbar.gif" width="10" height="55" alt=""></td>
        </tr>
        <tr>
          <td colspan="2" width="503" height="37" bgcolor="#000066"><img src="../images/heading_refer.gif" width="503" height="37" alt="SyncIT Tell A Friend"></td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
<%
Dim email, mailsend,refercnt

function goodemail (e)
	Dim eok
	eok = (instr(e,".")>0 and instr(e,"@") > 0)
	'if not eok then response.write (e & " is bad<br>") else response.write (e & " is good<br>")
	goodemail = eok
end function

email = request.querystring("email")
if email = "" then email = Session("email")
if email = "" then email = request.cookies("email")
if request.form("toemail") <> "" then
	Dim toemail, rs, name,allemail
	Dim subject, body, note,i

	email = request.form("email")
	allemail = request.form("toemail")
	if email <> "" and allemail <> "" then
		on error resume next
		set rs = conn.execute ("SELECT name FROM person WHERE email = '" & email & "'")
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		if not rs.eof then
			name = rs("name")
		else
			name = email
		end if
		rs.close
		ID = Session("ID")
		if ID = "" then ID = "0"
		note = replace(trim(request.form("note")),"'","''")
				'          1         2         3         4         5         6         7
				'012345678901234567890123456789012345678901234567890123456789012345678901
		subject =	"You have been invited to..."
		body =		"...Sign up for BookmarkSync"
		if note <> "" then
			body = body & ":" & vbCrLf & vbCrLf & name & " (" & email & ") wrote:" & vbCrLf & makereply(note, "> ")
		else
			body = body & "." & vbCrLf
		end if

		body = body & vbCrLf
		body = body &	"BookmarkSync is the award-winning, free tool that synchronizes all of"	& vbCrLf
		body = body &	"your Internet favorites or bookmarks for instant access from any"	& vbCrLf
		body = body &	"computer in the world.  BookmarkSync also gives you the tools"	& vbCrLf
		body = body &	"to publish links to useful Internet sites in your own bookmarks for"	& vbCrLf
		body = body &	"your friends, colleagues or the entire BookmarkSync community - it's"	& vbCrLf
		body = body &	"a great way to share information!"					& vbCrLf
		body = body & vbCrLf
		body = body &	"For more information, or to download BookmarkSync now, visit our web"	& vbCrLf
		body = body &	"site http://www.bookmarksync.com."					& vbCrLf
		body = body & vbCrLf
		body = body &	"Our commitment: helping you get the most out of the Internet. We look"	& vbCrLf
		body = body &	"forward to serving you as a customer."					& vbCrLf
		body = body & vbCrLf
		body = body &	"Best Regards," & vbCrLf & vbCrLf & "Michael Berneis" & vbCrLf & "President, SyncIT.com" & vbCrLf
		body = body &	"michaelb@syncit.com"

		Do
			I = instr(allemail,vbCrLf)
			if I = 0 then exit do
			toemail = left(allemail,I-1)
			allemail = right(allemail,len(allemail)-I)
			if goodemail(toemail) then domail email, toemail, subject, body			
		Loop
		if goodemail(allemail) then domail email, allemail, subject, body
		domail email, "refer@bookmarksync.com","Referral" & " from " & email,allemail

		response.write "<br><h2>Thank you for your referral!</h2>"
		response.write "<h3>We have sent your friends the following invitation E-mail:</h3>"
		response.write "<img src=""../images/psep.gif"" width=""491"" height=""10"" alt=""""><br>"
		response.write "<pre><b>Subject:</b> " & subject & "</b><br><b>Body:</b><br>" & body & "</pre>"
		mailsend = true

		on error resume next
		ID = Session("ID")
		if ID <> "" and ID <> 0 then
			set rs = conn.execute ("{ call bumprefer(" & ID & ")}")
			if err then dberror err.description & " [" & err.number & "]"
			on error goto 0
			refercnt = rs("refercnt")
			rs.close
		else
			refercnt=0
		end if
		if refercnt = 0 then		
			body = "Dear SyncIt member," & vbCrLf & vbCrLf
			body = body & "Thank you very much for referring your friends and colleagues to SyncIt." & vbCrLf
			body = body & "Your referral means the world to us!" & vbCrLf
			body = body & "Be sure to check out our site's new look at http://www.syncit.com and try out SyncIt Express and MySync Collections."& vbCrLf
			body = body & "" & vbCrLf
			body = body & "Best Regards,"								& vbCrLf
			body = body & ""									& vbCrLf
			body = body & "Michael Berneis"								& vbCrLf
			body = body & "President SyncIt.com"									& vbCrLf
			body = body & ""									& vbCrLf
			body = body & "michaelb@syncIt.com"								& vbCrLf
			body = body & "http://www.SyncIt.com"								& vbCrLf
			body = body & vbCrLf
			body = body & "How are we doing?  Please tell us how we can make you a happy customer."	& vbCrLf
			body = body & "Fill in our survey online at"						& vbCrLf
			body = body & "http://SyncIt.com/common/survey.asp"					& vbCrLf
			body = body & "We have a strict privacy policy (we are TRUSTe licensed) - "		& vbCrLf
			body = body & "we do not release names or contact information to any party outside of"	& vbCrlf
			body = body & "SyncIT.com."								& vbCrLf
			domail "support@bookmarksync.com",email,"Thank you for your referral",body
		end if
	else
		mailsend = false
	end if
end if
if not mailsend then
%>
            <br><h2>Recommend us to Your friends</h2>
            <p>Please enter the required information below so we can send an
            invitation to your friends.</p>
<script language="JavaScript"><!--
function validate(theForm)
{

  if (theForm.email.value == "")
  {
    alert("Please enter a value for the 'Your E-mail' field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.toemail.value == "")
  {
    alert("Please enter a value for the 'Your friend's e-mail addresses' field.");
    theForm.toemail.focus();
    return (false);
  }
  return (true);
}
//--></script><form method="POST" action="refer.asp" onsubmit="return validate(this)">
<table border="0">
  <tr>
    <td class="menub" align="right">Your E-mail:</td>
    <td>
        <p><input class="register" type="text" name="email" size="30" value="<%= email %>"></p>
    </td>
  </tr>
  <tr>
    <td class="menub" align="right">Your friend's&nbsp;<br>
 E-mail Addresses:<br>
      <br>
      (Enter as many as you like&nbsp; One Address per line)</td>
    <td>
        <textarea rows="6" name="toemail" cols="30" class="register"></textarea>
    </td>
  </tr>
  <tr>
    <td class="menub" align="right">Additional&nbsp;<br>
 Information&nbsp;<br>
 to send:&nbsp;</td>
    <td>
        <textarea rows="6" name="note" cols="30" class="register"></textarea>
    </td>
  </tr>
  <tr>
    <td class="menub" align="right"></td>
    <td align="right">
        <input border="0" src="../images/btninvite.gif" name="I1" type="image" width="62" height="16" alt="Invite">
    </td>
  </tr>
</table>
      </form>
<% end if %>
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
