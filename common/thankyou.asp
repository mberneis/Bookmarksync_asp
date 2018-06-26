<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<!-- #include file = "../inc/mail.inc" -->
<%
mainmenu = "PURCHASE"
submenu  = ""

function cbinvalid (seed,cbpop)
	Dim validator
	Const secretkey="SyncIt_Rocks"
	set validator = Server.CreateObject("KV.Validator")
	cbinvalid  =  (0 = validator.Valid(seed, cbpop, secretkey))
	set validator=Nothing
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
<meta http-equiv="Content-Language" content="en-us">
<title>Purchase Confirmation</title>
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
          <td colspan="2" width="503" height="37" bgcolor="#000066"><img src="../images/heading_registration.gif" width="503" height="37" alt="SyncIT Purchase Confirmation"></td>
        </tr>
        <tr>
          <td width="12">&nbsp;</td>
          <td width="491" valign="top">
            <br><h2>Purchase Confirmation</h2>
<p>
<%
	Dim mailbody,oldseed,newseed,cbpop,cbreceipt,sql,name,pid,rs,newvalid,alertmsg,sdate,fdate,ndate,sale,x,email,errmsg
	errmsg=""
	oldseed = Session("seed")
	newseed = request("seed")
	cbpop = request("cbpop")
	cbreceipt = request("cbreceipt")
	pid = ndecrypt(request("pid"))
	sale=request("s")+0
	
	'if oldseed<>newseed then
	'	errmsg = "Bad seed: " & oldseed & "|" & newseed
	'	errmsg = errmsg & "Q:" & request.querystring()
	'end if
	if errmsg="" then
		if pid=0 then 
			errmsg="Bad User ID" & vbCrLf
			errmsg = errmsg & "Q:" & request.querystring()
		end if
	end if
	if errmsg="" then
		if cbinvalid(newseed,cbpop) then 
			errmsg = "Invalid seed decode for user #" & pid & vbCrLf
			errmsg = errmsg &  "Seed = " & newseed & " | cbpop=" & cbpop & " | check=" & cbinvalid (newseed,cbpop) & vbCrLf
			errmsg = errmsg & "Q:" & request.querystring()
		end if
		'errmsg="X"
	end if
	if errmsg="" then
		on error resume next
		set rs = conn.execute ("select * from cblog where receipt='" & cbreceipt & "'")
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		if not rs.eof then
			errmsg = "Known receipt " & cbreceipt & " purchased from " & rs("person_id") & " at " & rs("purchased") & vbCrLf
			errmsg = errmsg & "Q:" & request.querystring()
		end if
		rs.close
	end if
	if errmsg = "" then ' Finally - move to stored procedure
		sql = "Insert into cblog (person_id,receipt,seed,purchased,amount,query) values ("
		sql = sql & pid & ","
		sql = sql & "'" & cbreceipt & "',"
		sql = sql & "'" & newseed & "',"
		sql = sql & "getdate(),"
		select case sale
			case 1:	sql = sql & "40,"
			case 2:	sql = sql & "20,"
			case 3:	sql = sql & "50,"
			case else: sql = sql & sale & ","
		end select
		sql = sql & "'" & request.querystring() & "')"
		'response.write sql
		connexecute (sql)
		set rs = conn.execute ("select name,email from person where personid=" & pid)
		name = rs("name")
		email = rs("email")
		rs.close
		x = request("x") + 0
		mailbody = "Congratulations " & name & "!<br>"
		if (x and 1)=1 then 
			connexecute ("update person set pc_license=GetDate() where personid=" & pid )
			mailbody = mailbody & "You have purchased a PC license.<br>"
		end if
		if (x and 2)=2 then 
			connexecute ("update person set mac_license=GetDate() where personid=" & pid )
			mailbody = mailbody & "You have purchased a MacOS license.<br>"
		end if
		mailbody=mailbody & "<br>Your account has been activated. If you have not done so already please <b><a href=""../download/default.asp"">download</a></b> the client software from our site.<br>"
		mailbody = mailbody & "<BR><B>Your credit card statement will show a charge from ClickBank/Keynetics.<hR>Contact Information:</b><br><BR>SyncIt.com<br>315 7th Avenue, #21-D<br>New York, NY 10001<br>Tel: (212) 242-5442<br>"
		Mailbody = mailbody & "<hr><b>Your receipt# is " & cbreceipt & "</b>"
		Session("license") = true
		domail "sales@syncit.com",email,"SyncIt License Sales Receipt",htmltxt(MailBody)
		domail "sales@syncit.com","syncitlicense@berneis.com","SyncIt License Sales Receipt",htmltxt(MailBody)
	end if
		
	if errmsg <> "" then
		errmsg = Now() & ": Purchase Error"  & vbCrLf & errmsg
		response.write "<p><b><font color='red'>An error has occured with your transaction</font></b></p>"
		response.write "The error response is <hr>" & replace(errmsg,vbCrLf,"<br>") & "<hr>"
		response.write "An email has been sent to the SyncIt system administrators<br>"
		response.write "You will be contacted shortly at your email: <a href=""mailto:" & Session("email") & """>" & Session("email") & "</a><br>"
		response.write "If your email account is not valid please contact SyncIT.com with your receipt number (" & cbreceipt & ") at <a href=""mailto:sales@syncit.com"">sales@syncit.com</a>"
		domail "license@syncit.com","license@berneis.com","Transaction Error with UserID " & pid,htmltxt(errmsg)
	else
		response.write mailbody				
	end if
%>	
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
