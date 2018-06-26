<script language="javascript">
<!--
function showhelp(topic) {
	if (topic == 0) { // Default
		helpurl=''; 
	}
	if (topic == 1) { // FAQ
		helpurl='html/frequently_asked_questions.htm'; 
	}
	if (topic == 2) { // MAC Client
		helpurl='html/release_notes_for_mac_version_1_0.htm'; 
	}
	if (topic == 3) { // PC Client
		helpurl='html/syncit_client_functions.htm'; 
	}
	if (topic == 4) { // Bookmarksync Menu
		helpurl='html/bookmarksync_navigation.htm'; 
	}
	if (topic == 5) { // Collections Menu
		helpurl='html/mysync_collections_navigation.htm'; 
	}
	if (topic == 6) { // News Menu
		helpurl='html/quicksync_news_navigation.htm'; 
	}
	if (topic == 7) { // Express Menu
		helpurl='html/syncit_express_navigation.htm'; 
	}
	if (topic == 8) { // Common Menu
		helpurl='html/main_menu_navigation.htm'; 
	}
		
	var helpwnd = window.open ("/help-2.0/default.asp?start=" + helpurl,"help","menubar=0,status=1,scrollbars=1,resizable=1");
	helpwnd.focus();
}
//-->
</script>
<% 	Dim partner_id,hrs,partnerurl,partnerlogo

	Session("ad") = Session("ad") + 1		
	partner_id = request.querystring("sp")	
	if partner_id = "" then 
		partner_id = Session("partner_id")
	else
		Session("partner_id") = partner_id
	end if
	if partner_id = "" then partner_id=0
	Select case partner_id
		case 2: ' Jurisline
			partnerlogo = "partners/jurisline/logo.gif"
			partnerurl = "http://www.jurisline.com"
			Session("partnerpub") = "712382"
		case 3: ' Hieros Gamos
			partnerlogo = "partners/hierosgamos/logo.gif"
			partnerurl = "http://www.hg.org"
			Session("partnerpub") = "716817"
		case 4: ' doct.org
			partnerlogo = "partners/doct.org/logo.gif"
			partnerurl = "http://www.doct.org"
			Session("partnerpub") = "716260"
		case 5: ' Tailwind
			partnerlogo = "partners/tailwind/logo.gif"
			partnerurl = "http://www.tailwind.com"
			Session("partnerpub") = "716457"
		case 6: ' MusicPlay
			partnerlogo = "partners/musicplay/logo.gif"
			partnerurl = "http://www.musicplay.de"
			Session("partnerpub") = "718693"
		case 7: ' Oktoberfest
			partnerlogo = "icons/octoberfest.gif"
			partnerurl = "http://germannews.com/oktoberfest/"
			Session("partnerpub") = "719518"
		case else:
			Session("partner_id") = 0
			Session("partnerpub") = ""
	end select
	Session("partnerlogo") = partnerlogo
	Session("partnerurl") = partnerurl
%>
<!--
<table align="center" width="630" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="130" align="center"><a target="_top" href="../"><img src="../images/logo1.gif" width="113" height="52" xvspace="7" hspace="9" border="0" alt="SyncIT.com"></a>
    <% if partnerlogo <> "" then response.write "<br><a href=""" & partnerurl & """><img border=0 src=""/" & partnerlogo & """></a>"
    %>
    </td>
    <td width="500" align="center">
<% if mainmenu = "PRIVACY" then
	response.write "<img src=""../images/oldbanner.gif"">"
else 
	Dim randid
	randomize
	'randid = 100000 + int(rnd (1) * 1000000)
	randid=123456 ' for now... - experiment with random values later
%>
<IFRAME WIDTH=468 HEIGHT=60 SRC="http://ad-adex7.flycast.com/server/iframe/SyncITcom/runofsite/<%= randid %>" scrolling="no" marginwidth=0 marginheight=0 frameborder=0 vspace=0 hspace=0>
<A target="_top" HREF="http://ad-adex7.flycast.com/server/click/SyncITcom/runofsite/<%= randid %>"><IMG BORDER=0 WIDTH=468 HEIGHT=60 SRC="http://ad-adex7.flycast.com/server/img/SyncITcom/runofsite/<%= randid %>"></A>
</IFRAME>

<% end if %>     

 
  </td></tr>
  <tr><td colspan=2><img src="../images/tpix.gif" width="630" height="5" alt=""></td></tr>
</table>
-->
