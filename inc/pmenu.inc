      <table width="117" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="18" valign="top" class="bgmenu"><img src="../images/left_news.gif" width="18" height="62" alt=""></td>
          <td width="99" valign="top" bgcolor="#660099" background="../images/menu_news.gif"><img src="../images/newslogo.gif" width="99" height="77" alt="SyncIT News">
            <div class="newsdiv<%= BrowserType() %>">
              <p class="menuw"><br>
            
<%
Private pmenu, pimages
'pmenu = Application("pmenu")
if not IsArray(pmenu) then
	pmenu = Array(	_
		"SETUP", 		"../news/",			_
		"SETUP|VIEW", 		"",	_
		"-",			"-",				_
		"PROFILE",		"../common/profile.asp",	_
		"PRIVACY",		"../common/policy.asp",	_
		"HELP",			"javascript:showhelp(6)",		_
		"-",			"-""")
	Application("pmenu") = pmenu
end if

//		"PARTNERSHIP",		"../common/partner.asp",	_

pimages = Application("pimages")
If Not IsArray(pimages) Then
	pimages = Array(_
		"src=""../images/ypmark.gif"" width=""9"" height=""15"" align=""absmiddle"" alt=""*""",	_
		"src=""../images/pmark.gif"" width=""9"" height=""15"" align=""absmiddle"" alt=""&gt;""",	_
		"src=""../images/ndot_on.gif"" width=""26"" height=""10"" align=""absmiddle"" alt=""*""",	_
		"src=""../images/ndot_off.gif"" width=""26"" height=""10"" align=""absmiddle"" alt=""o""")

	Application("pimages") = pimages
End If

ShowMenu pmenu, pimages
%>
              </p>
            </div></td>
        </tr>
        <tr>
          <td colspan="2" width="117" valign="top"><img src="../images/flagnews.gif" width="117" height="84" alt=""></td>
        </tr>
      </table>
