      <table width="117" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="18" valign="top" class="bgmenu"><img src="../images/leftbar1.gif" width="18" height="56" vspace=5 alt=""></td>
          <td width="99" valign="top" bgcolor="#33FF00" background="../images/menu_express.gif"><img vspace=0 src="../images/expresslogo.gif" width="99" height="103" alt="">
            <div class="exdiv<%= BrowserType() %>">
              <p class="menub">
                <br>
<%
Private emenu, eimages

'emenu = Application("emenu")
If Not IsArray(emenu) Then
	emenu = Array(	_
		"HOME", 		"../express/default.asp",		_
		"DELIVER", 		"../express/deliver.asp",		_
		"RETRIEVE",		"../express/retrieve.asp",		_
		"ACTIVITY",		"../express/activity.asp",		_
		"-",			"-",				_
		"PROFILE",		"../common/profile.asp",	_
		"PRIVACY",		"../common/policy.asp",	_
		"HELP",			"javascript:showhelp(7)",			_
		"-",			"-")
	Application("emenu") = emenu
End If

//		"PARTNERSHIP",		"../common/partner.asp",	_

eimages = Application("eimages")
If Not IsArray(eimages) Then
	eimages = Array(_
		"src=""../images/yemark.gif"" width=""9"" height=""15"" align=""absmiddle"" alt=""*""",		_
		"src=""../images/emark.gif"" width=""9"" height=""15"" align=""absmiddle"" alt=""&gt;""",	_
		"src=""../images/edot_on.gif"" width=""26"" height=""10"" align=""absmiddle"" alt=""*""",	_
		"src=""../images/edot_off.gif"" width=""26"" height=""10"" align=""absmiddle"" alt=""o""")
	Application("eimages") = eimages
End If

ShowMenu emenu, eimages
%>
              </p>
            </div></td>
        </tr>
        <tr>
          <td colspan="2" width="117" valign="top"><img src="../images/flagexpress.gif" width="117" height="84" alt=""></td>
        </tr>
      </table>
