      <table width="117" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="18" valign="top" class="bgmenu"><img src="../images/leftbar1.gif" width="18" height="56" alt=""></td>
          <td width="99" valign="top" class="menulist" background="../images/menu_bluebg.gif" bgcolor="#6633FF"><img src="../images/black.gif" width="99" height="1" alt=""><br>
            <div class="commondiv<%= BrowserType() %>">
              <p class="menuw">
          
<%
Private bmenu, bimages
'bmenu = Application("bmenu")
If Not IsArray(bmenu) then
	bmenu = Array(	_
		"PROFILE",		"../common/profile.asp",		_
		"PROFILE|REMOVE ME",	"../common/removeuser.asp",		_
		"-",			"-",					_
		"ABOUT US",		"../common/aboutus.asp",		_
		"-",			"-")
	Application("bmenu") = bmenu
End If




bimages = Application("bimages")
If Not IsArray(bimages) then
	bimages = Array( _
		"src=""../images/ybmark.gif"" width=""9"" height=""15"" align=""absmiddle"" alt=""*""",		_
		"src=""../images/bmark.gif"" width=""9"" height=""15"" align=""absmiddle"" alt=""&gt;""",	_
		"src=""../images/bdot_on.gif"" width=""26"" height=""10"" align=""absmiddle"" alt=""*""",	_
		"src=""../images/bdot_off.gif"" width=""26"" height=""10"" align=""absmiddle"" alt=""o""")
	Application("bimages") = bimages
End If

ShowMenu bmenu, bimages
%>
            </div>
          </td>
        </tr>
        <tr>
          <td colspan="2" width="117" valign="top"><img src="../images/flagblue.gif" width="117" height="89" alt=""></td>
        </tr>
      </table>
