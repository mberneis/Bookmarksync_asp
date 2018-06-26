      <table width="119" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="8" valign="top"> <img src="../images/leftbar.gif" width="8" height="56" alt=""><br>
              <br><br><br>
            <br>
          </td>
          <td width="111" background="../images/vermenu.gif" valign="top" bgcolor="#FF9933"><img src="../images/black.gif" width="111" height="1" alt=""><br>
            <div class="colldiv<%= BrowserType() %>">
              <p class="menub">
                
<%
Private cmenu, cimages
'cmenu = Application("cmenu")
if not IsArray(cmenu) then
	cmenu = Array(	_
		"HOME", 		"../collections/default.asp",			_
		"PUBLISHED",		"../collections/publication.asp",		_
		"PUBLISHED|VIEW",	"../tree/cview.asp?ref=PUBLISHED",		_
		"PUBLISHED|INVITE",	"../collections/invite.asp?ref=PUBLISHED",	_
		"PUBLISHED|MODIFY",	"../collections/editpub.asp?ref=PUBLISHED",	_
		"ADD NEW",		"../collections/addpub.asp",			_
		"SUBSCRIBED",		"../collections/subscription.asp",		_
		"SUBSCRIBED|VIEW",	"../tree/cview.asp?ref=SUBSCRIBED",		_
		"SUBSCRIBED|INVITE",	"../collections/invite.asp?ref=SUBSCRIBED",	_
		"-",			"-",						_
		"POPULAR",		"../collections/popular.asp",			_
		"POPULAR|VIEW",		"../tree/cview.asp?ref=TOP20",			_
		"POPULAR|INVITE",	"../collections/invite.asp?ref=TOP20",		_
		"POPULAR|SUBSCRIBE",	"../collections/subscription.asp?addid=-1",		_
		"RECENT",		"../collections/recent.asp",			_
		"RECENT|VIEW",		"../tree/cview.asp?ref=RECENT",			_
		"RECENT|INVITE",	"../collections/invite.asp?ref=RECENT",		_
		"RECENT|SUBSCRIBE",	"../collections/subscription.asp?addid=-1",		_
		"BROWSE",		"../collections/showcategory.asp",		_
		"BROWSE|VIEW",		"../tree/cview.asp?ref=BROWSE",			_
		"BROWSE|INVITE",	"../collections/invite.asp?ref=BROWSE",		_
		"BROWSE|SUBSCRIBE",	"../collections/subscription.asp?addid=-1",		_
		"SEARCH",		"../collections/searchcol.asp",			_
		"SEARCH|VIEW",		"../tree/cview.asp?ref=SEARCH",			_
		"SEARCH|INVITE",	"../collections/invite.asp?ref=SEARCH",		_
		"-",			"-",						_
		"PROFILE",		"../common/profile.asp",	_
		"-",			"-")
	Application("cmenu") = cmenu
end if

//		"PARTNERSHIP",		"../common/partner.asp",	_

cimages = Application("cimages")
If Not IsArray(cimages) then
	cimages = Array( _
		"src=""../images/yomark.gif"" width=""17"" height=""15"" align=""absmiddle"" alt=""*""",	_
		"src=""../images/omark.gif"" width=""17"" height=""15"" align=""absmiddle"" alt="">""",		_
		"src=""../images/cdot_on.gif"" width=""33"" height=""10"" align=""absmiddle"" alt=""*""",	_
		"src=""../images/cdot_off.gif"" width=""33"" height=""10"" align=""absmiddle"" alt=""o""")
	Application("cimages") = cimages
End If

ShowMenu cmenu, cimages
%>
              </p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
            <p><br>
            </p>
          </div></td>
        </tr>
        <tr>
          <td colspan="2" width="119" valign="top"><img src="../images/flag.gif" width="119" height="78" alt=""></td>
        </tr>
      </table>
