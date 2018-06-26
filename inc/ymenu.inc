<script language=javascript>
<!--
function showremote() {
	var uagent=navigator.userAgent;
	var tempIndex = uagent.indexOf("MSIE");
	if(tempIndex != -1){
		bVer = parseInt(uagent.charAt(tempIndex + 5));
		if (bVer > 5) {
			open ("../tree/remote.asp?target=_main","_search");
			return;
		}
	} 
	window.open("../tree/remote.asp","sremote","width=300,height=screen.availHeight,menubar=0,toolbar=0,resizable=1,scrollbars=1");
}
//-->
</script>
      <table width="117" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="18" valign="top"><img src="../images/leftbar1.gif" width="18" height="56" alt=""></td>
          <td width="99" valign="top" background="../images/menu_bmbg.gif" bgcolor="#FFCC00" height="1" class="menub">
            <div class="bmdiv<%= BrowserType() %>"><img src="../images/black.gif" width="99" height="1" alt=""><br>
              <p class="menub">
          
<%
Private ymenu, yimages

'ymenu = Application("ymenu")
if not IsArray(ymenu) then
	ymenu = Array(	_
		"HOME", 		"../bms/",		_
		"VIEW", 		"../tree/view.asp",	_
		"VIEW|FRAME",		"../bms/frame.asp",	_
		"VIEW|REMOTE",		"#"" target=""sremote"" onclick=""showremote();", _
		"EDIT",			"../tree/edit.asp", _
		"UTILITIES",		"../bms/util.asp",	_
		"UTILITIES|QUICK ADD",	"../bms/quickadd.asp",	_
		"UTILITIES|AVANTGO",	"../bms/avantgo.asp",	_
		"UTILITIES|EXPORT",	"../bms/export.asp",	_
		"UTILITIES|UNDELETE", 		"../bms/undelete.asp",	_
		"UTILITIES|DUPLICATES",		"../bms/duplicate.asp",	_
 		"SEARCH",		"../bms/searchbm.asp",	_
		"-",			"-",			_
		"DOWNLOAD",		"../download/",		_
		"-",			"-",			_
		"PROFILE",		"../common/profile.asp")




	Application("ymenu") = ymenu
end if

		'"UTILITIES|IMPORT",	"../bms/export.asp",	
		'"UTILITIES|ROAMING",		"../bms/roaming.asp",	

yimages = Application("yimages")
If Not IsArray(yimages) then
	yimages = Array( _
		"src=""../images/rmark.gif"" width=""9"" height=""13"" align=""absmiddle"" alt=""*""",		_
		"src=""../images/ymark.gif"" width=""9"" height=""13"" align=""absmiddle"" alt=""&gt;""",	_
		"src=""../images/yydot_on.gif"" width=""23"" height=""8"" align=""absmiddle"" alt=""*""",		_
		"src=""../images/yydot.gif"" width=""23"" height=""8"" align=""absmiddle"" alt=""o""")
	Application("yimages") = yimages
End If

ShowMenu ymenu, yimages
%>
              </p>
            <p>&nbsp;</p>
            <p><br>
            </p>
          </div></td>
        </tr>
        <tr>
          <td colspan="2" width="117" valign="top"><img src="../images/flagbm.gif" width="117" height="47" alt=""></td>
        </tr>
      </table>
     
