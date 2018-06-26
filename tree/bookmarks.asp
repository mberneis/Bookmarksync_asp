<%
code = request.querystring("code")
if code = "" then code = request.querystring("gcode")
if code <> "" then
	response.redirect "cview.asp?invite=" & code
else
	response.redirect "view.asp"
end if
%>