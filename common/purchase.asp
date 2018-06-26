<!-- #include file="../inc/asp.inc" -->
<%
	dim id,seed,fs,strTemp,pcarg,macarg,pc_license,mac_license,sale,url,x
	id = CLng(("0" & Session("PayID")))
	if id=0 then response.end	
	Set fs = CreateObject("Scripting.FileSystemObject")
	strTemp = fs.GetBaseName(fs.GetTempName)
	seed = Right(strTemp, Len(strTemp) - 3)
	strTemp = fs.GetBaseName(fs.GetTempName)
	seed = seed & Right(strTemp, Len(strTemp) - 3)
	set fs=Nothing
	Session("seed") = seed
	pc_license=(Session("pc_license") <> "")
	mac_license=(Session("mac_license") <> "")
	' 1 = $35: First License
	' 2 = $15: Additional License
	' 3 = $50: Both Licenses
	pcarg = (request("pc")=1)
	macarg = (request("mac")=1)
	if (pcarg and not(mac_license)) or (macarg and not(pc_license)) then 
		sale=1
	else
		sale=2
	end if
	if pcarg and macarg then sale=3
	select case sale
		case 1: url="1/Single_license"
		case 2: url="2/Additional_license"
		case 3: url="3/Combo_license"
	end select
	if pcarg and not macarg then url = url & "_PC"
	if not pcarg and macarg then url = url & "_MacOS"
	if pcarg and macarg then url = url & "_PC,MacOS"
	x = 0
	if pcarg then x = x + 1
	if macarg then x = x + 2
	
	response.redirect "http://www.clickbank.net/sell.cgi?link=syncit/" & url & "&x=" & x & "&s=" & sale & "&seed=" & seed & "&pid=" & ncrypt(id)
%>