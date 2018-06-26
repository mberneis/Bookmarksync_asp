<script language="JavaScript">
<!--

var head="display:''"
img1=new Image()
img1.src="../tree/f.gif"
img2=new Image()
img2.src="../tree/o.gif"

function nchange(f) {
	f1.style.display='';
	return false;
}

function change(){
   if(!document.all) {
      return
     }
   var noff,soff,myid;
   myid=event.srcElement.id;
   noff=5;soff=4;
   //alert (myid);
   if (myid =="fx1") {
   	myid="fx";
   	soff=0;noff=1;
   }
   if (myid=="fx") {
      var srcIndex = event.srcElement.sourceIndex
      var nested = document.all[srcIndex+noff]
      if (nested.style.display=="none") {
         document.all[srcIndex+soff].src="o.gif"
         nested.style.display=''
         //event.srcElement.style.listStyleImage="url(../tree/o.gif)"
      }
      else {
         document.all[srcIndex+soff].src="f.gif"
         nested.style.display="none"
         //event.srcElement.style.listStyleImage="url(../tree/f.gif)"
      }
   }
}

document.onclick=change

//-->

</script>

<%
sub tree(pubid,doedit)
	Dim target, sql, vs, rs,mytitle,NS6
	Dim vpath, catid, title, desc, uid, xpath
	Dim delim, showit, level, foldernum, vlen
	Dim prefix, i, myurl(32),MSIE,dirlevel
	
	MSIE = instr(request.servervariables ("HTTP_USER_AGENT") ,"MSIE") > 0
	NS6 = instr(request.servervariables ("HTTP_USER_AGENT") ,"Netscape6") > 0
	NS6 = false
	target=Session("target")
	if target="" then target="_blank"
	pubid = ndecrypt(pubid)
	if pubid=0 then
		sql = "{ call showbookmarks(" & ID & ") }"
	else
		sql = "Select User_ID,Path,Title, category_id, Description from publish where (PublishID = " & pubid  & ")"
		on error resume next
		Set vs = conn.execute (sql)
		if err then dberror err.description & " [" & err.number & "]"
		on error goto 0
		if not vs.eof then
			vpath = vs("path")
			catid = vs("Category_ID")
			title = vs("Title")
			desc = vs("Description")
			uid = vs("User_ID")
			xpath = vpath
			do
				i = instr(xpath,"\")
				if i = 0 then exit do
				xpath = right(xpath,len(xpath) - i)
			loop
			vlen = len(vpath) +1 '- len(xpath)
		end if
		sql = "{ call showpublication(" & pubid & ") }"
	end if
	on error resume next
	set rs = conn.execute (sql)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	delim = ""
	showit = true
	level = 1
	foldernum = 0
	if pubid > 0 then
		if uid = NewsID then
			if Left(lcase(title),5) = "news\" then title = right(title,len(title)-5)
			response.write "<H2>" & title & "</H2>"
		else
			response.write "<H2>" & title & "</H2><H4>" & desc & "</H4>"
		end if
	end if
	'response.write "vpath=" & vpath & " vlen=" & len(vpath) & "<hr>"
	'@@@response.write "<ul>"
	while not rs.eof
		Dim p, url
		p = rs("path")
		if vlen > 0 then p = right(p,len(p)-vlen)
		if delim = "" then
			delim = left(p,1)
			prefix = delim
		end if
		if delim <> left(p,1) then delim = left(p,1)
		while lcase(left(p,len(prefix))) <> lcase(prefix)
				response.write myurl(level)
				level = level - 1
				'@@@response.write "</ul>"	& vbCrLf
				response.write "</div>"	& vbCrLf
			do
				if len(prefix) = 0 then
					prefix = delim
					level = level + 1
					'@@@response.write "<ul>"	& vbCrLf
					response.write "<div>"	& vbCrLf
				else
					prefix = left(prefix,len(prefix)-1)
				end if
				if right(prefix,1) = delim then exit do
			loop
		wend
		Do
			dim f,fm

			i = instr(len(prefix)+1,p,delim)
			if i = 0 then exit do
			f = unprep(mid(p,len(prefix)+1,i-len(prefix)-1))
			fm = trim(left(p,i))
			foldernum = foldernum + 1
			if MSIE then
				if doedit then
					response.write "<div id=""fx"" class=""fi""><a href=""conf.asp?f=" & Server.UrlEncode(fm) & """><img src=""g.gif"" border=0></a><img src=""h.gif"" height=14 width=" & level*15 & "><img id=""fx1"" src=""f.gif"">" & f & "</div>" & vbCrLf
				else
					response.write "<div id=""fx"" class=""fi""><img src=""s.gif"" width=""0""><img src=""s.gif"" width=""0""><img src=""s.gif"" height=10 width=" & level*15 & "><img id=""fx1"" src=""f.gif"">" & f & "</div>" & vbCrLf
				end if
				response.write "<div id=""fl"" style=""display:none"" style=&{head};>" & vbCrLf
			else 
				if (NS6) then
					if doedit then
						response.write "<div id=""fx"" class=""fi""><a href=""conf.asp?f=" & Server.UrlEncode(fm) & """><img src=""g.gif"" border=0></a><img src=""h.gif"" height=14 width=" & (level)*15 & "><img id=""fx1"" src=""f.gif"">" & f & "</div>" & vbCrLf
					else
						response.write "<div id=""fx"" class=""fi""><img src=""s.gif"" width=""0""><img src=""s.gif"" width=""0""><img src=""s.gif"" height=10 width=" & (level)*15 & "><a href=""javascript:nchange(this);""><img border=0 id=""fx1"" src=""f.gif"">" & f & "</a></div>" & vbCrLf
					end if
					response.write "<div id=""fl"" style=""display:none"" style=&{head};>" & vbCrLf
					'response.write "<div id=""fl"" style=&{head};>" & vbCrLf
				else 
					if doedit then
						response.write "<div id=""fx"" class=""fi""><a href=""conf.asp?f=" & Server.UrlEncode(fm) & """><img src=""g.gif"" border=0></a><img src=""h.gif"" height=14 width=" & (level -1)*15 & "><img id=""fx1"" src=""o.gif"">" & f & "</div>" & vbCrLf
					else
						response.write "<div id=""fx"" class=""fi""><img src=""s.gif"" width=""0""><img src=""s.gif"" width=""0""><img src=""s.gif"" height=10 width=" & (level-1)*15 & "><img id=""fx1"" src=""o.gif"">" & f & "</div>" & vbCrLf
					end if
					response.write "<div id=""fl"" style=&{head};>" & vbCrLf
				end if
			end if			
			level = level + 1
			myurl(level) = ""
			prefix = left(p,i)
		Loop
		url = rs("url")	& ""
		if (instr(url,":") < 1) and (url <> "") then url = "http://" & url
		mytitle = right(p,len(p)-len(prefix))
		mytitle = replace(mytitle, "<","&lt;")
		mytitle = replace(mytitle, ">","&gt;")
		mytitle = unprep(mytitle)
		if MSIE then
			if doedit then
				if url <> "" then myurl(level) = myurl(level) & "<div class=""bi""><a href=""conf.asp?p=" & Server.UrlEncode(p) & """><img src=""g.gif"" border=0></a><img src=""h.gif"" height=14 width=" & level*15 & "><img src=""b.gif""><a target=""" & target & """ href=""" & url & """>" & mytitle & "</a></div>"	& vbCrLf
			else
				if url <> "" then myurl(level) = myurl(level) & "<div class=""bi""><img src=""s.gif"" width=""0""><img src=""s.gif"" height=10 width=" & level*15 & "><img src=""a.gif""><a target=""" & target & """ href=""" & url & """>" & mytitle & "</a></div>"	& vbCrLf
			end if
		else
			if doedit then
				if url <> "" then myurl(level) = myurl(level) & "<div class=""bi""><a href=""conf.asp?p=" & Server.UrlEncode(p) & """><img src=""g.gif"" border=0></a><img src=""h.gif"" height=14 width=" & level*15 & "><img src=""b.gif""><a target=""" & target & """ href=""" & url & """>" & mytitle & "</a></div>"	& vbCrLf
			else
				if url <> "" then myurl(level) = myurl(level) & "<div class=""bi""><img src=""s.gif"" width=""0""><img src=""s.gif"" height=10 width=" & level*15 & "><img src=""a.gif""><a target=""" & target & """ href=""" & url & """>" & mytitle & "</a></div>"	& vbCrLf
			end if
		end if
		rs.movenext
	WEnd
	while level > 0
		response.write myurl(level) & "</div>"
		level = level - 1
	wend
end sub
%>