<%
function crypt (x) ' poor mans encryption ;-)
	Dim i, s

	for i = 1 to len(x)
		s = s & hex(asc(mid(x,i,1))-30+i)
	next
	crypt = s
end function

function decrypt (x)
	dim s, k, i, h
	s = ""
	k = 30
	for i = 1 to len(x) step 2
		h =  Int("&H" & mid(x,i,2))
		k = k - 1
		s = s & chr(h+k)
	next
	decrypt = s
end function

function ncrypt (x)
	ncrypt = crypt(CStr(Clng (x) * 18))
end function

function ndecrypt (x)
	dim ntemp, ntemp2

	if x = "" then
		ndecrypt = 0
	else
		on error resume next
		ntemp = CLng(decrypt(x))
		if err then ntemp=-1234
		ntemp2 = CLng(ntemp / 18)
		if ntemp2 * 18 = ntemp then
			ndecrypt = ntemp2
		else
			ndecrypt = 0
		end if
	end if
end function

Dim qid

sub setbmid(id) 
	Session("id") = id
	qid=ncrypt(id)
end sub

function getbmid()
	id = request.form("id")
	if id="" then id=ndecrypt(request.querystring("qid"))
	if id="" then id=Session("id")
	if id<>"" then 
		setbmid(id)
		getbmid=id
	else
		getbmid=0
	end if
end function

tempid=getbmid()

%>