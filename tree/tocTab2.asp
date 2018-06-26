<%@ Language=VBScript %>
<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->

<%
	Function Samepath (arr1, arr2)  ' tests two arrays to see if two links have the same path
		Dim path
		path = Ubound(arr1)
		Dim i
		Samepath = true
		For i = 0 to (path-1)
			if Strcomp(arr1(i), arr2(i),1) = 1 then 
				Samepath = false
				exit for
			end if
		Next
	End Function

' there are a few more parts in the following code that can be put into separate functions, since they are used two or three times in the code,
' such as printing out the bookmarks, saving things into the queue, changing the menu string and iterating through the queue

	Dim SQL, res, pubid
	Dim count, queue_num, queue_max, level, max_level
	Dim queue() ' 1 is the path, 2 is the url, 3 is the level - not really a queue, but close enough
	Dim folder, split_cur, split_last, cur_len, i, j, k  'cur_len is the length of the path for the current folder
	Dim menu, trans, length, split_menu, path, url
	Dim path_arr, title
	
	pubid = ndecrypt(Request.QueryString("publication"))
	ID = ndecrypt(Request.QueryString("person"))
	if ID=0 then ID = Session("ID")
	if ID="" then ID=1
	menu = "0"
	response.write "// <pre>" & vbCrLf
	' this part might not work properly because ndecrypt is caled on the querystring: if ndecrypt is called on a blank value, then an actual value might result
	if pubid <> 0 then   ' if no publication is in the querystring
		SQL = "{ call showpublication(" & pubid & ") }"
	else
		SQL = "{ call showbookmarks(" & ID & ") }"
	End if
	
	on error resume next
	set res = conn.execute(SQL)
	if err then dberror err.description & " [" & err.number & "]"
	on error goto 0
	i = 0
	level = 0
	max_level = 0
	count = -1
	queue_num = 0
	queue_max = 20
	Redim queue(3, queue_max)
	cur_len = 1 ' base case with \url
	response.Write "var tocTab = new Array(); var ir=0;" & vbNewLine
	Response.Write "tocTab[ir++] = new Array (""Top"", ""Your Bookmarks"", """");" & vbNewLine
	count = 1
	split_last = split("\", "\")  ' start with nothing 
	do while not res.eof 
		' resizing queue if necessary
		if queue_num = queue_max then
			queue_max = queue_max * 2
			Redim Preserve queue(3, queue_max)
		End if
		
		' this part skips over any paths that are only folders, not actual links
		folder = true
		while folder = true and not res.eof
			if (right(res("path"),1) = "\") then
				res.movenext
			elseif len(res("url")) = 0 then
				res.movenext
			else
				folder = false
			End if
		wend
		if res.eof then exit do
				
		' split path of current bookmark into array
		path = res("path")
		url = res("url")
		if left(path, 1) <> "\" then path = "\" & path
		split_cur = Split(path, "\")
		
		' this is the first part of the if..elseif...else statement
		' if the number of folders in the current bookmark path is the same as in the last path
		' then it is either another bookmark in the same folder or the first bookmark in another folder at the same level as the previous bookmark
		if UBound(split_cur) = cur_len then
			if Samepath(split_cur, split_last) = true then ' if another bookmark in same folder, then save into the queue
				queue(1, queue_num) = path
				queue(2, queue_num) = url
				queue(3,queue_num) = level
				queue_num = queue_num + 1
			else  ' print out the bookmarks for the last folder delete them from the queue and save the current bookmark into the first empty space in the queue
				if queue_num>0 then i = queue_num - 1
				do while queue(3, i) = level  'loop to find index of first bookmark to be printed out by checking the level on each bookmark in the queue starting at the end
					i = i - 1
					if i<1 then exit do
				loop
				i = i+1
				j = i   ' hold place of next empty slot in queue
				' flush queue for that level, changing menu appropriately "1.1, 1.2, ....."  
				do while i < queue_num
					split_menu = Split(menu, ".") 
					trans = Clng(split_menu(Ubound(split_menu)))
					trans = trans +1
					menu = Left(menu, len(menu)- len(split_menu(Ubound(split_menu)))) & trans   
					path_arr = Split(queue(1,i), "\")
					title = path_arr(UBound(path_arr))
					title = replace(title,"""","\""")	
					response.write "tocTab[ir++] = new Array (""" & menu & """, """ & title & """, """ & replace(queue(2,i), """", "\""") & """);" & vbNewLine
					queue(1,i) = ""
					queue(2,i) = ""
					queue(3,i) = ""
					i = i+1
					count = count+1
				loop
				' add current link to queue
				split_menu = Split(menu, ".")
				if Ubound(split_menu) > 0 then menu = Left(menu, len(menu) - (len(split_menu(Ubound(split_menu)))+1))
				queue_num = j
				split_menu = Split(menu, ".")
				trans = Clng(split_menu(Ubound(split_menu)))
				trans = trans + 1
				menu = Left(menu, len(menu)- len(split_menu(Ubound(split_menu)))) & trans 
				Response.Write "tocTab[ir++] = new Array (""" & menu & """, """ & replace(split_cur(cur_len-1),"""","\""") & """, """");" & vbNewLine
				menu = menu + ".0"
				count = count +1
				queue(1,queue_num) = path
				queue(2,queue_num) = url
				queue(3,queue_num) = level
				queue_num = queue_num +1	
			End if
		' else if level of the current link is less than the last (path is shorter)
		elseif UBound(split_cur) < cur_len then ' folder level goes down
			'empty out folder levels until back down to level of current bookmark
			do while cur_len <> Ubound(split_cur)
				cur_len = cur_len-1
				if queue_num>0 then i = queue_num - 1
				do while queue(3, i) = level
					i = i - 1
					if i<1 then exit do
				loop
				i = i+1
				j = i 
				do while i < queue_num  
					split_menu = Split(menu, ".")
					trans = Clng(split_menu(Ubound(split_menu)))
					trans = trans+1
					menu = Left(menu, (len(menu) - len(split_menu(Ubound(split_menu))))) & trans 
					path_arr = Split(queue(1,i), "\")
					title = path_arr(UBound(path_arr))
					title = replace(title,"""","\""")
					response.write "tocTab[ir++] = new Array (""" & menu & """, """ & title & """, """ & replace(queue(2,i),"""","\""") & """);" & vbNewLine
					queue(1,i) = ""
					queue(2,i) = ""
					queue(3,i) = ""
					i = i+1
					count = count+1
				loop
				'right here too
				split_menu = Split(menu, ".")
				if Ubound(split_menu) > 0 then menu = Left(menu, len(menu) - (len(split_menu(Ubound(split_menu)))+1))
				split_menu = Split(menu, ".")
				trans = Clng(split_menu(Ubound(split_menu)))
				trans = trans + 1
				queue_num = j
				level = level-1
			loop
			' now deal with the current bookmark
			queue(1,queue_num) = path
			queue(2,queue_num) = url
			queue(3,queue_num) = level
			
		' else if the current bookmark is in a folder inside the current folder
		else   
			do while cur_len <> Ubound(split_cur) ' print out lines for each new folder in path while necessary
				level = level + 1
				if level > max_level then max_level = level ' for the "var nCols = ..." at the end of the file 
				cur_len = cur_len+1
				' add new subfolder
				split_menu = Split(menu, ".")
				trans = Clng(split_menu(Ubound(split_menu)))				
				trans = trans+1
				menu = Left(menu, len(menu)- len(split_menu(Ubound(split_menu)))) & trans
				Response.Write "tocTab[ir++] = new Array (""" & menu & """, """ & replace(split_cur(cur_len-1),"""","\""") & """, """");" & vbNewLine
				count = count+1
				menu = menu & ".0"
			loop
			queue(1,queue_num) = path
			queue(2,queue_num) = url
			queue(3,queue_num) = level
			queue_num = queue_num +1
		End if
		split_last = split_cur
		res.movenext
	loop	
	' empty out the rest of the queue here after last bookmark in the recordset
	i = 0
	do while i < queue_num
		if level <> 0 then  ' remove levels from menu string until the appropriate one is reached for the first bookmark in the queue
			split_menu = Split(menu, ".")
			k = Ubound(split_menu)
			while k > queue(3,i)
				menu = left(menu, len(menu) - len(split_menu(k))+1)
				k = k-1
			wend
		End if
		split_menu = Split(menu, ".")
		trans = Clng(split_menu(Ubound(split_menu)))
		trans = trans+1
		menu = Left(menu, len(menu)- len(split_menu(Ubound(split_menu)))) & trans
		path_arr = Split(queue(1,i), "\")
		title = path_arr(UBound(path_arr))
		title = replace(title,"""","\""")	
		response.write "tocTab[ir++] = new Array (""" & menu & """, """ & title & """, """ & replace(queue(2,i), """", "\""")  & """);" & vbNewLine
		count = count +1
		i = i+1
'		Response.Write len(queue(2,i)) & vbNewLine
	loop
	response.Write "var nCols = " & max_level+1 & ";" 
%>