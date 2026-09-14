<%
'sub paging(page,order,iPageSize,strConn,strSql)
	
	' VARIABLES
	' page		current page
	' order		order type
	' iPageSize	number of record 4 page
	' strConn	connession string
	' strSql	sql string whit order type

	' Retrieve page to show or default to 1
	If page = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = CInt(page)
	End If
	' ---------------

	' Create recordset and set the page size
	Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.PageSize = iPageSize

	' You can change other settings as with any RS
	'objPagingRS.CursorLocation = adUseClient
	objPagingRS.CacheSize = iPageSize
	' Open RS
	'objPagingRS.Open strSql, StrConn, adOpenStatic, adLockReadOnly, adCmdText
objPagingRS.Open strSql, StrConn, 3, 3

	' Get the count of the pages using the given page size
	iPageCount = objPagingRS.PageCount

	' If the request page falls outside the acceptable range,
	' give them the closest match (1 or max)
	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
	If iPageCurrent < 1 Then iPageCurrent = 1

	' Check page count to prevent bombing when zero results are returned!
	If iPageCount = 0 Then
		noFound = true ' No records found!
	Else
		' Move to the selected page
		objPagingRS.AbsolutePage = iPageCurrent
	end if
'END sub
	
' ---------------------- LOOPING -----------------------------
'	' Start output with a page x of n line
'	' Loop through our records and ouput 1 row per record
'	iRecordsShown = 0
'	Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
'
'
'		' Increment the number of records we've shown
'		iRecordsShown = iRecordsShown + 1
'		objPagingRS.MoveNext
'	Loop
'	' Close DB objects and free variables
'	objPagingRS.Close
'	Set objPagingRS = Nothing
' ------------------------------------------------------------

' --------------PREV & NEXT BUTTON----------------------------
'	<%NameOfFile = Request.ServerVariables("Script_Name")
'	If iPageCurrent <> 1 Then%>
<!-- <a href="<%=NameOfFile%>?page=<%= iPageCurrent - 1 %>&order=<%=order%>">Prec.</a> --><% 
'   End If	    
'	If iPageCurrent < iPageCount Then %>
<!--	<a href="<%=NameOfFile%>?page=<%= iPageCurrent + 1 %>&order=<%=order%>">Avanti</a> --><%
'	End If
' ------------------------------------------------------------%>
