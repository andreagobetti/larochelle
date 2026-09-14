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
	
	iPageStart=iPageSize*(iPageCurrent-1)
	
	' ---------------

	' Create recordset and set the page size
	Set objPagingRS = Server.CreateObject("ADODB.Recordset")

	' You can change other settings as with any RS
	'objPagingRS.CursorLocation = adUseClient
	' Open RS
	'objPagingRS.Open strSql, StrConn, adOpenStatic, adLockReadOnly, adCmdText
	strSQL = strSQL&" LIMIT "&iPageStart&" , "&iPageSize&""
	'strSQL="select SQL_CALC_FOUND_ROWS * FROM log LIMIT 0, 10"
	
	'response.write strSQL
    objPagingRS.Open strSQL, conn, 1
    'lngTotalRecords = conn.Execute("Select Found_Rows();")(0).Value
    'response.write "lngTotalRecords: "&lngTotalRecords

	'response.end

	' Get the count of the pages using the given page size
	lngTotalRecords=conn.Execute("Select Found_Rows();")(0).Value
	'Response.write "numrecord:"&lngTotalRecords
	iPageCount = int(clng(lngTotalRecords)/ipagesize)
	'Response.write "iPageCount:"&iPageCount

	' If the request page falls outside the acceptable range,
	' give them the closest match (1 or max)
	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
	If iPageCurrent < 1 Then iPageCurrent = 1

	' Check page count to prevent bombing when zero results are returned!
	If iPageCount = 0 Then
		noFound = true ' No records found!
	Else
		' Move to the selected page
		'objPagingRS.AbsolutePage = iPageCurrent
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
