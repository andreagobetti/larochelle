<!--#include virtual="/setup.asp" -->
<!--#include virtual="/sub_head.asp" -->
<%
	
	val=""
	response.write "isnumeric ("&val&")"&isnumeric(val)&"<br>"
	val=1
	response.write "isnumeric ("&val&")"&isnumeric(val)&"<br>"
	val="1,1"
	response.write "isnumeric ("&val&")"&isnumeric(val)&"<br>"
	val="1a"
	response.write "isnumeric ("&val&")"&isnumeric(val)&"<br>"
	val="1.1"
	response.write "isnumeric ("&val&")"&isnumeric(val)&"<br>"
	val="b"
	response.write "isnumeric ("&val&")"&isnumeric(val)&"<br>"
	%>