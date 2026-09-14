<%
	response.cookies("io")="io"
	response.cookies("io").expires = date()+100
	
	response.redirect "default.asp"
	%>