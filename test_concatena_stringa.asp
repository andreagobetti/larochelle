<!--#include virtual="/setup.asp" -->
<%
	
	val="ciao"

	call concatena_stringa(val,"<br>","ciao2")
	
	call concatena_stringa(val,"<br>","ciao3")
response.write val

	%>