<%
	if session("contatore")="" then
		session("contatore")=1
	else
		session("contatore")=session("contatore")+1
	end if
	
	
	response.write session("contatore")
	
%>