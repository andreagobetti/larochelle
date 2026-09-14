<%
Response.Status="301 Moved Permanently"
Response.AddHeader "Location","http://"&Request.ServerVariables("SERVER_NAME") &"/index.asp"
%>