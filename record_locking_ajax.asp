<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()
tabella=request.QueryString("tabella")
id_tabella=request.QueryString("id_tabella")
inattivo=request.querystring("inattivo")
prendi=request.querystring("prendi")
if prendi<>"" then
	rilascia=true
	result=prenota_record(tabella,id_tabella,inattivo,rilascia)
end if
rilascia=false
result=prenota_record(tabella,id_tabella,inattivo,rilascia)
response.write result
call rsclose()
%>