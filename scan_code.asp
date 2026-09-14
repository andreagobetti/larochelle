<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->

<%
	call add2log("chiamato"&queryeform(),0)
codice=request("code")
if codice<>"" then
	if session("lettura_codice")="" then session("lettura_codice")=0
	session("lettura_codice")=session("lettura_codice")+1
	add2log "scansionato il codice "&codice,1
	
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		js("isError")="true"
		js("errorMessage")="Questo codice articolo esiste già"
		js("numerolettura")=session("lettura_codice")
		js.Flush

	
end if
%>