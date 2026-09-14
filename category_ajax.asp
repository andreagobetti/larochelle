<%questofile="/category.asp"%>

<!--#include virtual="/setup.asp" -->
<!--#include virtual="/category_inc.asp" -->
<!--#include virtual="/dettaglio_inc.asp" -->

<%
	idsettore=request("idsettore")
	page = Request("page")
	ordinaPer=request("ordinaper")
	ipSize=request("ipSize")
	idmodello=request("idmodello")
	idtag=request("idtag")
	prezzo_nascosto=0
	if idsettore <>"" and not isnumeric(idsettore) then
		call determina_sqlinjection(idsettore)
		idsettore=""
	end if
	if page <>"" and not isnumeric(page) then
		call determina_sqlinjection(page)
		page=""
	end if
	if ordinaPer <>"" and not isnumeric(ordinaPer) then
		call determina_sqlinjection(ordinaPer)
		ordinaPer=""
	end if
	if ipSize <>"" and not isnumeric(ipSize) then
		call determina_sqlinjection(ipSize)
		ipSize=""
	end if
	if idmodello <>"" and not isnumeric(idmodello) then
		call determina_sqlinjection(idmodello)
		idmodello=""
	end if
	if idtag <>"" and not isnumeric(idtag) then
		call determina_sqlinjection(idtag)
		idtag=""
	end if
	
'session("categoryajax")=now()&"idsettore:"&idsettore&"page:"&page&" ordinaPer:"&ordinaper&" ipsize:"&ipsize&" idmodello:"&idmodello
if idsettore="" then idsettore=0
call carica_dizionario()
call elenco_articoli()

call rsclose()
%>

