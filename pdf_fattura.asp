<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pdf_fattura_inc.asp" -->
<%

idfat=request("idfat")
if not isnumeric(idfat) or idfat="" then
	call esci(0,0)
end if
if session("iduser")="" then
	add2log "Tentativo di visualizare PDF fattura da utente non loggato"&vbcrlf&allCookies(),3
	call esci(8,idfat)
end if
n=pdf_fattura(idfat,"pdf")

call connclose()
call CheckConnChiusa()
%>


