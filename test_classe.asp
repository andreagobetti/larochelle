<!--#include virtual="/setup.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%

'set ordine= (new ClasseOrdine)(array("ordine","new"))
set ordine= (new ClasseOrdine)(array("preventivi",752))

'set ordine= new ClasseOrdine
'response.write a.campo("idord")
'ordine.elenco_campi
'response.end
response.write ordine.esiste&"<br>"
response.write ordine.campo("idord")&"<br>"
val=ordine.calcolatotalemerce(false)
Response.write "totale_ok:"&val&"<br>"
response.write "nullo:"&isnull(val)&"<br>"
response.write "vuoto:"&(val="")&"<br>"
response.write "totale db:"&ordine.campo("totale")&"<br>"
response.write "ricalcolo:"&ordine.calcolatotale()&" verifica:"&ordine.totale_ok&"<br>"
response.write "totale merce db:"&ordine.campo("totale_merce_ordine")&"<br>"
response.write "ricalcolo merce:"&ordine.calcolatotalemerce(false)&" verifica merce:"&ordine.totale_merce_ok&"<br>"
response.write "totale db:"&ordine.campo("totale")&"<br>"
on error resume next
ordine.aggiungi("vocelibera")
ordine.elenco_articoli()
'ordine.elenco_campi


%>