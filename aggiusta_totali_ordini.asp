<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
sql="select * from ordini"
Set rs_ordini = Server.CreateObject("ADODB.Recordset")
Set rs_or = Server.CreateObject("ADODB.Recordset")

rs_ordini.Open sql, conn, 3, 3
do while not rs_ordini.eof
	totale_merce=0
	totale_ordine=0
	totale_imponibile_ordine=0
	'if DateDiff("d","01/01/2011",rs_ordini("data"))<0 then 'prima del 01/01/2011
	'	metodo_conteggio=1
	if DateDiff("d","01/01/2015",rs_ordini("data"))<0 then 'prima del 01/01/2015
		metodo_conteggio=2
	else
		metodo_conteggio=3
	end if
	
	sql="select * from ordini_dett where idord="&rs_ordini("idord")

	rs_or.Open sql, conn, 3, 3
	do while not rs_or.eof
		if metodo_conteggio<3 then
			totale_riga_tmp=round((rs_or("prezzo")-(rs_or("sconto_prodotto")/100)*rs_or("prezzo"))*rs_or("quantita"),2)
		else
			totale_riga_tmp=totale_riga(rs_or("prezzo"),rs_or("sconto_prodotto"),rs_or("quantita"))
		end if
		rs_or("totale_riga")=totale_riga_tmp
		totale_merce=totale_merce+totale_riga_tmp
		rs_or.update
		rs_or.movenext
	loop
	rs_or.close
	rs_ordini("totale_merce_ordine")=totale_merce
	if isnull(rs_ordini("spese_bancarie_ordine")) then rs_ordini("spese_bancarie_ordine")=0

	if metodo_conteggio=1 then
		totale_imponibile_ordine=totale_merce+rs_ordini("spese_bancarie_ordine")
		
		if esenzione_iva(rs_ordini("trattamento_iva_ordine")) or  rs_ordini("iva_0") then
			valore_iva=0
		else
			'Calcolo iva arrotondata
			valore_iva=roundup(totale_imponibile_ordine*rs_ordini("iva_ordine")/100,2)
		end if
		rs_ordini("imposta_ordine")=valore_iva
		totale_ordine=totale_imponibile_ordine+valore_iva-rs_ordini("sconto_ordine")+rs_ordini("arrotondamento")+roundup(rs_ordini("trasporto"),2)
		
		
		
		
		
	else
		totale_imponibile_ordine=totale_merce+roundup(rs_ordini("trasporto"),2)+rs_ordini("spese_bancarie_ordine")
		
		if esenzione_iva(rs_ordini("trattamento_iva_ordine")) or  rs_ordini("iva_0") then
			valore_iva=0
		else
			'Calcolo iva arrotondata
			valore_iva=roundup(totale_imponibile_ordine*rs_ordini("iva_ordine")/100,2)
		end if
		rs_ordini("imposta_ordine")=valore_iva
		totale_ordine=totale_imponibile_ordine+valore_iva-rs_ordini("sconto_ordine")+rs_ordini("arrotondamento")
	end if


	
	
	if isnull(rs_ordini("arrotondamento")) then rs_ordini("arrotondamento")=0
	
	if totale_ordine<>roundup(rs_ordini("bk_totale_ordine"),2) then
		response.write "Totale differente:"&rs_ordini("idord")&" bk_totale:"&rs_ordini("bk_totale_ordine")&" totale:"&totale_ordine&"spese bancarie"&rs_ordini("spese_bancarie_ordine")&"<br>"
	else
		response.write "Totale ok:"&rs_ordini("idord")&"<br>"
	end if
	
	
	
	rs_ordini.update
	rs_ordini.movenext
loop
%>