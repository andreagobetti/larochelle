<%
'Verifica chiusure 09_12_2015
%>
<!--#include virtual="/setup.asp" -->
<%
on error resume next
if session("iduser")=1 then
	debug_pagina=true
else
	debug_pagina=false
end if
item_number = Request.Form("item_number")
if item_number="" then
	add2log "Iniziata procedura IPN Paypal transazione "& Request.Form("txn_id")&", item_number mancante,"&queryeform(),3
	response.end
end if
set rs=conn.execute ("select * from ordini where idord="&item_number)
if not rs.eof then nord=rs("nord")
rs.close
set rs = Nothing
add2log "Iniziata procedura IPN Paypal transazione "& Request.Form("txn_id")&" per [ordine="&Request.Form("item_number")&"]"&nord&"[/ordine]",2
Dim Item_name, Item_number, Payment_status, Payment_amount
Dim Txn_id, Receiver_email, Payer_email
Dim objHttp, str

'leggo la post da PayPal e preparo il comando
str = Request.Form & "&cmd=_notify-validate"

' rimando a paypal la post per la validazione
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
if debug_pagina=false then
	objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
else
	objHttp.open "POST", "https://www.sandbox.paypal.com/cgi-bin/webscr", false
end if
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str

' assegno le variabili mandate da paypal a variabili locali
item_name = Request.Form("item_name")
'item_number = Request.Form("item_number")
payment_status = Request.Form("payment_status")
txn_id = Request.Form("txn_id")
parent_txn_id = Request.Form("parent_txn_id")
receiver_email = Request.Form("receiver_email")
payer_email = Request.Form("payer_email")
reason_code = Request.Form("reason_code")
business = Request.Form("business")
quantity = Request.Form("quantity")
invoice = Request.Form("invoice")
custom = Request.Form("custom")
tax = Request.Form("tax")
option_name1 = Request.Form("option_name1")
option_selection1 = Request.Form("option_selection1")
option_name2 = Request.Form("option_name2")
option_selection2 = Request.Form("option_selection2")
num_cart_items = Request.Form("num_cart_items")
pending_reason = Request.Form("pending_reason")
payment_date = Request.Form("payment_date")
mc_gross = Request.Form("mc_gross")
mc_fee = Request.Form("mc_fee")
mc_currency = Request.Form("mc_currency")
settle_amount = Request.Form("settle_amount")
settle_currency = Request.Form("settle_currency")
exchange_rate = Request.Form("exchange_rate")
txn_type = Request.Form("txn_type")
first_name = Request.Form("first_name")
last_name = Request.Form("last_name")
payer_business_name = Request.Form("payer_business_name")
address_name = Request.Form("address_name")
address_street = Request.Form("address_street")
address_city = Request.Form("address_city")
address_state = Request.Form("address_state")
address_zip = Request.Form("address_zip")
address_country = Request.Form("address_country")
address_status = Request.Form("address_status")
payer_email = Request.Form("payer_email")
payer_id = Request.Form("payer_id")
payer_status = Request.Form("payer_status")
payment_type = Request.Form("payment_type")
notify_version = Request.Form("notify_version")
verify_sign = Request.Form("verify_sign")

'informazioni di sottoscrizione
subscr_date = Request.Form("subscr_date")
period1 = Request.Form("period1")
period2 = Request.Form("period2")
period3 = Request.Form("period3")
amount1 = Request.Form("mc_amount1")
amount2 = Request.Form("mc_amount2")
amount3 = Request.Form("mc_amount3")
recurring = Request.Form("recurring")
reattempt = Request.Form("reattempt")
retry_at = Request.Form("retry_at")
recur_times = Request.Form("recur_times")
username = Request.Form("username")
password = Request.Form("password")
subscr_id = Request.Form("subscr_id")

'informazioni utente
for_auction = Request.Form("for_auction")
auction_buyer_id = Request.Form("auction_buyer_id")
auction_closing_date = Request.Form("auction_closing_date")

' controllo la notifica
if objHttp.status <> 200 then
' gestisco l'errore HTTP
elseif (objHttp.responseText = "VERIFIED") then
elseif (objHttp.responseText = "INVALID") then
else
' error
end if
set objHttp = nothing

set rsp = Server.CreateObject("ADODB.Recordset")
rsp.open "pagamenti_paypal", conn, 2, 2
rsp.addnew

'aggiorno la tabella dati di pagamento paypal PAYMENT

rsp.Fields("item_name") = item_name 
rsp.Fields("item_number") = item_number
rsp.Fields("payment_date") = payment_date
rsp.Fields("pp_txn_id") = txn_id
rsp.Fields("parent_txn_id") = parent_txn_id
rsp.Fields("payment_status") = payment_status
rsp.Fields("pending_reason") = pending_reason
rsp.Fields("reason_code") = reason_code
rsp.Fields("txn_type") = txn_type
rsp.Fields("payment_type") = payment_type
rsp.Fields("mc_gross") = mc_gross
rsp.Fields("mc_fee") = mc_fee
rsp.Fields("payment_currency") = mc_currency
rsp.Fields("settle_amount") = settle_amount
rsp.Fields("settle_currency") = settle_currency
rsp.Fields("exchange_rate") = exchange_rate
rsp.Fields("payer_email") = payer_email
rsp.Fields("payment_status") = payer_status
rsp.Fields("cust_firstname") = first_name
rsp.Fields("cust_lastname") = last_name
rsp.Fields("cust_biz_name") = payer_business_name
rsp.Fields("gift_address_name") = address_name
rsp.Fields("cust_address_street") = address_street
rsp.Fields("cust_address_city") = address_city
rsp.Fields("cust_address_state") = address_state
rsp.Fields("cust_address_zip") = address_zip
rsp.Fields("cust_address_country") = address_country
rsp.Fields("cust_address_status") = address_status
rsp.Fields("notify_version") = notify_version
rsp.Fields("for_auction") = for_auction
rsp.Fields("auction_buyer_id") = auction_buyer_id
rsp.Fields("auction_closing_date") = auction_closing_date

'finish up
rsp.Update
rsp.Close
set rsp = Nothing
'add2log "Registrazione pagamento Paypal [ordine="&item_number&"]"&nord&"[/ordine] effettuata",2
'se non c'é un numero d'ordine salta l'aggiornamento del record dell'ordine, paypal invia ipn anche per transazioni via email
if instr(lcase(item_name),"ordine")>0 and item_number<>"" then
	'controllo se c'è una fattura
	idfat=conn.execute("SELECT IFNULL( (select idfat from ordini_fatture  where idord="&item_number&") ,null)")(0)
	
	
	
	
	'controllo se c'é l'ordine e se c'è lo metto a pagato
	mc_gross=replace(mc_gross,".",",")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "incassi", conn, 2, 2
	rs.addnew
	rs("data")=now()
	rs("importo")=mc_gross
	rs("causale")=3 'ALTRO
	rs("note_pagamento")="Pagemento Paypal transazione:" & txn_id&" Importo:"&mc_gross
	rs("idord")=item_number
	if not isnull(idfat) then
		rs("idfat")=clng(idfat)
	end if
	rs("tipo_pagamento")="PAYPAL AUTOMATICO"
	rs.update
	
	idincasso=Get_last_id("incassi")
	rs.close
	set rs = Nothing

	'add2log "Inserimento incasso effettuato",1
	

	
	add2log "Inserimento [incasso="&idincasso&"] Paypal per [ordine="&item_number&"]"&nord&"[/ordine] effettuato",2

	'fine aggiornamento ordine
end if


	if err.number<>0 then
		add2log "ERRORE in"&questofile&":"&Err.Description,0
	end if 

call connclose()
%>