<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
'Valorizzo il campo totale_riga   
sql="select * from ordini_dett"
Set rs_or = Server.CreateObject("ADODB.Recordset")

rs_or.Open sql, conn, 3, 3
do while not rs_or.eof
	rs_or("totale_riga")=totale_riga(rs_or("prezzo"),rs_or("sconto_prodotto"),rs_or("quantita"))
	rs_or.update
	rs_or.movenext
loop
rs_or.close
response.write "ordini finiti<br>"
sql="select * from preventivi_dett"
Set rs_or = Server.CreateObject("ADODB.Recordset")

rs_or.Open sql, conn, 3, 3
do while not rs_or.eof
	rs_or("totale_riga")=totale_riga(rs_or("prezzo"),rs_or("sconto_prodotto"),rs_or("quantita"))
	rs_or.update
	rs_or.movenext
loop
rs_or.close
response.write "preventivi finiti<br>"
sql="select * from ordini_fornitori_dett"
Set rs_or = Server.CreateObject("ADODB.Recordset")

rs_or.Open sql, conn, 3, 3
do while not rs_or.eof
	rs_or("totale_riga")=totale_riga(rs_or("prezzo"),rs_or("sconto_prodotto"),rs_or("quantita"))
	rs_or.update
	rs_or.movenext
loop
rs_or.close
response.write "fornitori finiti<br>"

%>