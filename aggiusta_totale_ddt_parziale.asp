<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Crea intestazioni<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	
	set rs_ddt=conn.execute("select * from ddt where tipo='1'")
	do while not rs_ddt.eof
		idord=rs_ddt("idord")
		totale_merce=0
		set rs_dett=conn.execute("select ddt_dett_ordini.quantita, ordini_dett.prezzo, ordini_dett.sconto_prodotto from ddt_dett_ordini inner join ordini_dett on ddt_dett_ordini.iddett=ordini_dett.iddett where idddt="&idord)
		do while not rs_dett.eof 
			totale_merce= totale_merce+totale_riga(rs_dett("prezzo"),rs_dett("sconto_prodotto"),rs_dett("quantita"))
			rs_dett.MoveNext
		loop
		conn.execute("update ordini set totale_merce_ordine="&aggiusta_decimale(totale_merce,"mysql")&" where idord="&idord)
		call add2log("calcolo_totale_merce_ddt idord:"&idord &" totalemerce:"&totale_merce,0)
		rs_ddt.MoveNext
	loop
	
	
	response.write "ESEGUITO"
end if
call connclose()

%>