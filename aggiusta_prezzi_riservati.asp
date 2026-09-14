<!--#include virtual="/setup.asp" -->
<%
sql_base="select prodotti.*, settori.nascondi_prezzi, categorie.nascondi_prezzi, categorie.nascondi, settori_prodotti.ordine FROM (settori INNER JOIN (prodotti INNER JOIN settori_prodotti ON prodotti.IDpro = settori_prodotti.IDpro) ON settori.idsettore = settori_prodotti.IDsettore) INNER JOIN categorie ON prodotti.idcat = categorie.IDcat"
rs.Open sql_base, conn, 3, 3
do while not rs.eof
		if  rs("settori.nascondi_prezzi") or rs("categorie.nascondi_prezzi") then
			 rs("prezzo_riservato")=true
			 rs.update
			 response.write rs("idpro")&"<br>"
		end if
	rs.movenext
loop

		

%>
