<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
'Cerca fatture senza PIVA
sql="select fatture.data, fatture.nfat, ordini.iduser, ordini.nome, ordini.cognome, ordini.piva,ordini.azienda, ordini.indirizzo, ordini.citta, ordini.provincia, ordini.regione, fatture.idord, ordini.nord, sommadiincassi.SommaDiimporto FROM ((utenti INNER JOIN (ordini INNER JOIN fatture ON ordini.idord = fatture.idord) ON utenti.iduser = ordini.iduser) LEFT JOIN (select incassi.idfat, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idfat ) AS sommadiincassi ON fatture.IDFat = sommadiincassi.idfat) where ordini.piva=''"
set rs=conn.execute (sql)
do while not rs.eof
	set rs_utente=conn.execute("select utenti.* from utenti where iduser="&rs("iduser"))

	if not rs_utente.eof then
		if rs_utente("piva")<>"" then
			response.write "fattura:"& rs("nfat")&"/"&year(rs("data"))&" Trovata piva:"&rs_utente("piva")&"<br>"
			
			conn.execute("update ordini set piva='"&rs_utente("piva")&"' where idord="&rs("idord"))
			
			
		end if
		
		
	end if


	rs.movenext
loop

%>
