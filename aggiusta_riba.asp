<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%   
'Rigenera scadenze per fatture riba

sql="SELECT SQL_CALC_FOUND_ROWS sommadiincassi.SommaDiimporto, fatture.IDFat, fatture.Nfat, fatture.totale_fattura, fatture.data, utenti.Cognome, utenti.Nome, utenti.Azienda, agenti.Nominativo, agenti.IDagente, utenti.iduser, fatture.pagamento, fatture.PA FROM ((fatture LEFT JOIN (select incassi.idfat, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idfat ) AS sommadiincassi ON fatture.IDFat = sommadiincassi.idfat) INNER JOIN utenti ON fatture.iduser = utenti.iduser) LEFT JOIN agenti ON utenti.idagente = agenti.IDagente where totale_fattura>0 and (SommaDiimporto<totale_fattura or isnull(SommaDiimporto))  order by fatture.idfat"

set rs=conn.execute(sql)


		do while not rs.eof
				if pagamento_fattura(rs("pagamento"),4) then
					response.write "idfat:"&rs("idfat")&"<br>"
					IDfat=rs("idfat")
					Nfat=rs("Nfat")
					idord=0
					data=rs("data")
					totale_fattura=cdbl(rs("totale_fattura"))
					pagamento=rs("pagamento")
					n=scadenze_riba(IDfat,Nfat,idord,data,totale_fattura,pagamento)
				end if
		
			rs.movenext
		loop







%>
