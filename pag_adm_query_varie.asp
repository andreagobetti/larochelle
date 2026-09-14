<%


'Elimina record magazzino per idvara inesistente
sql="DELETE magazzino.idvara, varianti_a.IDvara, magazzino.* FRom magazzino INNER JOIN varianti_a oN magazzino.idvara = varianti_a.IDvara WHERE (((magazzino.idvara)<>0) AND ((varianti_a.IDvara) Is Null));"

'Elimina record magazzino per idvarb inesistente
sql="DELETE magazzino.idvarb, magazzino.*, varianti_b.IDvarb FRom magazzino INNER JOIN varianti_b oN magazzino.idvarb = varianti_b.IDvarb WHERE (((magazzino.idvarb)<>0) AND ((varianti_b.IDvarb) Is Null));"
carrello_esteso="select carrello.*, prodotti.codice, prodotti.articolo, prodotti.variante1, prodotti.variante2, prodotti.prezzo, prodotti.costo, prodotti.Promozione, prodotti.Sconto, prodotti.Prodata, varianti_a.variante_a, varianti_a.prezzo_ve_va, varianti_b.variante_b, prodotti.idfor, prodotti.fileimg, prodotti.aggiungi_a_ordine FRom varianti_b RIGHT JOIN (varianti_a RIGHT JOIN (carrello INNER JOIN prodotti oN carrello.idpro = prodotti.IDpro) oN varianti_a.IDvara = carrello.idvara) oN varianti_b.IDvarb = carrello.idvarb"

ordini_dett_esteso="select ordini_dett.*, magazzino.quantita_magazzino, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, varianti_a.variante_a, varianti_b.variante_b
FRom ((magazzino RIGHT JOIN ordini_dett oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JOIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JOIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara "

cerca_azienda="select utenti.iduser, utenti.Azienda, utenti.Cognome, utenti.Nome, azienda & " " & nome & " " & cognome AS ricerca FROM utenti "

cerca_articoli_new="select prodotti.*, idpro & codice & articolo & parole_ricerca AS ricerca FROM prodotti"




SELect ordini.idord, ordini.nord, ordini.iduser, ordini.data, ordini.stato, ordini.nome, ordini.cognome, ordini.azienda,ordini.regione, ordini.provincia, ordini.verde,  ordini.totale, ordini.indirizzo, ordini.citta,sommadiincassi.SommaDiimporto FROM ((ordini LEFT JOIN (select  incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi force index (idord) GROUP BY incassi.idord ) AS sommadiincassi ON ordini.idord = sommadiincassi.idord) LEFT JOIN utenti ON ordini.iduser = utenti.iduser) LEFT JOIN fatture on ordini.idord = fatture.idord


SELect ordini.idord, ordini.nord, ordini.iduser, ordini.data, ordini.stato, ordini.nome, ordini.cognome, ordini.azienda,ordini.regione, ordini.provincia, ordini.verde,  ordini.totale, ordini.indirizzo, ordini.citta,sommadiincassi.SommaDiimporto FROM (((ordini force index(PRIMARY) LEFT JOIN (select  incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi force index (idord) GROUP BY incassi.idord ) AS sommadiincassi ON ordini.idord = sommadiincassi.idord) LEFT JOIN utenti ON ordini.iduser = utenti.iduser) LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDFat order by ordini.idord desc LIMIT 0 , 20

SELect ordini.idord, ordini.nord, ordini.iduser, ordini.data, ordini.stato, ordini.nome, ordini.cognome, ordini.azienda,ordini.regione, ordini.provincia, ordini.verde, utenti.email, utenti.cellulare, utenti.telefono, ordini.totale, ordini.indirizzo, ordini.citta, ordini_fatture.idfat, ddt.IDddt, ddt.Nddt, fatture.Nfat,sommadiincassi.SommaDiimporto, utenti.voto, utenti.tipologia FROM ((((ordini force index(PRIMARY) LEFT JOIN (select  incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi force index (idord) GROUP BY incassi.idord ) AS sommadiincassi ON ordini.idord = sommadiincassi.idord) LEFT JOIN utenti ON ordini.iduser = utenti.iduser) LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDFat) LEFT JOIN ddt on ordini.idord = ddt.idord where ordini.eliminato=false and year(ordini.data)>=2015  order by ordini.idord desc LIMIT 0 , 20
%>