<!--#include virtual="/multidict_cl.asp" -->
<!--#include virtual="/pag_adm_ordini_elenchi.asp" -->
<!--#include virtual="/ClasseMetodipagamento.asp" -->

<%
ordini_dett_esteso="select ordini_dett.*, magazzino.quantita_magazzino, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, varianti_a.codicevara,  varianti_a.variante_a, varianti_b.variante_b FRom ((magazzino RIGHT JOIN ordini_dett oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JOIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JOIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara "

if session("txt_aggiorna_ordine")<>"" then
	add2log "Errore txt_aggiorna_ordine:"&session("txt_aggiorna_ordine"),0
	Session.Contents.remove("txt_aggiorna_ordine")
end if

if session("scarica_distintabase")<>"" then
	add2log "scarica_distintabase:"&session("scarica_distintabase"),0
	Session.Contents.remove("scarica_distintabase")
end if
if session("mag_db")<>"" then
	add2log "Errore: mag_db:"&session("mag_db"),0
	session("mag_db")=""
	Session.Contents.remove("mag_db")
end if
if session("proc_ordina")<>"" then
	add2log "proc_ordina:"&session("proc_ordina"),0
	Session.Contents.remove("proc_ordina")
end if



'-------------------------Funzioni per incassi INIZIO
const incremento_ordine=10
if session("Errore incasso_su_fattura")<>"" then
	add2log session("incasso_su_fattura"),0
	Session.Contents.remove("incasso_su_fattura")
end if
sub elimina_incasso(idincasso)
	'Verifica chiusure 03_12_2015
	dim txt_nfat,txt_nord
	tipo_pagamento=split(Application("metodi_incasso"),vbcrlf)
	Set rs_incassi = Server.CreateObject("ADODB.Recordset")
	sql="select incassi.idincasso, ordini.idord, ordini.Nord, fatture.IDfat, fatture.Nfat FROM (incassi LEFT JOIN ordini ON incassi.idord = ordini.idord) LEFT JOIN fatture ON incassi.idfat = fatture.IDfat WHERE incassi.idincasso="&idincasso&";"
	rs_incassi.Open sql, conn, 1, 1
	idord=rs_incassi("idord")
	idfat=rs_incassi("idfat")
	txt_nord=rs_incassi("nord")
	txt_nfat=rs_incassi("nfat")
	rs_incassi.Close

	sql="select incassi.* FROM incassi WHERE incassi.idincasso="&idincasso&";"
	rs_incassi.Open sql, conn, 1, 3
	txt="Eliminato [incasso="&rs_incassi("idincasso")&"] da [ordine="&idord&"]"&txt_nord&"[/ordine]"
	if not isnull(idfat) then txt=txt&" e [fattura="&idfat&"]"&txt_nfat&"[/fattura]"
	txt=txt&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"

	txt=txt&vbcrlf&"Data:"&rs_incassi("data")&"<br>Importo:"&formatcurrency(rs_incassi("importo"))&"<br>Causale:"&causale_pagamento(rs_incassi("causale"))&"<br>Tipo:"&rs_incassi("tipo_pagamento")
	rs_incassi.delete
	rs_incassi.close
	set rs_incassi=nothing
	add2log txt ,2
	oper="view"
end sub
sub elimina_scadenza(idscadenza)
	'Verifica chiusure 03_12_2015
	dim txt_nfat,txt_nord
	Set rs_incassi = Server.CreateObject("ADODB.Recordset")
	sql="select * FROM scadenze WHERE idscadenza="&idscadenza&";"
	rs_incassi.Open sql, conn, 1, 3
	txt="Eliminato [scadenza="&rs_incassi("idscadenza")&"] "
	if not isnull(idfat) then txt=txt&" e [fattura="&idfat&"]"&txt_nfat&"[/fattura]"
	txt=txt&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"

	txt=txt&vbcrlf&"Data:"&rs_incassi("data")&"<br>Importo:"&formatcurrency(rs_incassi("importo"))&"<br>Tipo:"&rs_incassi("tipo_pagamento")
	rs_incassi.delete
	rs_incassi.close
	set rs_incassi=nothing
	add2log txt ,2
	oper="view"
end sub

sub carica_incasso(idord,idfat)
	'Verifica chiusure 03_12_2015
	'idord=request.form("idord")
	'idfat=request.form("idfat")
	'nfat=request.form("nfat")

	Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class

	dim carica_incasso, incassi_caricati, importo_cumulato
	carica_incasso=true
	incassi_caricati=false
	importo_cumulato=0


	log_txt="carica_incasso(idord:"&idord&",idfat:"&idfat&")<br>"
	session("incasso_su_fattura")=""
	importo=ccur(aggiusta_decimale(request.form("importo"),"asp"))
	var_causale=request.form("causale")
	if mod_larochelle then
		procedura="Incasso larochelle"
		log_txt=log_txt&"Incasso su fattura<br>"
		Set rs_incassi = Server.CreateObject("ADODB.Recordset")
		rs_incassi.Open "incassi", conn, 3, 3
		rs_incassi.addnew
		rs_incassi("data")=request.form("data")
		rs_incassi("importo")=importo
		rs_incassi("note_pagamento")=request.form("note_pagamento")
		tipo_pagamento=split(Application("metodi_incasso"),vbcrlf)
		rs_incassi("tipo_pagamento")=tipo_pagamento(cint(request.form("tipo_pagamento")))
		rs_incassi("idord")=0
		rs_incassi("idfat")=idfat
		rs_incassi("causale")=var_causale
		rs_incassi.update
		idincasso=get_last_id("incassi")
		rs_incassi.close
		set rs_incassi = nothing
		nfat=conn.execute("select nfat, idfat from fatture where idfat="&idfat)(0)
		if request.form("idscadenza")<>"" then
			conn.execute("delete from scadenze where idscadenza="&request.form("idscadenza"))
		end if

		log_txt=log_txt&"Inserimento [incasso="&idincasso&"] "&importo&"[/incasso] su [fattura="&idfat&"]"&nfat&"[/fattura]<br>"
		log_n=1

	else
		if idfat>0 then
			procedura="Incasso su fattura"
			log_txt=log_txt&"Incasso su fattura<br>"
			session("incasso_su_fattura")=session("incasso_su_fattura")&"idfat:"&idfat&"<br>"
			'Prendo il totale da pagare dall'importo fattura e il totale già pagato dalla somma degli incassi della fattura
			sql="select fatture.IDFat, fatture.totale_fattura, fatture.tipo_fattura, fatture.nfat, Sum(incassi.importo) AS SommaDiimporto, fatture.pagamento FROM fatture LEFT JOIN incassi ON fatture.IDFat = incassi.idfat GROUP BY fatture.IDFat, fatture.totale_fattura, fatture.nfat, fatture.pagamento HAVING (((fatture.IDFat)="&idfat&"));"

			set rs=conn.execute(sql)
			nfat=rs("nfat")
			tipo_fattura=rs("tipo_fattura")
			totale_fattura=cdbl(rs("totale_fattura"))
			somma_di_acconti=rs("SommaDiimporto")
			if isnull(somma_di_acconti) then
				somma_di_acconti=0
			else
				somma_di_acconti=cdbl(somma_di_acconti)
			end if
			totale_saldare=totale_fattura-somma_di_acconti
			n_scadenze=metodipagamento.campo(rs("pagamento"),"numero_scadenze")






			'Rivedere query con sub_idord
			'sql="select ordini_fatture.idfat, Sum(ordini.Totale) AS SommaDiTotale, Count(ordini.idord) AS ConteggioDiidord FROM ordini_fatture left JOIN ddt on ordini_fatture.idord = ddt.idord left join  ordini ON ddt.sub_idord = ordini.idord GROUP BY ordini_fatture.idfat HAVING (((ordini_fatture.idfat)="&idfat&"));"

			if tipo_fattura="ddt" then
				sql="select ordini_fatture.idfat, Sum(o2.Totale) AS SommaDiTotale, Count(o2.idord) AS ConteggioDiidord from ordini_fatture inner join ordini on ordini_fatture.idord =ordini.idord left join ddt on ordini_fatture.idord = ddt.idord  left join ordini o2 on ddt.sub_idord = o2.idord  GROUP BY ordini_fatture.idfat HAVING idfat="&idfat

			else
				sql="select ordini_fatture.idfat, Sum(o2.Totale) AS SommaDiTotale, Count(o2.idord) AS ConteggioDiidord from ordini_fatture inner join ordini on ordini_fatture.idord =ordini.idord left join ddt on ordini_fatture.idord = ddt.idord  left join ordini o2 on ordini_fatture.idord = o2.idord  GROUP BY ordini_fatture.idfat HAVING idfat="&idfat
			end if




			set rs=conn.execute(sql)
			n_ordini=clng(rs("ConteggioDiidord"))
			totale_ordini=cdbl(rs("SommaDiTotale"))
			set rs = Nothing


			log_txt=log_txt&"Importo fattura: "&totale_fattura&"<br>"
			log_txt=log_txt&"Numero ordini: "&n_ordini&"<br>"
			log_txt=log_txt&"Somma totale ordini: "&totale_ordini&"<br>"
			if totale_ordini<>totale_fattura then log_txt=log_txt&"<b>Totale fattura "&totale_fattura&" differente da totale ordini "&totale_ordini&" per "&(totale_fattura-totale_ordini) &"</b><br>"
			log_txt=log_txt&"Totale da saldare su fattura: "&totale_saldare&"<br>"
			if n_scadenze=0 then n_scadenze=1


		else
			procedura="Incasso su singolo ordine"
			log_txt=log_txt&"Incasso su singolo ordine<br>"
			'sql="select ordini.idord, ordini.totale, ordini.nord, Sum(incassi.importo) AS SommaDiimporto, ordini.tipopagamento FROM ordini LEFT JOIN incassi ON ordini.idord = incassi.idord GROUP BY ordini.idord, ordini.totale, ordini.pagamento HAVING (((fatture.IDFat)="&idfat&"));"
			sql="select ordini.idord, ordini.Nord, ordini.TipoPagamento, ordini.trattamento_iva_ordine, ordini.totale, Sum(incassi.importo) AS SommaDiimporto FROM ordini LEFT JOIN incassi ON ordini.idord = incassi.idord GROUP BY ordini.idord, ordini.Nord, ordini.TipoPagamento, ordini.trattamento_iva_ordine, ordini.totale HAVING (((ordini.idord)="&idord&"));"


			set rs=conn.execute(sql)
			nord=rs("nord")
			log_txt=log_txt&procedura&" "&nord&"<br>"
			totale_ordini=cdbl(rs("totale"))
			log_txt=log_txt&"Importo ordine: "&totale_ordini&"<br>"
			somma_di_acconti=rs("SommaDiimporto")
			if isnull(somma_di_acconti) then
				somma_di_acconti=0
			else
				somma_di_acconti=cdbl(somma_di_acconti)
			end if
			log_txt=log_txt&"Totale acconti: "&somma_di_acconti&"<br>"
			totale_saldare=totale_ordini-somma_di_acconti
			log_txt=log_txt&"Totale da saldare su ordine: "&totale_saldare&"<br>"
			if totale_saldare<0 then
				log_txt=log_txt&"<b>Totale da saldare negativo: "&totale_saldare&"</b><br>"
			end if
			n_scadenze=1

		end if

		log_txt=log_txt&"Importo incasso: "&importo&"<br>"
		log_txt=log_txt&"Totale acconti incassati: "&somma_di_acconti&"<br>"
		if importo>totale_saldare then
			log_txt=log_txt&"<b>Importo incasso: "&importo&" maggiore del totale da saldare: "&totale_saldare&"</b><br>"
			carica_incasso=false
		end if

		log_txt=log_txt&"scadenze: "&n_scadenze&"<br>"
		session("incasso_su_fattura")=session("incasso_su_fattura")&log_txt
		if carica_incasso then
			totale_fattura=totale_fattura/n_scadenze
			log_txt=log_txt&"Totale da versare per questa rata: "&totale_fattura&"<br>"
			'calcolo il rapporto versato/dovuto
			'add2log "1step: "&log_txt,1
			'session("verifica")="importo:"&importo&" totale_saldare:"&totale_saldare

			rapporto=cdbl(importo/totale_saldare)
			log_txt=log_txt&"Rapporto versato:"&importo&"/dovuto:"&totale_saldare&"="&rapporto&"<br>"


			'Apro ordini e incassi su ordini
			if idfat>0 then
				'Rivedere query

				if tipo_fattura="ddt" then
					log_txt=log_txt&"Applico query ddt<br>"

					sql="select ordini_fatture.idfat, o2.idord, o2.Nord, o2.totale, Sum(incassi.importo) AS SommaDiimporto from ordini_fatture inner join ordini on ordini_fatture.idord =ordini.idord left join ddt on ordini_fatture.idord = ddt.idord  left join ordini o2 on ddt.sub_idord = o2.idord LEFT JOIN incassi ON o2.idord = incassi.idord GROUP BY ordini_fatture.idord HAVING ordini_fatture.idfat="&idfat


				'sql="select ordini_fatture.idfat, ordini_fatture.idord, ordini.idord as idddt, ordini.nord as nddt, ordini.data as dataddt, ordini.tipo_documento,ordini.trasporto as trasporto_ddt, ddt.tipo_ddt, ddt.sub_idord, o2.idord, o2.nord, o2.trasporto, o2.data as dataordine,  o2.sconto_ordine, o2.iva_ordine,o2.trattamento_iva_ordine, ddt.annotazioni from ordini_fatture inner join ordini on ordini_fatture.idord =ordini.idord left join ddt on ordini_fatture.idord = ddt.idord  left join ordini o2 on ddt.sub_idord = o2.idord where idfat="&idfat




				else
					log_txt=log_txt&"Applico query ordine<br>"
					sql="select ordini_fatture.idfat, o2.idord, o2.Nord, o2.totale, Sum(incassi.importo) AS SommaDiimporto from ordini_fatture inner join ordini on ordini_fatture.idord =ordini.idord left join ddt on ordini_fatture.idord = ddt.idord  left join ordini o2 on ordini_fatture.idord = o2.idord LEFT JOIN incassi ON o2.idord = incassi.idord  GROUP BY ordini_fatture.idfat HAVING ordini_fatture.idfat="&idfat
				end if




			'	sql="select ordini_fatture.idfat, ordini.idord, ordini.Nord, Sum(incassi.importo) AS SommaDiimporto, ordini.Totale FROM (ordini_fatture INNER JOIN ordini ON ordini_fatture.idord = ordini.idord) LEFT JOIN incassi ON ordini.idord = incassi.idord GROUP BY ordini_fatture.idfat, ordini.idord, ordini.Nord, ordini.Totale HAVING (((ordini_fatture.idfat)="&idfat&")) ORDER BY ordini.Totale;"
			else
				sql="select ordini.idord, ordini.Nord, Sum(incassi.importo) AS SommaDiimporto, ordini.Totale FROM ordini LEFT JOIN incassi ON ordini.idord = incassi.idord GROUP BY ordini.idord, ordini.Nord, ordini.Totale HAVING (((ordini.idord)="&idord&")) "
			end if

			Set rs_ordini = Server.CreateObject("ADODB.Recordset")
			rs_ordini.Open sql, conn, 3, 3
			if n_ordini>1 then
				tipo_incasso=" ripartito"
			else
				tipo_incasso=" singolo"
			end if
			'apro incassi
			sql = "select * FROM incassi WHERE idincasso=0"
			Set rs_incassi = Server.CreateObject("ADODB.Recordset")
			rs_incassi.Open sql, conn, 3, 3

			n=1
			do while not rs_ordini.eof
				importo_incasso=rs_ordini("SommaDiimporto")
				if isnull(importo_incasso) then
					importo_incasso=0
				else
					importo_incasso=cdbl(importo_incasso)
				end if
				da_saldare=roundup(cdbl(rs_ordini("totale")),2)-importo_incasso
				log_txt=log_txt&"da_saldare=rs_ordini.totale"&rs_ordini("totale")&"-importo_incasso:"&importo_incasso&"="&da_saldare&"<br>"
				importo_rata=roundup(rapporto*da_saldare,2)
				log_txt=log_txt&"Processo ordine: "&rs_ordini("nord")&" acconto:"&importo_incasso&" totale:"&rs_ordini("totale")&" da saldare:"&da_saldare&" importo rata:"&importo_rata&"<br>"

				'SE importo_rata maggiore di da saldare ALLORA rettifico importo_rata
				if importo_rata>da_saldare then
					importo_rata=da_saldare
					log_txt=log_txt&"<b>Importo rata MAGGIORE del da saldare!</b><br>"
					log_txt=log_txt&"<b>Nuovo Importo rata: "&importo_rata&"</b><br>"
				end if

				if da_saldare>0 then
					incassi_caricati=true
					'SE ultima rata ALLORA conguaglio
					if n=n_ordini then
						importo_rata=importo-importo_cumulato
						log_txt=log_txt&"Conguaglio: importo_rata=importo:"&importo&"-importo_cumulato:"&importo_cumulato&"="&importo_rata&"<br>"
					end if
					'importo_rata=roundup(importo_rata,2)
					rs_incassi.addnew
					rs_incassi("data")=request.form("data")
					rs_incassi("importo")=importo_rata
					rs_incassi("note_pagamento")="Incasso"&tipo_incasso&" - "&request.form("note_pagamento")
					tipo_pagamento=split(Application("metodi_incasso"),vbcrlf)
					rs_incassi("tipo_pagamento")=tipo_pagamento(cint(request.form("tipo_pagamento")))
					rs_incassi("idord")=rs_ordini("idord")
					rs_incassi("idfat")=idfat
					rs_incassi("causale")=var_causale
					rs_incassi.update
					log_txt=log_txt&"Inserimento [incasso="&rs_incassi("idincasso")&"] "&importo_rata&"[/incasso] su [ordine="&rs_ordini("idord")&"]"&rs_ordini("nord")&"[/ordine]<br>"
					importo_cumulato=importo_cumulato+importo_rata

				end if
				rs_ordini.movenext
				n=n+1
			loop
			rs_ordini.close
			set rs_ordini=nothing
			rs_incassi.close
			set rs_incassi=nothing

			if importo_cumulato<>importo then
				log_txt=log_txt&"<b>Differenza tra importo versato "&importo&" e totale incassi caricati "&importo_cumulato&" di "&(importo-importo_cumulato)&"</b><br>"
			end if
			log_txt=log_txt&"<b>Totale importo incassi caricati: "&importo_cumulato&"</b><br>"
		end if	'if carica_incasso then
		if incassi_caricati=false then log_txt=log_txt&"<b>NESSUN INCASSO CARICATO</b><br>"
		if importo_cumulato>0 then
			log_n=2
			'visualizza_toastr("success|"&"<b>Inserimento incasso</b><br>" &log_txt)
		else
			visualizza_toastr("error|"&"<b>Inserimento incasso</b><br>" &log_txt)
			log_n=1
		end if

	end if	'id mod_larochelle


	add2log "<b>Distribuzione incasso</b><br>"&procedura &"<br>"&log_txt&vbcrlf&"Procedura:"&procedura,log_n
	'session("incasso_su_fattura")=""
	Session.Contents.remove("incasso_su_fattura")
end sub

function collega_incassi(idord,idfat)
if idord>0 then
	sql="UPDATE incassi SET incassi.idfat = "&idfat&" WHERE (((incassi.idord)="&idord&"));"
else
	sql="UPDATE incassi INNER JOIN ordini_fatture ON incassi.idord = ordini_fatture.idord SET incassi.idfat = "&idfat&" WHERE (((ordini_fatture.idfat)="&idfat&"));"
end if
conn.execute sql,num
collega_incassi=num
end function

function calcola_totale_fattura(idfat)
	dim vecchio_conteggio
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="select fatture.* FROM fatture WHERE idfat= " & idfat &";"
	rs1.open sql1,conn,3,3
	if DateDiff("d","01/01/2015",rs1("data"))<0 then 'prima del 01/01/2015
		vecchio_conteggio=true
	else
		vecchio_conteggio=false
	end if
	'Dati per generazione scadenze
	data=rs1("data")
	Nfat=rs1("nfat")
	fattura_sp=rs1("fattura_sp")
	pagamento=rs1("pagamento")

	spese_bancarie_v=cdbl(rs1("spese_bancarie"))
	spese_bancarie_rs1=spese_bancarie_v
	sconto_totale=0
	tot_fattura=0
	tot_iva_ordini=0
	tot_ordini=0
	tot_ordini_ivati=0
	tot_arrotondamento=0
	sconto_ordine=0
	'sql_ordine="select SQL_CALC_FOUND_ROWS ordini_fatture.idfat, ordini.*, ddt.IDddt, ddt.tipo FROM (ordini INNER JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN ddt ON ordini.idord = ddt.idord "
	'sql_ordine=sql_ordine&" where idfat="&idfat&" order by ordini.totale"

	sql_ordine=get_sql_ordine_fattura(rs1("tipo_fattura"),idfat)

	Set rs_ordini = conn.execute (sql_ordine)
	lngTotalRecords=conn.Execute("Select Found_Rows();")(0).Value

	n_ordini=clng(lngTotalRecords)
	if spese_bancarie_v>0 then spese_bancarie_v=roundup(spese_bancarie_v/n_ordini,2)
	spese_bancarie_cumulativo=0
	conteggio_ordini=1
	'ciclo su ordini
	log_txt=""
	log_txt=log_txt&"Ricalcolo [fattura="&rs1("idfat")&"]"&rs1("nfat")&"[/fattura], tipo_fattura:"&rs1("tipo_fattura")&" <br>"
	if vecchio_conteggio then
		log_txt=log_txt&"Applico vecchio conteggio<br>"
	else
		log_txt=log_txt&"Applico nuovo conteggio<br>"
	end if

	idord_tmp=0

				idtrasportoordine=0

	do while not rs_ordini.eof

		tipo_documento=rs_ordini("tipo_documento")
		tot_ordine=0
		log_txt=log_txt&"Analisi ["&tipo_documento&"="&rs_ordini("idord")&"]"&rs_ordini("nord")&"[/"&tipo_documento&"]<br>"



		if rs_ordini("tipo_documento")="ddt" then
			if rs_ordini("tipo_ddt")="parziale" then
				sql_dettaglio="select ordini_dett.*,ddt_dett_ordini.quantita, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara inner join ddt_dett_ordini on ordini_dett.iddett = ddt_dett_ordini.iddett  where ddt_dett_ordini.idddt="&rs_ordini("idddt")&" order by ordine,iddett;"
			else
				sql_dettaglio="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara  where idord="&rs_ordini("sub_idord")&" order by ordine,iddett;"
			end if
		else
				sql_dettaglio="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara  where idord="&rs_ordini("idord")&" order by ordine,iddett;"
		end if



		on error resume next
		set rs_ordini_dett=conn.execute(sql_dettaglio)
		if err.number<>0 then
			call add2log("Errore query sql_dettaglio:"&sql_dettaglio,0)
			call connclose()
			response.end


		end if
		on error goto 0
		'ciclo su dettaglio ordini
		do while not rs_ordini_dett.eof

			'prezzoeff=rs_ordini_dett("prezzo")
			'prezzoeff=prezzoeff-(rs_ordini_dett("sconto_prodotto")/100)*prezzoeff
			'tot_ordine=tot_ordine+roundup(prezzoeff*rs_ordini_dett("quantita"),2)
            if vecchio_conteggio then
				prezzo=cdbl(rs_ordini_dett("prezzo"))
				sconto_prodotto=cdbl(rs_ordini_dett("sconto_prodotto"))
				m_totale_riga=round((prezzo-(sconto_prodotto/100)*prezzo)*rs_ordini_dett("quantita"),2)
				'tot_ordine=tot_ordine+round((rs_ordini_dett("prezzo")-(rs_ordini_dett("sconto_prodotto")/100)*rs_ordini_dett("prezzo"))*rs_ordini_dett("quantita"),2)
				tot_ordine=tot_ordine+m_totale_riga
	        else
				tot_ordine=tot_ordine+totale_riga(rs_ordini_dett("prezzo"),rs_ordini_dett("sconto_prodotto"),rs_ordini_dett("quantita"))
			end if



			rs_ordini_dett.movenext
		loop

		'Aggiungo il trasporto
		trasporto_ordine=0
		if rs_ordini("tipo_ddt")="parziale" then
			'trasporto ddt parziale


				if idtrasportoordine<>rs_ordini("idord") then
					trasporto_ordine=cdbl(rs_ordini("trasporto"))
					idtrasportoordine=rs_ordini("idord")
					end if

			if cdbl(rs_ordini("trasporto_ddt"))>0 then
				trasporto_ordine=trasporto_ordine+cdbl(rs_ordini("trasporto_ddt"))
			end if
		else
			trasporto_ordine=cdbl(rs_ordini("trasporto"))
		end if
		if idord_tmp<>rs_ordini("idord") then
			sconto_ordine=roundup(rs_ordini("sconto_ordine"),2)
		end if


		tot_ordine=tot_ordine+roundup(trasporto_ordine,2)

		'Conteggio lo sconto


		sconto_totale=sconto_totale+sconto_ordine
		'if isnull(sconto) then sconto=0

		'imposto spese_bancarie_ordine
		tot_fattura=tot_fattura+tot_ordine	'aggiorno cumulativo fattura prima di applicare iva ordine

		if conteggio_ordini=n_ordini then spese_bancarie_v=spese_bancarie_rs1-spese_bancarie_cumulativo

		log_txt=log_txt&"spese bancarie su ordine "&spese_bancarie_v&"<br>"

		if isnull(rs1("trattamento_iva")) then rs1("trattamento_iva")=0


		tot_ordine=tot_ordine+spese_bancarie_v
		log_txt=log_txt&"imponibile ordine "&tot_ordine&"<br>"


		'if vecchio_conteggio then
		'	totord=round(totord,2)
		'else
		'	totord=roundup(totord,2)
		'end if

		tot_ordini=tot_ordini+tot_ordine
		if rs1("trattamento_iva")=10 or fattura_sp then
			'Calcolo iva arrotondata
			log_txt=log_txt&"<b>Calcolo iva per trattamento_iva=10</b><br>"
			valore_iva=tot_ordine*rs_ordini("iva_ordine")/100
			log_txt=log_txt&"imposta ordine "&valore_iva&"<br>"
			arrotondamento=valore_iva-roundup(valore_iva,2)
			tot_arrotondamento=tot_arrotondamento+arrotondamento
			tot_arrotondamento=0
			valore_iva=roundup(valore_iva,2)
			log_txt=log_txt&"imposta ordine dopo arrotondamento "&valore_iva&"<br>"
			log_txt=log_txt&"arrotondamento imposta ordine "&formatnumber(arrotondamento,3)&"<br>"

		elseif esenzione_iva(rs1("trattamento_iva")) or rs1("iva_0") then
			valore_iva=0
			log_txt=log_txt&"Esenzione iva<br>"
		else
			'Calcolo iva arrotondata
			log_txt=log_txt&"<b>Calcolo iva su  tot_ordine: "&tot_ordine&" con iva_ordine:"&rs_ordini("iva_ordine")&"</b><br>"
			valore_iva=tot_ordine*rs_ordini("iva_ordine")/100
			log_txt=log_txt&"imposta ordine "&valore_iva&"<br>"
			arrotondamento=valore_iva-roundup(valore_iva,2)
			tot_arrotondamento=tot_arrotondamento+arrotondamento
			valore_iva=roundup(valore_iva,2)
			log_txt=log_txt&"imposta ordine dopo arrotondamento "&valore_iva&"<br>"
			log_txt=log_txt&"arrotondamento imposta ordine "&formatnumber(arrotondamento,3)&"<br>"

		end if


		if rs1("trattamento_iva")=10 then
			'tot_ordine
		else
			tot_ordine=tot_ordine+valore_iva-sconto
		end if


		tot_iva_ordini=tot_iva_ordini+valore_iva
		tot_ordini_ivati=tot_ordine
		log_txt=log_txt&"totale ordine ivato "&tot_ordine&"<br>"

		tot_ordine=roundup(tot_ordine,2)

		arrotondamento_ordine=0
		if conteggio_ordini=n_ordini then
			if esenzione_iva(rs1("trattamento_iva")) or rs1("iva_0") then
				valore_iva_fattura=0
			else
				valore_iva_fattura=roundup(tot_ordini*iva(rs1("data"))/100,2)
			end if
			if rs1("trattamento_iva")=10 then
				arrotondamento_ordine=0
				tot_arrotondamento=0
			else
			'tot_arrotondamento=roundup(tot_arrotondamento,2)
				tot_arrotondamento=valore_iva_fattura-tot_iva_ordini
				arrotondamento_ordine=tot_arrotondamento
				tot_ordine=tot_ordine+tot_arrotondamento
			end if
		end if

		'Aggiorno totale ordine
		if rs_ordini("tipo_ddt")<>"parziale" then
			call add2log("aggiorno valore_iva="&valore_iva,0)
			Set rs_ordine = Server.CreateObject("ADODB.Recordset")
			rs_ordine.open "select ordini.* from ordini where idord="&rs_ordini("idord"),conn,3,3
			rs_ordine("totale")=tot_ordine
			rs_ordine("imposta_ordine")=valore_iva
			rs_ordine("arrotondamento")=roundup(arrotondamento_ordine,2)
			spese_bancarie_v=roundup(spese_bancarie_v,2)
			rs_ordine("spese_bancarie_ordine")=spese_bancarie_v
			rs_ordine("trattamento_iva_ordine")=rs1("trattamento_iva")
			rs_ordine.Update
			rs_ordine.close
			set rs_ordine = Nothing
		end if

		'Azzero dati trasporto
		trasporto_ordine=0
		idord_tmp=rs_ordini("idord")


		rs_ordini.movenext
		spese_bancarie_cumulativo=spese_bancarie_cumulativo+spese_bancarie_v
		conteggio_ordini=conteggio_ordini+1
	loop
	'tot_fattura=round(tot_fattura,2)
	log_txt=log_txt&"<b>Fine analisi ordini</b><br>"

	log_txt=log_txt&"imponibile totale degli ordini "&tot_ordini&"<br>"
	rs1("totale_merce")=tot_fattura
	if spese_bancarie_rs1>0 then tot_fattura=tot_fattura+spese_bancarie_rs1
	log_txt=log_txt&"imponibile totale della fattura "&tot_fattura&"<br>"
	log_txt=log_txt&"totale imposte degli ordini "&tot_iva_ordini&"<br>"
	log_txt=log_txt&"totale arrotondamenti imposte su ordini "&formatnumber(tot_arrotondamento,2)&"<br>"
	if esenzione_iva(rs1("trattamento_iva")) or rs1("iva_0") then
		valore_iva=0
	else
		valore_iva=roundup(tot_fattura*iva(rs1("data"))/100,2)
		tot_fattura=tot_fattura+valore_iva
	end if
	rs1("imposta")=valore_iva
	rs1("sconto_ordine")=sconto_totale
	log_txt=log_txt&"totale imposte della fattura "&valore_iva&"<br>"
	log_txt=log_txt&"totale fattura "&tot_fattura&"<br>"
	tot_fattura=tot_fattura-sconto_totale
	log_txt=log_txt&"totale fattura dopo sconto "&tot_fattura&"<br>"
	tot_fattura=roundup(tot_fattura,2)
	rs_ordini_dett.close
	rs_ordini.close
	set rs_ordini_dett= nothing
	set rs_ordini= nothing
	rs1("totale_fattura")=tot_fattura
	rs1.update
	rs1.close
	Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class
	n=metodipagamento.genera_scadenze(IDfat,Nfat,0,data,tot_fattura,pagamento)
	set metodipagamento = Nothing


	call add2log (log_txt&" ricreata "&n&" scadenze",1)
	set rs1= nothing
	calcola_totale_fattura=tot_fattura
end function
function calcola_totale_fattura_workline(idfat)
	dim vecchio_conteggio
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="select fatture.* FROM fatture WHERE idfat= " & idfat &";"
	rs1.open sql1,conn,3,3
	if DateDiff("d","01/01/2015",rs1("data"))<0 then 'prima del 01/01/2015
		vecchio_conteggio=true
	else
		vecchio_conteggio=false
	end if

	spese_bancarie_v=cdbl(rs1("spese_bancarie"))
	spese_bancarie_rs1=spese_bancarie_v
	sconto_totale=0
	tot_fattura=0
	tot_iva_ordini=0
	tot_ordini=0
	tot_ordini_ivati=0
	tot_arrotondamento=0


	'sql_ordine="select SQL_CALC_FOUND_ROWS ordini_fatture.idfat, ordini.*, ddt.IDddt FROM (ordini INNER JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN ddt ON ordini.idord = ddt.idord "
	'sql_ordine=sql_ordine&" where idfat="&idfat&" order by ordini.totale"

	sql_ordine=get_sql_ordine_fattura(rs1("tipo_fattura"),idfat)


	Set rs_ordini = conn.execute (sql_ordine)
	lngTotalRecords=conn.Execute("Select Found_Rows();")(0).Value

	n_ordini=clng(lngTotalRecords)
	if spese_bancarie_v>0 then spese_bancarie_v=roundup(spese_bancarie_v/n_ordini,2)
	spese_bancarie_cumulativo=0
	conteggio_ordini=1
	'ciclo su ordini
	log_txt=""
	log_txt=log_txt&"Ricalcolo [fattura="&rs1("idfat")&"]"&rs1("nfat")&"[/fattura]<br>"
	if vecchio_conteggio then
		log_txt=log_txt&"Applico vecchio conteggio<br>"
	else
		log_txt=log_txt&"Applico nuovo conteggio<br>"
	end if

	do while not rs_ordini.eof
        if rs_ordini("tipo_documento")="ddt" then
		        tipo_ddt=rs_ordini("tipo_ddt")
		end if
		tot_ordine=0
		log_txt=log_txt&"Analisi [ordine="&rs_ordini("idord")&"]"&rs_ordini("nord")&"[/ordine]<br>"
		sql_dettaglio="select ordini_dett.* FROM ordini_dett  where idord="&rs_ordini("idord")

			if rs_ordini("tipo_documento")="ddt" then
				if tipo_ddt="parziale" then
					sql_dettaglio="select ordini_dett.*,ddt_dett_ordini.quantita, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara inner join ddt_dett_ordini on ordini_dett.iddett = ddt_dett_ordini.iddett  where ddt_dett_ordini.idddt="&rs_ordini("idddt")&" order by ordine,iddett;"
				else
					sql_dettaglio="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara  where idord="&rs_ordini("sub_idord")&" order by ordine,iddett;"
				end if
			else
					sql_dettaglio="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara  where idord="&rs_ordini("idord")&" order by ordine,iddett;"

			end if



		set rs_ordini_dett=conn.execute(sql_dettaglio)
		'ciclo su dettaglio ordini
		do while not rs_ordini_dett.eof

			'prezzoeff=rs_ordini_dett("prezzo")
			'prezzoeff=prezzoeff-(rs_ordini_dett("sconto_prodotto")/100)*prezzoeff
			'tot_ordine=tot_ordine+roundup(prezzoeff*rs_ordini_dett("quantita"),2)
            if vecchio_conteggio then
				prezzo=cdbl(rs_ordini_dett("prezzo"))
				sconto_prodotto=cdbl(rs_ordini_dett("sconto_prodotto"))
				m_totale_riga=round((prezzo-(sconto_prodotto/100)*prezzo)*rs_ordini_dett("quantita"),2)
				'tot_ordine=tot_ordine+round((rs_ordini_dett("prezzo")-(rs_ordini_dett("sconto_prodotto")/100)*rs_ordini_dett("prezzo"))*rs_ordini_dett("quantita"),2)
				tot_ordine=tot_ordine+m_totale_riga
	        else
				tot_ordine=tot_ordine+totale_riga(rs_ordini_dett("prezzo"),rs_ordini_dett("sconto_prodotto"),rs_ordini_dett("quantita"))
			end if



			rs_ordini_dett.movenext
		loop

		'Aggiungo il trasporto
		tot_ordine=tot_ordine+roundup(rs_ordini("trasporto"),2)

		'Conteggio lo sconto
		sconto=roundup(rs_ordini("sconto_ordine"),2)

		sconto_totale=sconto_totale+sconto
		if isnull(sconto) then sconto=0

		'imposto spese_bancarie_ordine
		tot_fattura=tot_fattura+tot_ordine	'aggiorno cumulativo fattura prima di applicare iva ordine

		if conteggio_ordini=n_ordini then spese_bancarie_v=spese_bancarie_rs1-spese_bancarie_cumulativo

		log_txt=log_txt&"spese bancarie su ordine "&spese_bancarie_v&"<br>"

		if isnull(rs1("trattamento_iva")) then rs1("trattamento_iva")=0


		tot_ordine=tot_ordine+spese_bancarie_v
		log_txt=log_txt&"imponibile ordine "&tot_ordine&"<br>"
		'if vecchio_conteggio then
		'	totord=round(totord,2)
		'else
		'	totord=roundup(totord,2)
		'end if
		tot_ordini=tot_ordini+tot_ordine
		if rs1("trattamento_iva")=10 then
			'Calcolo iva arrotondata
			valore_iva=tot_ordine*rs_ordini("iva_ordine")/100
			log_txt=log_txt&"imposta ordine "&valore_iva&"<br>"
			arrotondamento=valore_iva-roundup(valore_iva,2)
			tot_arrotondamento=tot_arrotondamento+arrotondamento
			valore_iva=roundup(valore_iva,2)
			log_txt=log_txt&"imposta ordine dopo arrotondamento "&valore_iva&"<br>"
			log_txt=log_txt&"arrotondamento imposta ordine "&formatnumber(arrotondamento,3)&"<br>"

		elseif esenzione_iva(rs1("trattamento_iva")) or rs1("iva_0") then
			valore_iva=0
		else
			'Calcolo iva arrotondata
			valore_iva=tot_ordine*rs_ordini("iva_ordine")/100
			log_txt=log_txt&"imposta ordine "&valore_iva&"<br>"
			arrotondamento=valore_iva-roundup(valore_iva,2)
			tot_arrotondamento=tot_arrotondamento+arrotondamento
			valore_iva=roundup(valore_iva,2)
			log_txt=log_txt&"imposta ordine dopo arrotondamento "&valore_iva&"<br>"
			log_txt=log_txt&"arrotondamento imposta ordine "&formatnumber(arrotondamento,3)&"<br>"

		end if

		if rs1("trattamento_iva")=10 then valore_iva=0
		tot_iva_ordini=tot_iva_ordini+valore_iva
		tot_ordine=tot_ordine+valore_iva-sconto
		tot_ordini_ivati=tot_ordine
		log_txt=log_txt&"totale ordine ivato "&tot_ordine&"<br>"

		tot_ordine=roundup(tot_ordine,2)

		arrotondamento_ordine=0
		if conteggio_ordini=n_ordini then
			if esenzione_iva(rs1("trattamento_iva")) or rs1("iva_0") then
				valore_iva_fattura=0
			else
				valore_iva_fattura=roundup(tot_ordini*iva(rs1("data"))/100,2)
			end if

			'tot_arrotondamento=roundup(tot_arrotondamento,2)
			tot_arrotondamento=valore_iva_fattura   -tot_iva_ordini
			arrotondamento_ordine=tot_arrotondamento
			tot_ordine=tot_ordine+tot_arrotondamento
		end if

		'Aggiorno totale ordine
		Set rs_ordine = Server.CreateObject("ADODB.Recordset")
		rs_ordine.open "select ordini.* from ordini where idord="&rs_ordini("idord"),conn,3,3
		rs_ordine("totale")=tot_ordine
		rs_ordine("imposta_ordine")=valore_iva
		rs_ordine("arrotondamento")=roundup(arrotondamento_ordine,2)
		spese_bancarie_v=roundup(spese_bancarie_v,2)
		rs_ordine("spese_bancarie_ordine")=spese_bancarie_v
		rs_ordine("trattamento_iva_ordine")=rs1("trattamento_iva")
		rs_ordine.Update
		rs_ordine.close
		set rs_ordine = Nothing
		rs_ordini.movenext
		spese_bancarie_cumulativo=spese_bancarie_cumulativo+spese_bancarie_v
		conteggio_ordini=conteggio_ordini+1
	loop
	'tot_fattura=round(tot_fattura,2)
	log_txt=log_txt&"<b>Fine analisi ordini</b><br>"

	log_txt=log_txt&"imponibile totale degli ordini "&tot_ordini&"<br>"
	rs1("totale_merce")=tot_fattura
	if spese_bancarie_rs1>0 then tot_fattura=tot_fattura+spese_bancarie_rs1
	log_txt=log_txt&"imponibile totale della fattura "&tot_fattura&"<br>"
	log_txt=log_txt&"totale imposte degli ordini "&tot_iva_ordini&"<br>"
	log_txt=log_txt&"totale arrotondamenti imposte su ordini "&formatnumber(tot_arrotondamento,2)&"<br>"
	if esenzione_iva(rs1("trattamento_iva")) or rs1("iva_0") then
		valore_iva=0
	else
		valore_iva=roundup(tot_fattura*iva(rs1("data"))/100,2)
		tot_fattura=tot_fattura+valore_iva
	end if
	rs1("imposta")=valore_iva
	rs1("sconto_ordine")=sconto_totale
	log_txt=log_txt&"totale imposte della fattura "&valore_iva&"<br>"
	log_txt=log_txt&"totale fattura "&tot_fattura&"<br>"
	tot_fattura=tot_fattura-sconto_totale
	log_txt=log_txt&"totale fattura dopo sconto "&tot_fattura&"<br>"
	tot_fattura=roundup(tot_fattura,2)
	rs_ordini_dett.close
	rs_ordini.close
	set rs_ordini_dett= nothing
	set rs_ordini= nothing
	rs1("totale_fattura")=tot_fattura
	rs1.update
	rs1.close
	if session("iduser")=1 then add2log log_txt,1
	set rs1= nothing
	calcola_totale_fattura=tot_fattura
end function

function max_fatt(pa,fattura_sp,anno)
	dim rs

	if fattura_sp=0 then
		SQL="select Max(fatture.nfat) AS MAX_IDfat FROM fatture WHERE anno="&anno &" and pa="&pa&" and fattura_sp=0 and data > '2023-05-29'"
	else
		SQL="select Max(fatture.nfat) AS MAX_IDfat FROM fatture WHERE anno="&anno &" and fattura_sp=1 and data > '2023-05-29'"
	end if
	set rs=conn.execute(SQL)
	MAX_IDfat=RS("MAX_IDfat")
	if isnull(MAX_IDfat) then MAX_IDfat=0
	max_fatt=MAX_IDfat+1
	set rs=Nothing
end function

function max_ord(tabella,anno)
	dim rs_o
	SQL="select Max("&tabella&".Nord) AS MAX_Nord FROM "&tabella&" WHERE data > '2023-05-29' and anno="&anno&";"
	set rs_o=conn.execute(SQL)
	MAX_Nord=rs_o("MAX_Nord")
	if isnull(MAX_Nord) then MAX_Nord=0
	max_ord=MAX_Nord+1
	rs_o.close
	set rs_o=nothing
end function
sub file_bollettino(idord)
		'response.ContentType="application/vnd.ms-excel"
		Dim objFSO, objTextFile
		'Creazione dell'istanza filesystem
		Set objFSO = CreateObject("Scripting.filesystemObject")
		'Apertura del file di testo
		Set objTextFile = objFSO.CreateTextFile(Server.MapPath("\public\tracciato.txt"))

		sql_ordine="select ordini.*, ddt.*,ddt.data as ddt_data, ddt.peso as ddt_peso, ddt.colli as ddt_colli, ddt.porto as ddt_porto, utenti.email, utenti.Cellulare FROM (ordini LEFT JOIN ddt ON ordini.idord = ddt.idord) INNER JOIN utenti ON ordini.iduser = utenti.iduser "
		sql_ordine=sql_ordine&" where ordini.idord="&idord&" "
		set rs_ordini=conn.execute(sql_ordine)
		'Scrivo la prima riga di codice
		separatore=";"
		riga=""
		'response.write len(trim(rs_ordini("azienda")))&"|"& rs_ordini("nominativo")&"|"
		'add2log "Bollettino ordine "& idord&", lunghezza d_indirizzo " & len(rs_ordini("d_indirizzo"))&isnull(rs_ordini("d_indirizzo"))&".",2
		lunghezza_campo=35
		if len(rs_ordini("d_indirizzo"))<3 or isnull(rs_ordini("d_indirizzo")) then
		'add2log "Bollettino ordine "& idord&", indirizzo principale",2
			'Ragione sociale
			if len(trim(rs_ordini("azienda")))>3 then
			riga_tmp=left(rs_ordini("azienda")&" ",lunghezza_campo)
			else
			riga_tmp=left(rs_ordini("cognome")&" "&rs_ordini("nome"),lunghezza_campo)
			end if
			riga=riga&riga_tmp&separatore
			'Indirizzo
			lunghezza_campo=35
			riga_tmp=left(rs_ordini("indirizzo"),lunghezza_campo)
			riga=riga&riga_tmp&separatore
			'Localita
			lunghezza_campo=30
			riga_tmp=left(rs_ordini("citta"),lunghezza_campo)
			riga=riga&riga_tmp&separatore
			'CAP
			lunghezza_campo=5
			riga_tmp=left(rs_ordini("cap"),lunghezza_campo)
			riga=riga&riga_tmp&separatore
			'Provincia
			lunghezza_campo=2
			riga_tmp=left(rs_ordini("provincia"),lunghezza_campo)
			riga=riga&riga_tmp&separatore
		else
		'add2log "Bollettino ordine "& idord&", indirizzo alternativo",2

			if len(trim(rs_ordini("d_azienda")))>3 then
			riga_tmp=left(rs_ordini("d_azienda")&" ",lunghezza_campo)
			else
			riga_tmp=left(rs_ordini("cognome")&" "&rs_ordini("nome"),lunghezza_campo)
			end if
			riga=riga&riga_tmp&separatore
			'Indirizzo
			lunghezza_campo=35
			riga_tmp=left(rs_ordini("d_indirizzo"),lunghezza_campo)
			riga=riga&riga_tmp&separatore
			'Localita
			lunghezza_campo=30
			riga_tmp=left(rs_ordini("d_citta"),lunghezza_campo)
			riga=riga&riga_tmp&separatore
			'CAP
			lunghezza_campo=5
			riga_tmp=left(rs_ordini("d_cap"),lunghezza_campo)
			riga=riga&riga_tmp&separatore
			'Provincia
			lunghezza_campo=2
			riga_tmp=left(rs_ordini("d_provincia"),lunghezza_campo)
			riga=riga&riga_tmp&separatore
		end if



		'DDT
		lunghezza_campo=10
		if request("colli")<>"" then
				peso=request.querystring("peso")
				colli=request.querystring("colli")
				porto=cint(request.querystring("porto"))
				tipo_pagamento=rs_ordini("tipopagamento")



		elseif isnull(rs_ordini("nddt")) then
			sql="select ordini_fatture.idord, fatture.nfat,fatture.data, fatture.colli, fatture.peso,fatture.porto,fatture.pagamento FROM ordini_fatture INNER JOIN fatture ON ordini_fatture.idfat = fatture.IDfat where ordini_fatture.idord="&idord
			'response.write sql
			set rs1=conn.execute (sql)
			if not rs1.eof then
				'Dati da fattura
				datida="Fattura"
				testo=pre_fattura(rs1("data"))&rs1("nfat")
				tabella="fatture"
				peso=rs1("peso")
				colli=rs1("colli")
				porto=rs1("porto")
				tipo_pagamento=rs1("pagamento")
			else
				'Dati da ordine
				datida="Ordine"
				testo=""
				tabella="ordini"
				peso=rs_ordini("peso")
				colli=rs_ordini("colli")
				porto=rs_ordini("porto")
				tipo_pagamento=rs_ordini("tipopagamento")
			end if
			rs1.close
		else
			'Dati da DDT
			datida="DDT"
			testo=rs_ordini("nddt")
			tabella="ordini"
			peso=rs_ordini("ddt_peso")
			colli=rs_ordini("ddt_colli")
			porto=rs_ordini("ddt_porto")
			tipo_pagamento=rs_ordini("tipopagamento")
		end if
		'add2log "Datida:"&datida&" Porto:"&porto&" Tipo pagamento:"&tipo_pagamento,1
		riga_tmp=left(testo,lunghezza_campo)
		riga=riga&riga_tmp&separatore

		'Data DDT
		lunghezza_campo=2
		riga_tmp=left(day(rs_ordini("ddt_data")),lunghezza_campo)
		'riga=riga&riga_tmp
		lunghezza_campo=2
		riga_tmp=left(month(rs_ordini("ddt_data")),lunghezza_campo)
		'riga=riga&riga_tmp
		lunghezza_campo=2
		riga_tmp=right(year(rs_ordini("ddt_data")),lunghezza_campo)
		riga=riga&separatore

		'colli
		lunghezza_campo=5
		riga_tmp=left(colli,lunghezza_campo)
		riga=riga&riga_tmp&separatore

		riga=riga&separatore 'incoterm

		'Peso
		lunghezza_campo=6
		riga_tmp=left(peso,lunghezza_campo)
		riga=riga&riga_tmp&separatore
		'Contrassegno
		riga_tmp=""
		if tipo_pagamento=3 then
			riga_tmp=formatnumber(rs_ordini("totale"),2) 'Importo contrassegno
		else
			riga_tmp=""
		end if
		riga=riga&riga_tmp&separatore 'Importo contrassegno

		riga=riga&separatore 'Note spedizione

		'Porto
		lunghezza_campo=1
		riga_tmp=""
		if porto=1 or porto=3 then riga_tmp="F"
		if porto=2 then riga_tmp="A"
		riga=riga&riga_tmp&separatore

		riga=riga&separatore 'Fermo deposito
		riga=riga&separatore 'Importo assicurazione
		riga=riga&separatore 'Peso volume
		riga=riga&separatore 'Tipo collo
		riga=riga&separatore 'Franco anticipata
		riga=riga&separatore 'Riferimenti cliente
		riga=riga&separatore 'Note aggiuntive

		'Codice cliente
		lunghezza_campo=30
		riga_tmp=left(rs_ordini("iduser"),lunghezza_campo)
		riga=riga&separatore

		riga=riga&separatore 'Valore merce
		riga=riga&separatore 'Primo ld collo
		riga=riga&separatore 'Ultimo ld collo

		'Email per notifica
		lunghezza_campo=70
		riga_tmp=left(rs_ordini("email"),lunghezza_campo)
		riga=riga&riga_tmp&separatore

		'Cellulare notifica
		lunghezza_campo=10
		riga_tmp=left(rs_ordini("cellulare"),lunghezza_campo)
		riga_tmp=""
		riga=riga&riga_tmp&separatore
		riga=riga&separatore 'Cellulare notifica
		riga=riga&separatore 'Servizi accessori
		if isnumeric(rs_ordini("tipopagamento")) then
			if rs_ordini("tipopagamento")=3 then
				riga_tmp=left(rs_ordini("pagamentobollettino"),lunghezza_campo)
			end if
		else
			riga_tmp=""
		end if
		riga=riga&riga_tmp&separatore 'Modalita incasso
		riga=riga&separatore 'Data
		riga=riga&separatore 'Note e orario


		'Scrivo la riga
		objTextFile.Write riga

		'Chiudo il file e i vari oggetti/istanze
		objTextFile.Close
		Set objTextFile = Nothing
		Set objFSO = Nothing

		if true then
		Response.Buffer = True
		Response.Clear

		Dim strFilePath, strFileName

		Const adTypeBinary = 1

		Set objStream = Server.CreateObject("ADODB.Stream")

		strFilePath = Server.MapPath("\public\tracciato.txt")

		objStream.Open
		objStream.Type = adTypeBinary
		objStream.LoadFromFile strFilePath

		Response.AddHeader "Content-Disposition", "attachment; filename=" & "bollettino.txt"
		Response.Charset = "UTF-8"
		Response.ContentType = "text/plain"
		'Google "Mime Types" for addtional Content Type definitions

		Response.BinaryWrite objStream.Read
		Response.flush
		objStream.Close
		Set objStream = Nothing
		else
		response.write riga
		end if
end sub
function vedi_voto(stato)
		vedi_voto=""
		select case stato
		case ""
		case "3"
		vedi_voto="<img src='images/alert.png' border='0'>"
		end select
end function


sub email_ordine(idord,cosa, tabella)
	Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class

	dim timertmp,timertmp2
	timertmp=timer()
	timertmp2=timertmp
	if cosa="ordine" then tipo_checkout="ordini"
	if cosa="preventivo" then tipo_checkout="preventivi"
	sqlm="select "&tipo_checkout&".*, utenti.trattamento_iva, utenti.email, utenti.telefono, utenti.fax, utenti.cellulare FROM "&tipo_checkout&" INNER JOIN utenti ON "&tipo_checkout&".iduser = utenti.iduser  where idord="&idord &";"

	Set rs_ordine = conn.execute (sqlm)
	if rs_ordine.eof=false then

		sql="select impostazioni.dati_trasporto FROM impostazioni"
		set rst=conn.execute(sql)
		dati_trasporto=split(rst("dati_trasporto"),vbcrlf)
		rst.close
		set rst=nothing
		stato=rs_ordine("stato")
		txtHTML=txtHTML&"<center><table border='0' cellpadding='0' cellspacing='0' style='border: 1px solid black;' width='750'>"
		txtHTML=txtHTML&"<tr><td align='left'>Riferimento "&cosa&": " & rs_ordine("nord")&"</td></tr>"
		txtHTML=txtHTML&"<tr><td align='left'>Data: " & formatdatetime(rs_ordine("data"),2)&"</td></tr>"
		if cosa="ordine" then txtHTML=txtHTML&"<tr><td align='left'>Stato ordine: " & stato_ordine(stato)&"</td></tr>"
		'txtHTML=txtHTML&"<tr><td align='left'><b>Comunicazioni per il tuo ordine:</b></td></tr>"
		txtHTML=txtHTML&"<tr><td align='left'><p>" & rs_ordine("conferma_ordine")&"</p></td></tr>"
		txtHTML=txtHTML&"</table>"
		txtHTML=txtHTML&"<br>"

		if stato=3 or (stato=6 and rs_ordine("creato_da_admin")=0) or tabella then
				txtHTML=txtHTML&"<table border='0' cellpadding='0' cellspacing='0' style='border: 1px solid black;' width='750'>"
				txtHTML=txtHTML&"<tr><td colspan='2' align='left'><b>Dati anagrafici:</b></td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Email:</td><td>" & rs_ordine("email")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Cognome:</td><td>" & rs_ordine("cognome")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Nome:</td><td>" & rs_ordine("nome")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Azienda:</td><td>" & rs_ordine("azienda")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Indirizzo:</td><td>" & rs_ordine("indirizzo")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Citta':</td><td>" & rs_ordine("citta")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Cap:</td><td>" & rs_ordine("cap")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Provincia:</td><td>" & rs_ordine("provincia")&"</td></tr>"&vbcrlf
				'txtHTML=txtHTML&"<tr><td>Regione:</td><td>" & regione(rs_ordine("regione"))&"</td></tr>"
				txtHTML=txtHTML&"<tr><td>Telefono:</td><td>" & rs_ordine("telefono")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Fax:</td><td>" & rs_ordine("fax")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Cellulare:</td><td>" & rs_ordine("cellulare")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Partita Iva:</td><td>" & rs_ordine("piva")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Codice Fiscale:</td><td>" & rs_ordine("cf")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td colspan='2' align='left'><b>Consegna della merce:</b></td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Azienda:</td><td>" & rs_ordine("D_azienda")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Indirizzo:</td><td>" & rs_ordine("D_indirizzo")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Citta':</td><td>" & rs_ordine("D_citta")&"</td></tr>"&vbcrlf
				'if rs_ordine("d_regione")>1 then txtHTML=txtHTML&"<tr><td>Regione:</td><td>" & regione(rs_ordine("d_regione"))&"</td></tr>"
				txtHTML=txtHTML&"<tr><td>Provincia:</td><td>" & rs_ordine("D_provincia")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Cap:</td><td>" & rs_ordine("D_cap")&"</td></tr>"&vbcrlf
				'Altri dati
				txtHTML=txtHTML&"<tr><td colspan='2' align='left'><b>Altri dati:</b></td></tr>"
				txtHTML=txtHTML&"<tr><td>Tipo pagamento:</td><td>" & metodipagamento.descrizione(rs_ordine("tipopagamento"))&"</td></tr>"
				txtHTML=txtHTML&"<tr><td>Tipo trasporto:</td><td>" & tipo_trasporto(rs_ordine("tipo_trasporto"))&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Note trasporto:</td><td>" & rs_ordine("note")&"</td></tr>"&vbcrlf
				txtHTML=txtHTML&"<tr><td>Note sui prodotti acquistati:</td><td> " & rs_ordine("noteacq")&"</td></tr>"&vbcrlf
				'txtHTML=txtHTML&"<tr><td>Riferimento ordine:</td><td>" & rs_ordine("idord")&"</td></tr>"
				txtHTML=txtHTML&"</table>"
				txtHTML=txtHTML&"<br>"
				txtHTML=txtHTML&"<table border='0' cellpadding='0' cellspacing='0' style='border: 1px solid black;' width='750'>"
				if cosa="ordine" then txtHTML=txtHTML&"<tr><td colspan='6'><b>Articoli ordinati:</b></td></tr>"
				txtHTML=txtHTML&"<tr><td style='border-bottom: 1px solid black;'><b>Codice</b></td><td style='border-bottom: 1px solid black;'><b>Articolo</b></td><td align='center' style='border-bottom: 1px solid black;'><b>Quantit&agrave;</b></td>"
				txtHTML=txtHTML&"<td align='center' style='border-bottom: 1px solid black;'><b>Prezzo</b></td><td align='center' style='border-bottom: 1px solid black;'><b>Sconto</b></td><td align='right' style='border-bottom: 1px solid black;'><b>Totale</b></td></tr>"
				sqlm="select * FROM "&tipo_checkout&"_dett "
				sqlbase="select xxx.*, magazzino.quantita_magazzino, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, varianti_a.variante_a, varianti_b.variante_b FRom ((magazzino RIGHT JoIN xxx oN (magazzino.idvarb = xxx.idvarb) aND (magazzino.idvara = xxx.idvara) aND (magazzino.idpro = xxx.idpro)) LEFT JoIN varianti_b oN xxx.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN xxx.idvara = varianti_a.IDvara where idord="&idord
				sqlm=replace(sqlbase,"xxx",tipo_checkout&"_dett")




				set rsm=conn.execute(sqlm)
				'Ciclo articoli
				totord=0
				do until rsm.eof
					if rsm("modificato") then
						rosso= " color:red;"
					else
						rosso=""
					end if
					txtHTML=txtHTML&"<tr valign='middle'><td style='border-bottom: 1px solid black;"&rosso&"'>&nbsp;<b>"&rsm("codice_ordine")&"</b>&nbsp;</td>"
					txtHTML=txtHTML&"<td style='border-bottom: 1px solid black;"&rosso&"'>&nbsp;<b>"&rsm("articolo_ordine")&"</b>&nbsp;<br>"
					txt=""

						if rsm("var1")<>"" then
							txt ="&nbsp;" & rsm("variante1_ordine") & ": " & rsm("var1")
						end if
						if rsm("var2")<>"" then
							txt =txt & "&nbsp;" & rsm("variante2_ordine") & ": " & rsm("var2")
						end if
					txtHTML=txtHTML& txt
					txtHTML=txtHTML&"</td>"
					'Quantità con unità di misura
					txtHTML=txtHTML&"<td align='center' valign='middle' style='border-bottom: 1px solid black;'>"&rsm("um")&"&nbsp;<b>"&rsm("quantita")&"</b></td>"&vbcrlf

					'Prezzo
					prezzo=cdbl(rsm("prezzo"))
					txtHTML=txtHTML&"<td align='right' style='border-bottom: 1px solid black;'>"&simbolo_valuta&formatnumber(prezzo,2) &"</td>"
					sconto_prodotto=cdbl(rsm("sconto_prodotto"))
					if sconto_prodotto>0 then
						prezzo=prezzo-(sconto_prodotto/100)*prezzo
						sconto=formatnumber(sconto_prodotto,2)&"%"
					else
						sconto="&nbsp;"
					end if
					txtHTML=txtHTML&"<td align='center' style='border-bottom: 1px solid black;'>"&sconto &"</td>"

					txtHTML=txtHTML&"<td align='right' style='border-bottom: 1px solid black;'>"&simbolo_valuta&formatnumber(rsm("totale_riga"),2) &"</td>"
					'totord=totord+(prezzo*rsm("quantita"))
					txtHTML=txtHTML&"</tr>"
					rsm.movenext


				loop
				rsm.close
				set rsm=nothing
				txtHTML=txtHTML&"<tr valign='middle'><td colspan='6' style='border-bottom: 1px solid black;'><b>Totali:</b></td></tr>"&vbcrlf
				trasporto=cdbl(rs_ordine("trasporto"))
				txtHTML=txtHTML&"<tr valign='middle' ><td colspan='6' align='right'>Spese di trasporto: "&simbolo_valuta&formatnumber(trasporto,2)&"</td></tr>"&vbcrlf
				totale_merce_ordine=cdbl(rs_ordine("totale_merce_ordine"))
				imponibile_ordine=totale_merce_ordine+trasporto
				txtHTML=txtHTML&"<tr valign='middle' ><td colspan='6' align='right'>Totale Iva esclusa: "&simbolo_valuta&formatnumber(imponibile_ordine,2)&"</td></tr>"&vbcrlf

				if esenzione_iva(rs_ordine("trattamento_iva")) then
					valore_iva=0
				else
					valore_iva=totord*iva(rs_ordine("data"))/100
					totord=totord*(1+iva(rs_ordine("data"))/100)
				end if
				testo_esenzione=" "&trattamento_iva(rs_ordine("trattamento_iva"))

				imposta_ordine=cdbl(rs_ordine("imposta_ordine"))
				txtHTML=txtHTML&"<tr valign='middle' align='right'><td colspan='6'>IVA "&iva(rs_ordine("data"))&"%  "&testo_esenzione&":"&formatcurrency(imposta_ordine,2)&"</td></tr>"&vbcrlf
				totale_ordine=imponibile_ordine+imposta_ordine
				txtHTML=txtHTML&"<tr valign='middle' align='right'><td colspan='6' >Totale IVA inclusa: "&formatcurrency(totale_ordine,2)&"</td></tr>"&vbcrlf
				'Se c'è applico lo sconto
				sconto=cdbl(rs_ordine("sconto_ordine"))
				if sconto>0 then

					txtHTML=txtHTML&"<tr valign='middle' align='right'><td colspan='6'>Sconto: "&simbolo_valuta&formatnumber(sconto,2)&"</td></tr>"&vbcrlf
					totale_ordine=totale_ordine-sconto
				end if
				'Totale
				txtHTML=txtHTML&"<tr valign='right' ><td colspan='6' align='right'><b>Totale ordine: "&simbolo_valuta&formatnumber(totale_ordine,2)&"</b></td></tr>"&vbcrlf
				txtHTML=txtHTML&"</table>"
		end if
		toemail=rs_ordine("email")
		timerstr=timerstr&"----creazione del testo: "&formatnumber(timer()-timertmp,3)&" - "&time()&"<br>"
		timertmp=timer()
		Const cdoSendUsingMethod        = "http://schemas.microsoft.com/cdo/configuration/sendusing"
		Const cdoSendUsingPort          = 2
		Const cdoSMTPServer             = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
		Const cdoSMTPServerPort         = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
		Const cdoSMTPConnectionTimeout  = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
		Const cdoSMTPAuthenticate       = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
		Const cdoBasic                  = 1
		Const cdoSendUserName           = "http://schemas.microsoft.com/cdo/configuration/sendusername"
		Const cdoSendPassword           = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
		Dim objConfig
		Dim objMessage
		Dim Campi
		Set objConfig = Server.CreateObject("CDO.Configuration")
		Set Campi = objConfig.Fields
		With Campi
		.Item(cdoSendUsingMethod)       = cdoSendUsingPort
		.Item(cdoSMTPServer)            = SMTPServer
		.Item(cdoSMTPServerPort)        = 25
		.Item(cdoSMTPConnectionTimeout) = 30
		.Update
		End With
		timerstr=timerstr&"----Configurazione email: "&formatnumber(timer()-timertmp,3)&" - "&time()&"<br>"
		timertmp=timer()

	'EMAIL AL CLIENTE
	Set objMessage = Server.CreateObject("CDO.Message")
	Set objMessage.Configuration = objConfig
			objMessage.Bodypart.Charset = "UTF-8"
			HTML = "<html><head><title>"&nomesito&"</title><meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1""></head>"&vbcrlf
			HTML = HTML & "</head><body style='text-align:center; width:600px;  margin-left: auto; margin-right: auto; font-family:Verdana, Arial;'>"
			if Application("img_email")<>"" then HTML =HTML&  "<img src='http://"&nomesito&Application("img_email")&"'><br>"
			if cosa="ordine" then
				sql="select * from news where  tipo='E2' order by ordine;"
			else
				sql="select * from news where  tipo='E3' order by ordine;"
			end if
			set rsb=conn.execute(sql)
			if not rsb.eof then
				HTML = HTML &"<br>" & rsb("testo") & "<br>"
			else
				add2log "Manca testo per email conferma "&cosa,2
			end if
			HTML = HTML & "<br>"
			HTML = HTML & txtHTML
			'Aggiunge pagamento PAYPAL

			if instr(lcase(metodipagamento.descrizione(rs_ordine("tipopagamento"))),"paypal")>0 and Application("codice_paypal")<>"" and cosa="ordine" and rs_ordine("stato")>=3 then
			  'HTML= HTML &"<p><a href=""http://" & lcase(nomesito) & "/order.asp?idord="&idord&"""><img src=""https://www.paypalobjects.com/it_IT/i/btn/x-click-but6.gif"" border=""0"" alt=""Effettua i tuoi pagamenti con PayPal. È un sistema rapido, gratuito e sicuro.""></a></p>"
			  HTML= HTML &"<p><a href=""http://" & lcase(nomesito) & "/orders-list.asp""><img src=""https://www.paypalobjects.com/it_IT/i/btn/x-click-but6.gif"" border=""0"" alt=""Effettua i tuoi pagamenti con PayPal. È un sistema rapido, gratuito e sicuro.""></a></p>"
		   end if
			if stato=3 then
				rsb.movenext
				if not rsb.eof then
					HTML = HTML &"<br>" & rsb("testo") & "<br>"
				else
					add2log "Manca testo coda per email conferma "&cosa,2
				end if
			end if
			rsb.close
			set rsb= nothing
			nord=rs_ordine("nord")
			HTML = HTML & "</body>"
			HTML = HTML & "</html>"
			objMessage.From = nomesito  & "<" & Application("email_sito") & ">"
			objMessage.To = toemail
			if stato=0 then
				objMessage.Subject = "Conferma ricevimento "&cosa&" " & nord & " su " & nomesito
			else
				objMessage.Subject = "Aggiornamento "& cosa&" " & nord & " su " & nomesito
			end if
			objMessage.HtmlBody = HTML
			on error resume next
			if toemail<>"" then objMessage.Send
			txtreport="email inviata"
			If err.number <> 0 Then
				if session("report")<>"" then session("report")=session("report")&"<hr>"
				session("report")=session("report")&"<span style=""color: red;""><strong>Errore nell'invio della mail per l'ordine "&nord&" all'indirizzo "&toemail&"</strong></span>"
				txtreport="<span style=""color: red;""><strong>Errore nell'invio della mail all'indirizzo "&toemail&"</strong></span>"
				add2log "Errore nell'invio della mail per l'ordine "&nord&" all'indirizzo "&toemail&"</strong>",1
			End If

			set objMessage = Nothing

			Set rs_ordine2 = Server.CreateObject("ADODB.Recordset")
			sql="select "&tipo_checkout&".* FROM "&tipo_checkout&" where idord="&idord &";"
			rs_ordine2.Open sql, conn, 1, 3
			rs_ordine2("comunicazioni_precedenti")=rs_ordine2("comunicazioni_precedenti")&"<hr>"&now()&" <b>"&session("nominativo")&":</b> "&txtreport&"<br>Messaggio: "&request.form("conferma_ordine")
			rs_ordine2.update
			rs_ordine2.close
			set rs_ordine2 = Nothing

			timerstr=timerstr&"----Inoltro al cliente: "&formatnumber(timer()-timertmp,3)&" - "&time()&"<br>"
			timertmp=timer()

	'EMAIL AL GESTORE
			Set objMessage = Server.CreateObject("CDO.Message")
			Set objMessage.Configuration = objConfig
			objMessage.Bodypart.Charset = "UTF-8"

			objMessage.From = nomesito  & "<" & Application("email_sito") & ">"
			objMessage.To = Application("email_suppl")&";"
			if stato=0 then

				objMessage.Subject = "Ricevuto "&cosa&" " & nord & " su " & nomesito
			else
			objMessage.Subject = "Inviato aggiornamento "& cosa &" "& nord & " da " & nomesito
			end if
			objMessage.HtmlBody = HTML
			objMessage.Send
			set objMessage = Nothing
			timerstr=timerstr&"----Inoltro al gestore: "&formatnumber(timer()-timertmp,3)&" - "&time()&"<br>"
			timerstr=timerstr&"--Esco da email_ordine, totale: "&formatnumber(timer()-timertmp2,3)&" - "&time()&"<br>"

			timertmp=timer()
			err.clear
			on error goto 0
	else
		error=cosa&" "&idord&" non trovato."
	end if
	set rs_ordine=nothing
end sub

sub imposta_verde(idord,verde)
	'call add2log ("imposta verde idord:"&idord&" verde:"&verde,0)
	conn.execute ("UPDATE ordini SET verde = "&converti_bool(verde)&" where idord="&idord)
end sub
function cancellare_cancellare_importo_trasporto(totale_merce)
	dim dati_trasporto
	dati_trasporto=split(application("dati_trasporto"),vbcrlf)
	'response.write "parametri trasporto 2:"&clng(dati_trasporto(2))&"  4:"&clng(dati_trasporto(4))
	if totale_merce<clng(dati_trasporto(2)) then
		importo_trasporto=clng(dati_trasporto(1))
	elseif totale_merce<clng(dati_trasporto(4)) then
		importo_trasporto=totale_merce*(clng(dati_trasporto(3))/100)
	else
		importo_trasporto=0
	end if
	importo_trasporto=roundup(importo_trasporto,2)
end function
function calcola_importo_trasporto(tipotrasporto,totale_merce)
	if tipotrasporto<=2 then
		dim dati_trasporto
		dati_trasporto=split(application("dati_trasporto"),vbcrlf)
		'response.write "parametri trasporto 2:"&clng(dati_trasporto(2))&"  4:"&clng(dati_trasporto(4))
		if totale_merce<clng(dati_trasporto(2)) then
			calcola_importo_trasporto=clng(dati_trasporto(1))
		elseif totale_merce<clng(dati_trasporto(6)) then
			calcola_importo_trasporto=clng(dati_trasporto(5))
		elseif totale_merce<clng(dati_trasporto(8)) then
			calcola_importo_trasporto=clng(dati_trasporto(7))

		elseif totale_merce<clng(dati_trasporto(4)) then
			calcola_importo_trasporto=totale_merce*(clng(dati_trasporto(3))/100)
		else
			calcola_importo_trasporto=0
		end if
		calcola_importo_trasporto=roundup(calcola_importo_trasporto,2)
	else
		calcola_importo_trasporto=0
	end if
end function

function crea_ordine_fornitore(idfor)
	nord=max_ord("ordini_fornitori",year(date()))

	Set rs_for = Server.CreateObject("ADODB.Recordset")
	sql="select * from ordini_fornitori where stato<=1 and eliminato=0 and iduser="&idfor&" order by idord desc"
	rs_for.Open sql, conn,3,3
	if rs_for.eof then
		sql="select * FROM utenti where iduser="&idfor
		set rs_utenti=conn.execute(SQL)
		rs_for.addnew
		rs_for("data")=now()

		'intestazione ordine
		rs_for("eliminato")=false
		rs_for("nord")=nord
		rs_for("iduser")=idfor
		rs_for("cognome")=rs_utenti("cognome")
		rs_for("nome")=rs_utenti("nome")
		rs_for("azienda")=rs_utenti("azienda")
		rs_for("indirizzo")=rs_utenti("indirizzo")
		rs_for("citta")=rs_utenti("citta")
		rs_for("cap")=rs_utenti("cap")
		rs_for("provincia")=rs_utenti("provincia")
		rs_for("regione")=rs_utenti("regione")
		rs_for("piva")=rs_utenti("piva")
		rs_for("cf")=rs_utenti("cf")
		rs_for("D_azienda")=rs_utenti("D_azienda")
		rs_for("D_indirizzo")=rs_utenti("D_indirizzo")
		rs_for("D_citta")=rs_utenti("D_citta")
		rs_for("D_regione")=rs_utenti("D_regione")
		rs_for("D_provincia")=rs_utenti("D_provincia")
		rs_for("D_cap")=rs_utenti("D_cap")
		denominazione_v=denominazione(rs_utenti("nome"),rs_utenti("cognome"),rs_utenti("azienda"))
		'fine intestazione ordine
		rs_for("stato")="0"
		comunicazioni_precedenti= now()&" <b>"&session("nominativo")&":</b> ordine creato da procedura automatica "
		rs_for("comunicazioni_precedenti")= comunicazioni_precedenti
		rs_for("creato_da_admin")=sessionIDUser
		rs_for.update
		crea_ordine_fornitore=Get_last_id("ordini_fornitori")
		session("nord_fornitore")=rs_for("nord")
		session("proc_ordina")=session("proc_ordina")&"creato nuovo [ordine_fornitore="&crea_ordine_fornitore&"]"&nord&"[/ordine_fornitore] "&denominazione_v&"<br> "
	else
		crea_ordine_fornitore=rs_for("idord")
		session("proc_ordina")=session("proc_ordina")&"trovato [ordine_fornitore="&rs_for("idord")&"], "&denominazione_v&"<br>"
		session("nord_fornitore")=rs_for("nord")
	end if

	rs_for.close

	set rs_for=nothing
	set rs_utenti=nothing
end function

function tabella_singolare(tabella)
	select case tabella
	case "ordini"
		tabella_singolare="ordine"
	case "preventivi"
		tabella_singolare="preventivo"
	case "ordini_fornitori"
		tabella_singolare="ordine fornitore"
	end select
end function
function tabella_cosa(tabella)
	select case tabella
	case "ordini"
		tabella_cosa="ordine"
	case "preventivi"
		tabella_cosa="preventivo"
	case "ordini_fornitori"
		tabella_cosa="ordine_fornitore"
	end select
end function

function crea_fattura(tipo)
	dim fattura_pa,rs_tmp
	call add2log(queryeform(),0)
	Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class
	idord=request.form("idord")
	'verifico se esiste già una fattura
	sql="select ordini_fatture.idord FROM ordini_fatture WHERE (((ordini_fatture.idord)="&idord&"));"
	set rs_tmp=conn.execute(SQL)
	rs_eof=rs_tmp.eof
	set rs_tmp=nothing
	if not rs_eof then
		'esiste già
		crea_fattura="-1"
	else
		sql="select ordini.* FROM ordini where ordini.idord="&idord
		Set rs_ordini = Server.CreateObject("ADODB.Recordset")
		rs_ordini.Open sql, conn, 3, 3
		nord=rs_ordini("nord")
		if rs_ordini("tipo_documento")="ddt" then
			call add2log("Crea fattura per ddt idddt:"&idord&" nddt"&nord,0)
			fatturaper="[ddt="&idord&"]"&nord&"[/ddt]"
			tipo_fattura="ddt"
		else
			fatturaper="[ordine="&idord&"]"&nord&"[/ordine]"
			tipo_fattura="ordine"
		end if

		if request.form("fattura_pa")="" then


			if rs_ordini("trattamento_iva_ordine")=10 and  DateDiff("d","31/03/2015",date())>0 then 'dopo il 31/03/2015 then
				fattura_pa=1
			else
				fattura_pa=0
			end if
			tmp_trattamento_iva=rs_ordini("trattamento_iva_ordine")

		else
			fattura_pa=request.form("fattura_pa")

		end if
		if fattura_pa=1 then
			tmp_trattamento_iva=10
		else
	 		tmp_trattamento_iva=rs_ordini("trattamento_iva_ordine")
		end if
		call add2log("form fattura_pa:"&request.form("fattura_pa")&queryeform(),0)
		anno=request.form("anno")
		if anno="" then
			data=date()
			anno=year(data)
		else
			'Anno impostato quindi precedente
			'Cerco la data
			data=conn.execute("select max(data) from fatture where anno="&anno&" and pa="&fattura_pa)(0)
		end if
		Nfat=max_fatt(fattura_pa,fattura_sp,anno)

		if request("piva_editabile")<>"NO" then
			rs_ordini("piva")=ucase(request.form("piva"))
			rs_ordini("cf")=ucase(request.form("cf"))
		end if
		sql="select * from utenti where iduser="&rs_ordini("iduser")
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		if request.form("aggiorna_dati_utente")<>"" then
			rs("pag_accordato")=request.form("pagamento")
			rs("iban")=ucase(request.form("iban"))
			rs("banca_appoggio")=ucase(request.form("banca_appoggio"))
			if request("piva_editabile")<>"NO" then
				rs("piva")=ucase(request.form("piva"))
				rs("cf")=ucase(request.form("cf"))
			end if
			rs.update
		end if
		spese_0=rs("spese_0")
		'cerco banca predefinita
		idbanca=rs("idbanca")
		if idbanca="" or idbanca=0 then idbanca=Application("banca")
		rs.close
		'Aggiungo fattura
		sql="select fatture.* FROM fatture"
		rs.Open sql, conn, 1, 3
		rs.addnew
		RS("nfat")=Nfat
		rs("data")=data
		rs("tipo_fattura")=tipo_fattura
		rs("anno")=year(data)
		rs("iduser")=rs_ordini("iduser")
		rs("idintestazione")=rs_ordini("idintestazione")
		rs("idconsegna")=rs_ordini("idconsegna")
		rs("idord")=idord
		rs("annotazioni")=request.form("annotazioni")
		rs("data_inizio")=request.form("data_inizio")
		rs("incaricato")=request.form("incaricato")
		rs("pagamento")=request.form("pagamento")
		if request.form("CodiceDestinatario")<>"" then rs("CodiceDestinatario")=request.form("CodiceDestinatario")
		if spese_0 or isnull(rs("pagamento")) or cdbl(rs_ordini("totale"))<0 then
			spese_bancarie_v=0
		elseif request.form("spese_bancarie")<>"" then
			spese_bancarie_v=aggiusta_decimale(request.form("spese_bancarie"),"asp")
		else
			spese_bancarie_v=metodipagamento.spese_bancarie(rs("pagamento"))
		end if
		rs("spese_bancarie")=spese_bancarie_v
		rs("banca_appoggio")=request.form("banca_appoggio")
		rs("iban")=ucase(request.form("iban"))
		rs("creata_da")=session("nominativo")
		rs("eliminato")=0
		rs("fine_mese")=0
		rs("aliquota_iva")=iva(date())
		rs("trattamento_iva")=tmp_trattamento_iva
		rs("idbanca")=idbanca
		if tipo="accompagnatoria" then
			set rs_ddt=conn.execute("select ddt.causale from ddt where idord="&idord)
			if not rs_ddt.eof then
				rs("causale")=rs_ddt("causale")
			else
				rs("causale")=request.form("causale")
			end if
			set rs_ddt= Nothing
			rs("porto")=request.form("porto")
			rs("imballo")=request.form("imballo")
			rs("colli")=request.form("colli")
			rs("peso")=request.form("peso")
			rs("dimensione")=request.form("dimensione")
			rs("vettore")=request.form("vettore")
		else
			rs("causale")=1
			rs("fine_mese")=1
		end if
		rs("pa")=fattura_pa
		pagamento=rs("pagamento")
		rs.update
		IDfat=get_last_id("fatture")
		rs.Close
		sql="select ordini_fatture.* FROM ordini_fatture"
		rs.Open sql, conn, 1, 3
		rs.addnew
		rs("idord")=idord
		rs("idfat")=idfat

		rs.Update

		'creo scadenze

		testo="Creata [fattura="&IDfat&"]"&pre_fattura(data)&Nfat&"/"&year(data)&"[/fattura] per "&fatturaper
		rs_ordini("spese_bancarie_ordine")=spese_bancarie_v
		rs_ordini.update
		rs.close
		set rs=nothing
		rs_ordini.close
		set rs_ordini=nothing

		totale_fattura=calcola_totale_fattura(IDfat)
		testo=testo&" di importo "&formatcurrency(totale_fattura,2)
		n_incassi=collega_incassi (0,IDfat)
		if totale_fattura>0 then
			n=metodipagamento.genera_scadenze(IDfat,Nfat,idord,data,totale_fattura,pagamento)
		else
			n=0
		end if
		testo=testo&", eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&n_incassi&" incassi precedenti"&vbcrlf&n&" scadenze incasso"&vbcrlf&"Idord: "&IDORD
		add2log testo,2
		crea_fattura=IDfat&"|"&Nfat
	end if
end function


function crea_fattura_fattura(iduser)
	dim fattura_pa,rs_tmp
	call add2log(queryeform(),0)
	Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class
	anno=request.form("anno")
	if anno="" then
		data=date()
		anno=year(data)
	else
		'Anno impostato quindi precedente
		'Cerco la data
		data=conn.execute("select max(data) from fatture where anno="&anno&" and pa="&fattura_pa)(0)
	end if
	fattura_pa=0
	Nfat=max_fatt(fattura_pa,0,anno)

	sql="select * from utenti where iduser="&iduser
	set rs_utente= conn.execute("select * from utenti where iduser="&iduser)


	Set rs = Server.CreateObject("ADODB.Recordset")
	'Aggiungo fattura
	sql="select fatture.* FROM fatture"
	rs.Open sql, conn, 1, 3
	rs.addnew
	RS("nfat")=Nfat
	rs("data")=data
	rs("tipo_fattura")="fattura"
	rs("anno")=year(data)
	rs("iduser")=iduser
	rs("stato_fattura")=1

	nome=AllFirstUp(trim(request.form("nome")))
	cognome=AllFirstUp(trim(request.form("cognome")))
	azienda=AllFirstUp(trim(request.form("azienda")))


	idintestazione= aggiungi_intestazione(iduser,azienda,cognome,nome,ucase(trim(request.form("cf"))),trim(request.form("piva")),trim(request.form("Indirizzo")),allfirstup(trim(request.form("citta"))),trim(request.form("cap")),trim(request.form("provincia")))



	rs("idintestazione")=idintestazione

	'Creo dati_consegna
	idconsegna=0
	if request.form("d_indirizzo")<>"" then
		idconsegna=aggiungi_consegna(iduser,request.form("d_azienda"),request.form("d_indirizzo"),request.form("d_citta"),request.form("d_cap"))
	end if
	if request.form("consegna")<>"" then
		idconsegna=request.form("consegna")
	end if

	rs("idconsegna")=idconsegna
	rs("idord")=0
	rs("annotazioni")=request.form("annotazioni")
	if request.form("data_inizio")<>"" then
		rs("data_inizio")=request.form("data_inizio")
	end if
	rs("incaricato")=request.form("incaricato")
	rs("pagamento")=request.form("pagamento")
	if request.form("CodiceDestinatario")<>"" then
		rs("CodiceDestinatario")=request.form("CodiceDestinatario")
	end if
	spese_bancarie_v=0
	rs("spese_bancarie")=0
	rs("banca_appoggio")=request.form("banca_appoggio")
	rs("iban")=ucase(request.form("iban"))
	rs("creata_da")=session("nominativo")
	rs("eliminato")=0
	rs("fine_mese")=0
	rs("aliquota_iva")=iva(date())
	rs("trattamento_iva")=request.form("trattamento_iva")
	rs("idbanca")=request.form("idbanca")
	rs("causale")=request.form("causale")
	rs("porto")=request.form("porto")
	rs("imballo")=request.form("imballo")
	if request.form("colli")<>"" then
		rs("colli")=request.form("colli")
	end if
	rs("peso")=request.form("peso")
	rs("dimensione")=request.form("dimensione")
	rs("vettore")=request.form("vettore")
	rs("pa")=fattura_pa
	pagamento=rs("pagamento")
	rs.update
	IDfat=get_last_id("fatture")
	rs.Close

	'creo scadenze
	testo="Creata [fattura="&IDfat&"]"&pre_fattura(data)&Nfat&"/"&year(data)&"[/fattura] per "&fatturaper
	set rs=nothing


	testo=testo&", eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&n_incassi&" incassi precedenti"&vbcrlf&n&" scadenze incasso"&vbcrlf&"Idord: "&IDORD
	add2log testo,2
	crea_fattura_fattura=IDfat&"|"&Nfat
end function






function verifica_scadenze(idfat)
	verifica_scadenze=0
	dim rs
	sql="select Sum(incassi.importo) AS Somma, Count(incassi.idincasso) AS Conteggio FROM incassi WHERE incassi.idfat="&idfat&";"
	set rs=conn.execute (sql)
	somma=rs("somma")
	if not isnull(somma) then somma=cdbl(somma)
	conteggio=rs("conteggio")
	set rs=nothing
	sql="select fatture.* from fatture where idfat="&idfat
	set rs=conn.execute (sql)
	importo=cdbl(rs("totale_fattura"))
	set rs=nothing
	session("verifica_scadenze")="conteggio incassi:"&conteggio&" somma:"&somma&" importo:"&importo

	if importo<=somma then
		session("verifica_scadenze")=session("verifica_scadenze")&"<br>Elimino scadenze"
		conn.execute "update scadenze set pagato=1 WHERE idfat="&idfat,verifica_scadenze
	end if
end function
sub reset_tabella_ordini(tabella)
	Application.Lock
	application("cache_tabella_"&tabella)=""
	Application.UnLock
end sub
function get_tabella_ordini(tabella)
	dim stringa, rs,time_start
	if application("cache_tabella_"&tabella)="" then
		time_start=timer()
		where_tipo_documento=""
		if tabella="ordini" then
			link="ordini"
			where_tipo_documento=" and tipo_documento='ordine'"
		elseif tabella="preventivi" then
			link="preventivi"
		elseif tabella="ordini_fornitori" then
			link="ordini_fornitori"
		end if


		stringa="<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tabella1""><tr>"

		if tabella="ordini" then
			sql= "select Count(*) AS ConteggioDistato FROM "&tabella&" WHERE eliminato=0"&where_tipo_documento
			sql=sql&" and da_stampare=1 "
			set rs=conn.execute(sql)
			if clng(rs("ConteggioDistato"))>0 then
				stringa=stringa&"<td align=""center"" ><a href=""pag_adm_"&link&".asp?cercain=da_stampare"">Da stampare: "&rs("ConteggioDistato")&"</a></td>"
			end if
			rs.close
		end if
		sql= "select ordini.stato, Count(ordini.stato) AS ConteggioDistato FROM ordini WHERE ordini.eliminato=0"&where_tipo_documento&" GROUP BY ordini.stato ORDER BY ordini.stato;"
		sql=replace(sql,"ordini",tabella)
		set rs=conn.execute(sql)
		do until rs.EOF
			if clng(rs("ConteggioDistato"))>0 then
				if tabella="ordini" then
					testo=stato_ordine(rs("stato"))
				elseif tabella="preventivi" then
					testo=stato_preventivo(rs("stato"))
				elseif tabella="ordini_fornitori" then
					testo=stato_ordine_fornitore(rs("stato"))
				end if

				stringa=stringa&"<td align=""center"" class=""stato_"&rs("stato")&"""><a href=""pag_adm_"&link&".asp?cercain=stato"&rs("stato")&""">"&testo&": "&rs("ConteggioDistato")&"</a></td>"
			end if
			rs.movenext
		loop
		rs.close
		set rs=nothing

		Application.Lock
		Application("cache_tabella_"&tabella)=stringa&"</tr></table>"
		Application.Unlock
		call add2log("Ricreata cache cache_tabella_"&tabella &" in "&round(timer()-time_start,2),1)
	end if
	get_tabella_ordini=Application("cache_tabella_"&tabella)
end function
sub riordina_dettaglio(idord,tabella)
	sql="select * from "&tabella&"_dett where idord="&idord&" order by ordine, iddett"
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	i=2
	do while not rs.eof
		rs("ordine")=i
		rs.update
		rs.movenext
		i=i+2
	loop
	rs.close
end sub
sub carica_tipologia()
	array_tipologia=split(Application("str_tipologia"),"||")
	for n=0 to ubound(array_tipologia)
		rec_tipologia=split(array_tipologia(n),"|")
		Dict_tipol.SetKey = cint(rec_tipologia(0)) 'Key record
		Dict_tipol.SetField "tipologia", rec_tipologia(1)
		Dict_tipol.SetField "colore", rec_tipologia(2)
		Dict_tipol.Update 'Bind the record and preapre for the next record
	next
	erase array_tipologia
end sub
Dim Dict_tipol    'The Dict Object
Set Dict_tipol = New MultiDimensionalDictionary 'Create an Instance of the Class

call carica_tipologia()


function CancellettoSeNull (numero)
	if isnull(numero) then
		CancellettoSeNull="#"
	else
		CancellettoSeNull=numero
	end if
end function

function elenca_seriali(iddett,quantita)
	'Cerco nei seriali
	txt_seriale=""
	conteggio_seriali=0
	set seriali=conn.execute ("select numeri_seriali.* from numeri_seriali where iddett="&iddett)
	do while not seriali.eof
		if txt_seriale<>"" then txt_seriale=txt_seriale&", "
		txt_seriale=txt_seriale&seriali("seriale")
		conteggio_seriali=conteggio_seriali+1
		seriali.movenext
	loop
	if txt_seriale<>"" then
		txt_seriale= "Numeri di serie: "&txt_seriale
	end if
	if quantita>0 then
		mancanti=cint(quantita)-conteggio_seriali
		if mancanti>0 then
			if txt_seriale<>"" then txt_seriale=txt_seriale&", "
			txt_seriale=txt_seriale& mancanti&" numeri di serie mancanti"
		end if
	end if
	elenca_seriali=txt_seriale
end function


sub tabella_incassi(idord,idfat,totale,script,chiusa)
	dim sql
	if chiusa then
		nascosto="display: none;"

	end if
	if idfat>0 then
		sql="select incassi.*, fatture.IDFat, ordini.nord FROM incassi left JOIN ordini_fatture ON incassi.idord = ordini_fatture.idord inner JOIN fatture ON incassi.idfat = fatture.IDFat left JOIN ordini ON incassi.idord = ordini.idord where fatture.idfat="&idfat
	else
		sql="select * FROM incassi where idord="&idord &" order by data"
		sql="select incassi.*, ordini.nord FROM (incassi left JOIN ordini ON incassi.idord = ordini.idord) where ordini.idord="&idord
		'style=""
	end if
	if totale="null" then
		if idord>0 then
			set rs=conn.execute ("select ordini.totale, ordini.nord from ordini where idord="&idord)
			if not rs.eof then
				totale=rs("totale")
			else
				call add2log("Errore 1 tabella_incassi: ordini.totale non trovato per  idord="&idord&" idfat="&idfat&"totale="&totale&"script="&script,0)
			end if
		end if
		if idfat>0 then
			set rs=conn.execute ("select fatture.totale_fattura from fatture where idfat="&idfat)
			if not rs.eof then
				totale=rs("totale_fattura")
			else
				call add2log("Errore 2 tabella_incassi: fatture.totale_fattura non trovato per  idord="&idord&" idfat="&idfat&"totale="&totale&"script="&script,0)
			end if
		end if
		if totale="null" then
			call add2log("Errore 3 tabella_incassi: impossibile trovare totale per  idord="&idord&" idfat="&idfat&"totale="&totale&"script="&script,0)
		end if
		if isnull(totale) then
			call add2log("Errore 4 tabella_incassi: TOTALE NULL per [ordine="&idord&"]"&rs("nord")&"[/ordine] idord="&idord&" idfat="&idfat&"totale="&totale&"script="&script,0)
			totale=0
		end if
		set rs = nothing
	end if
	totale=cdbl(totale)
	da_saldare=true
	bgcolor="FFCCCC"
	set rs_incassi=conn.execute(sql)
	ci_sono_incassi=false
	%>
	<div class="div_tabella_incassi">
    <!-- incassi -->

	<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1" id="tabella_incassi">
        <tr><td colspan=5 class="ui-widget-header">Incassi</tr>
		<tbody id="body_incassi" style="<%=nascosto%>">
        <%if not rs_incassi.eof then
        	ci_sono_incassi=true

        %>
	        <tr>
	          <td align="center" bgcolor="#E5E5E5"></td>
	          <td align="center" bgcolor="#E5E5E5">Data</td>
	          <td align="center" bgcolor="#E5E5E5">Importo</td>
	          <td align="center" bgcolor="#E5E5E5">Causale / Tipo pagamento</td>
	          <td align="center" bgcolor="#E5E5E5">Note</td>
	        </tr>
	            <%
				tot_incassi=0
	            do while not rs_incassi.eof
				cancellabile=false
				if idfat>0 then
					note="Ordine: "&rs_incassi("nord")
				end if
				%>
	            <tr>
				<td align="center" style="border-bottom:1px solid;">
					<%if ha_il_permesso("B3") then%>
		              <div class="div_icona" ><span class="ui-icon ui-icon-triangle-1-s ui-corner-all left_menu_incassi bg-gray" id="incasso_<%=rs_incassi("idincasso")%>"></span></div><%=rs_incassi("idincasso")%>
		              <%end if%>
			  </td>


	              <td align="center" style="border-bottom:1px solid;"><%= formatDateTime(rs_incassi("data"), vbShortDate)%>
	         </td>
	              <td align="center" style="border-bottom:1px solid;"><%if rs_incassi("importo")<>"" then%><%= formatcurrency(rs_incassi("importo"), 2)%><%end if%></td>
	              <td align="center" style="border-bottom:1px solid;"><%=causale_pagamento(rs_incassi("causale"))%> - <%=rs_incassi("tipo_pagamento")%></td>
	              <td style="border-bottom:1px solid;"><%=rs_incassi("note_pagamento")&note%> </td>
	            </tr>
	                        <%

				tot_incassi=tot_incassi+cdbl(rs_incassi("importo"))
				rs_incassi.movenext
				loop
				'Calcolo totali
			saldo=roundup(totale-tot_incassi,2)
			testo="SALDATO"
			if saldo>0 then
				bgcolor="FFCCCC" 'Da saldare
				da_saldare=true
				testo="DA SALDARE "
				if tot_incassi=0 then
					testo=testo&"<br><strong>Nessuno</strong>"
				else
					testo=testo&"<br>Totale versato "&formatcurrency(tot_incassi)&" <b>Da saldare "&formatcurrency(saldo)&"</b>"
				end if
			elseif saldo <0 then
				bgcolor="CC00CC" 'Saldato in più
				testo_viola="(+"&formatcurrency(saldo*-1)&")"
				da_saldare=false
				testo="SALDATO"
			else
				da_saldare=false
				bgcolor="99FF99" 'Saldato
				testo="SALDATO"
			end if
			end if
			rs_incassi.close
			set rs_incassi=Nothing
			'Verifiche fuori dal ciclo incassi

			if isnull(totale) then
				call add2log("Errore tabella_incassi: TOTALE NULL per [ordine="&idord&"]"&rs("nord")&"[/ordine] idord="&idord&" idfat="&idfat&"totale="&totale&"script="&script,0)
				totale=0
			end if

			if isnull(stato) then
				call add2log("Errore tabella_incassi: STATO NULL per [ordine="&idord&"]"&rs("nord")&"[/ordine] idord="&idord&" idfat="&idfat&"totale="&totale&"script="&script,0)
				totale=0
			end if
			if isnull(verde) then
				call add2log("Errore tabella_incassi: VERDE NULL per [ordine="&idord&"]"&rs("nord")&"[/ordine] idord="&idord&" idfat="&idfat&"totale="&totale&"script="&script,0)
				totale=0
			end if

			if verde<>"" then
				da_saldare=false
				bgcolor="CCFF99" ''Verdino chiaro reso tutto
				testo=verde
			elseif totale=0 and stato>5 then
				da_saldare=false
				bgcolor="99FF99" 'Saldato
				testo="IMPORTO 0"
			end if

			%>
</tbody>
<%
	if false then
'if da_saldare=false then
 %>
			<tr bgcolor="<%=bgcolor%>">
			    <td colspan=5 align="center"><b><%=testo%></b> <%=testo_viola%>
			      <%if ci_sono_incassi then %><input type="button" id="pulsante_tabella_incassi" onClick="showhide_body_incassi()" value="Mostra incassi" style="font-size:8pt;" class="right"><%end if %>
			</tr>
<%'else
	end if
%>
			<tr bgcolor="<%=bgcolor%>">
	            <td colspan=5 align="center">
		            <b><%=testo%></b> <%=testo_viola%>

				<span class="right">
		            <%if 	ci_sono_incassi then %>
			            <input type="button" id="pulsante_tabella_incassi" onClick="showhide_body_incassi()" value="Mostra incassi" style="font-size:8pt">
		            <%end if%>
		            <%if da_saldare then %>
	            <%if esiste_fattura  then %>
	              <input type="button" id="agg_incasso"  value="Aggiungi incasso sulla fattura" style="font-size:8pt">
              <%else %>
	              <input type="button" id="agg_incasso"  value="Aggiungi incasso" style="font-size:8pt">

              <%end if
	             end if
              %>

              </span>
          </tr>
<%'end if%>

	</table>
	</div>

<%
if script then
	if idord="" then idord=0
	if idfat="" then idfat=0

	%>
	<script>

	$(function(){
			var idord=<%=idord%>;
			var idfat=<%=idfat%>;
			//$("#agg_incasso").click(function(e) {
			$(document).on("click","#agg_incasso",function(e) {


			//$("#dialog").dialog({
			//	height: 700,
			//	width: 700,
			//	modal: true,
			//	title:"Aggiungi incasso "
			//});
			$("#dialog").html("<center>Attendi...</center>");
			$.ajax({
				url     : "pag_adm_dialog_incasso.asp?idord="+idord+"&idfat="+idfat,
				type    : "post",
				cache: false,
				//dataType: 'json',
				//data	: dati,
				success: function(data){
					//alert(data.Message);

					$("#dialog").html(data);
					$("#dialog").dialog({height: "auto"});
				}
				,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
				      console.log("xhr.statusText:"+xhr.statusText);
				      console.log("xhr.responseText:"+xhr.responseText);
				      console.log("textStatus:"+textStatus);
				      console.log("error:"+error);
					  }
			});
			$("#dialog").dialog({
				autoOpen: true,
				modal: true,
				resizable: false,
				width: 600,
				height: "auto",
				title: "Aggiungi incasso "
			});
		});
	})
	function showhide_body_incassi()
	{
		if($('#body_incassi').css('display') == 'none'){
			$('#body_incassi').fadeIn(800);
			$('#pulsante_tabella_incassi').prop('value', 'Nascondi incassi');
		}
		else
		{
			$('#body_incassi').fadeOut(200);
			$('#pulsante_tabella_incassi').prop('value', 'Mostra incassi');
		}
	}
	</script>
	<%
	end if
end sub
sub tabella_scadenze(idfat,idord)
	paga=false
	if idfat>0 then
		  sql="select scadenze.* FROM scadenze where pagato=0 and  idfat="&idfat
	elseif idord>0 then
		paga=true
		sql="select scadenze.* FROM scadenze where pagato=0 and idord="&idord
	else
		sql=""
	end if
	%>

	<div class="div_tabella_scadenze dirtyignore">


	<%
	if sql<>"" then
			colspan=4
			if paga then colspan=colspan+1
			sql=sql&" order by data"
			set rs_incassi=conn.execute(sql)
			if not rs_incassi.eof then
		  %>
          <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1" id="tabella_scadenze">
            <tr>
            <td colspan="<%=colspan%>" class="ui-widget-header">Scadenze pagamenti</tr>
            <tr>
              <td align="center" bgcolor="#E5E5E5"></td>
              <td align="center" bgcolor="#E5E5E5">Data</td>
              <td align="center" bgcolor="#E5E5E5">Importo</td>
              <td align="center" bgcolor="#E5E5E5">Note</td>
              <%if paga then %>
              <td align="center" bgcolor="#E5E5E5">Pagato</td>
              <%end if %>
            </tr>
            <%
            do while not rs_incassi.eof
            if rs_incassi("pagato")=1 then
	    		colore="bgcolor='#99FF99'" 'Verdino saldata per importo 0
            else
	            colore=""
	        end if
			if rs_incassi("data")<date() then
				classe="rosso"
			else
				classe=""
			end if

			%>
            <tr id="scadenza<%=rs_incassi("idscadenza")%>"  <%=colore%> class="<%=classe%>">



              <td align="center" style="border-bottom:1px solid;">
					<%if ha_il_permesso("B3") then%>
		              <div class="div_icona" ><span class="ui-icon ui-icon-triangle-1-s ui-corner-all left_menu_incassi bg-gray" id="scadenza_<%=rs_incassi("idscadenza")%>"></span></div>
		              <%end if%>


	              </td>
              <td align="center" style="border-bottom:1px solid;"><%= formatDateTime(rs_incassi("data"), vbShortDate)%>
         </td>
              <td align="center" style="border-bottom:1px solid;"><%if rs_incassi("importo")<>"" then%><%= formatcurrency(rs_incassi("importo"), 2)%><%end if%></td>
              <td style="border-bottom:1px solid;"><%=rs_incassi("note_pagamento")%></td>
              <%if paga then %>
              <td align="center">
					<%if ha_il_permesso("B3") then%>
						<input type="checkbox"  value="1" class="paga" <%if rs_incassi("pagato")=1 then response.write "checked"%>/>
		              <%end if%>
              </td>
              <%end if %>
            </tr>
                        <%

			rs_incassi.movenext
			loop
			rs_incassi.close
			%>
			</table>
			<%
			if paga then %>
				<script>
				$(function() {
					$(".paga").change(function(e){
						var id=$(this).closest("tr").attr("id").replace("scadenza","");
						var checked=$(this).is(':checked')?1:0;
						console.log(id+" "+checked);

                    $.ajax({
                        url: "ajax_function.asp",
                        type: "post",
                        dataType: 'html',
                        cache: false,
                        data	: {
	                        oper: "paga_scadenza",
	                        idscadenza: id,
	                        pagato: checked,
	                        idord: idord
                        },
                        success: function(data) {

                            //$("#div_tabella_incassi").html(data);
                            //Ciclo su più div
                            $(".div_tabella_scadenze").each(function() {
                                var self = $(this);
                                self.html(data);
                            });

                        },
                        error: function(xhr, textStatus, error) {
                            toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
                            console.log("xhr.statusText:" + xhr.statusText);
                            console.log("xhr.responseText:" + xhr.responseText);
                            console.log("textStatus:" + textStatus);
                            console.log("error:" + error);
                        }
                    });


					});


				});
				</script>
			<%
			end if
			end if
	end if
	%>
	</div>

	<%

end sub

sub tabella_scadenze_for(idfat)
	sql="select * from scadenze_for where idfat="&idfat &" order by scadenza"
	totale_scadenze=0
	set rs=conn.execute(sql)
	if not rs.eof then

	%>
	          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1" id="scadenze_for">
	            <tr>
	              <td colspan="4" class="ui-widget-header"><span style="float:left;">Scadenze</span>  <span style="float: right;"><input type="button" value="Aggiungi" id="aggiungi_scadenza"></span></td>
	            </tr>
	            <tr style="border-bottom:1px solid;">
	              <td width="10" align="center" nowrap bgcolor="#E5E5E5">&nbsp;</td>
	              <td bgcolor="#E5E5E5" >Scadenza</td>
	              <td bgcolor="#E5E5E5" >Importo</td>
	              <td bgcolor="#E5E5E5" align="center" >Pagato</td>
	            </tr>
	            <%

		do until rs.EOF
		totale_scadenze=totale_scadenze+cdbl(rs("importo"))
		if rs("pagato")=1 then
			colore="bgcolor='#99FF99'"
		else
			colore=""
		end if
%>
            <tr class="coprobox" <%=colore%> style="border-bottom:1px solid;" id="scadenza_<%=rs("id")%>">
              <td align="center" >
	              <div class="div_icona" ><span class="ui-icon ui-icon-triangle-1-s ui-corner-all left_menu_incassi bg-gray" ></span></div>
	              </td>
              <td ><%=rs("scadenza")%></td>
              <td ><%=formatcurrency(rs("importo"),2)%></td>
              <td align="center" ><input type="radio" class="paga_scadenza" /></td>
            </tr>
            <%
	rs.movenext
	loop
	rs.close
	Set rs = Nothing
	end if
	%>
          </table>
          <script>
			$(function() {
				$("#aggiungi_scadenza").click(function(){
					$("#dialog").html("Attendi...");
                    $.ajax({
                        url: "pag_adm_dialog_scadenza_for.asp?idfat=<%=idfat%>",
                        type: "post",
                        //dataType: 'json',
                        //data	: dati,
                        cache: false,
                        success: function(data) {
                            //alert(data.Message);

                            $("#dialog").html(data);
                        },
                        error: function(xhr, textStatus, error) {
                            toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
                            console.log("xhr.statusText:" + xhr.statusText);
                            console.log("xhr.responseText:" + xhr.responseText);
                            console.log("textStatus:" + textStatus);
                            console.log("error:" + error);
                        }
                    });
                    $("#dialog").dialog({
                        autoOpen: true,
                        modal: true,
                        resizable: false,
                        width: 600,
                        height: "auto",
                        title: "Aggiunfi scadenza"
                    });
				});



	          });


	      </script>

          <%
          if utente_andrea then response.write "totale scadenze:"&totale_scadenze

          %>
          <%


end sub


sub report_errore()
    if session("report")<>"" then
     %>
		<div class="ui-widget" id="report_errore">
			<div class="ui-state-error ui-corner-all" style="padding: 0 .7em;">
				<p>
					<span class="ui-icon ui-icon-alert"
						style="float: left; margin-right: .3em;"></span>
					<strong>Attenzione:</strong> <%=session("report")%>
				</p>
				<p>
				<a href="<%=questofile%>?cancella_report=si" id="cancella_report">cancella messaggio</a>
				</p>
			</div>
		</div>
		<script>
			$(function(){
					$("#cancella_report").click(function(e){
						e.preventDefault();
						$.ajax({
							url     : "ajax_function.asp?oper=cancella_report",
							type    : "post",
							cache: false,
							dataType: 'json',
							//data	: dati,
							success: function(data){
								console.log(data.success+data.message);
								$("#report_errore").remove();
							}
							,error: function(xhr, textStatus, error){
									toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
									toastr.error('Errore cancella_report');
							      console.log("xhr.statusText:"+xhr.statusText);
							      console.log("xhr.responseText:"+xhr.responseText);
							      console.log("textStatus:"+textStatus);
							      console.log("error:"+error);
								  }
						});




					});

				});
		</script>

    <%end if

end sub
function codifica_varianti_ordine(variante_a, nome_variante_a, variante_b, nome_variante_b)
	dim txt
	codifica_varianti_ordine=""
	txt=""
		if variante_a<>"" then
			txt = nome_variante_a & ": " & variante_a
		end if
		if variante_b<>"" then
			if txt<>"" then txt=txt&" - "
			txt =txt & nome_variante_b & ": " & variante_b
		end if
		if txt<>"" then
			codifica_varianti_ordine=txt
		end if
end function
sub tabella_accessori(iduser)


	%>
		    <div id="accessori">
		    <%
			set rs= conn.execute ("select mod_dip from utenti_clienti where iduser="&iduser)
			if not rs.eof then
				jsonString=rs(0)
				if jsonString="" then
				jsonString="{}"
				end if
			else
				jsonString="{}"
			end if
			set rs = nothing
			'jsonString="{""velcrocamicia"":1}"
			'Response.write(jsonString)
			set JSON = New JSONobject
			JSON.Parse(jsonString)

			%>
		<label>Velcro DX</label>
		<span class="option <%=green(json("velcrocamicia"))%>">Camicia</span><span class="option <%=green(json("velcrogiacca"))%>">Giacca</span>
		<span class="option <%=green(json("velcrogiubbino"))%>">Giubbino</span>

		<br />
		<label>Cintura e passanti</label><%if json("cintura")=1 then %> <span class="option green">SI</span><%else %><span class="option">NO</span><% end if %>

		<br>
		<label>Camicia M/L</label>
		<span class="option <%=green(json("camiciam"))%>">Militare</span><span class="option <%=green(json("camiciac"))%>">Civile</span>
		<br>
		<label>Camiciotto M/M militare</label>
		<span class="option <%=green(json("camiciottob"))%>">Tutto bottoni</span><span class="option <%=green(json("camiciottop"))%>">Polo</span>
		<br>
		<label>Interasse fori berretto</label><%=json("interasse")%> cm		<br>
		<label>Note</label><%=json("note")%><br>
	    </div>
	    <%
		    end sub


function recupera_taglia(tipo_prodotto,iddip)
	if tipo_prodotto=10 then
		sql="select arma,lato_arma from dipendenti where iddip="&iddip
		set rs=conn.execute(sql)
		if not rs.eof then
			recupera_taglia=rs(0)&" "&rs(1)
		else
			recupera_taglia=""
		end if
	elseif tipo_prodotto=14 then
		sql="select ornamento from (dipendenti left join utenti_gradi on dipendenti.grado = utenti_gradi.id) where iddip="&iddip
		recupera_taglia=conn.execute(sql)(0)
	elseif tipo_prodotto=16 then
		sql="select matricola from dipendenti where iddip="&iddip
		recupera_taglia=conn.execute(sql)(0)
	elseif tipo_prodotto=17 then
		sql="select lingue from dipendenti where iddip="&iddip
		lingue=conn.execute(sql)(0)
		if lingue<>"" then
			'call add2log("Lingua:"&lingue,0)
			lingue=split(lingue,",")
			if isarray(lingue) then
				recupera_taglia=""
					'call add2log("ubound(lingue):"&ubound(lingue),0)
				for n=0 to ubound(lingue)
					recupera_taglia=recupera_taglia&"<img src=""/images/"&trim(lingue(n))&"-flag.png"">"
					'call add2log("recupera_taglia aggiungo:"&lingue(n),0)

				next
				'call add2log("recupera_taglia array:"&recupera_taglia,0)
			else
				recupera_taglia="<img src=""/images/"&trim(lingue)&"-flag.png"">"
				'call add2log("recupera_taglia no array:"&recupera_taglia,0)

			end if
		end if
	else


		recupera_taglia=""
		sql="SELECT * FROM dipendenti_misure where iddip="&iddip&" order by data_rilievo desc limit 0,1 "
		set rs_misure=conn.execute (sql)
		if not rs_misure.eof then
			select case cint(tipo_prodotto)
				case 0
					'Nessuna
				case 1	'GIACCA
					'tipo_prodotto_spettanze="Giacca"
					'recupera_taglia="TG:"&rs_misure(3)&" Fondo:"&rs_misure(4)&" Manica:"&rs_misure(5)&" Torace:"&rs_misure(6)&" Vita:"&rs_misure(7)&" Spalle"&rs_misure(8)&""
					'if rs_misure(3)<>"" then
						'call concatena_stringa(recupera_taglia," ","TG:"&rs_misure(3))
						if rs_misure(40)<>"" then
							call concatena_stringa(recupera_taglia," ","TG: ("&rs_misure(40)&")")
						end if
					'end if
					'if rs_misure(4)<>"" then
						call concatena_stringa(recupera_taglia," ","Fondo:"&rs_misure(4))
						if rs_misure(27)<>"" then
							call concatena_stringa(recupera_taglia," ","("&rs_misure(27)&")")
						end if
					'end if
					'if rs_misure(5)<>"" then
						call concatena_stringa(recupera_taglia," ","Manica:"&rs_misure(5))
						if rs_misure(28)<>"" then
							call concatena_stringa(recupera_taglia," ","("&rs_misure(28)&")")
						end if
					'end if
					'if rs_misure(6)<>"" then
						'call concatena_stringa(recupera_taglia," ","Torace:"&rs_misure(6))
						if rs_misure(29)<>"" then
							call concatena_stringa(recupera_taglia," ","Torace: ("&rs_misure(29)&")")
						end if
					'end if
					'if rs_misure(7)<>"" then
						'call concatena_stringa(recupera_taglia," ","Vita:"&rs_misure(7))
						if rs_misure(30)<>"" then
							call concatena_stringa(recupera_taglia," ","Vita: ("&rs_misure(30)&")")
						end if
					'end if
					'if rs_misure(42)<>"" then
						'call concatena_stringa(recupera_taglia," ","Bacino:"&rs_misure(42))
						if rs_misure(43)<>"" then
							call concatena_stringa(recupera_taglia," ","Bacino: ("&rs_misure(43)&")")
						end if
					'end if
					'if rs_misure(8)<>"" then
						'call concatena_stringa(recupera_taglia," ","Spalle"&rs_misure(8))
						if rs_misure(31)<>"" then
							call concatena_stringa(recupera_taglia," ","Spalle: ("&rs_misure(31)&")")
						end if
					'end if
					'recupera_taglia="TG:"&rs_misure(3)&" Fondo:"&rs_misure(27)&" Manica:"&rs_misure(28)&" Torace:"&rs_misure(29)&" Vita:"&rs_misure(30)&" Spalle"&rs_misure(31)&""
						if rs_misure(48)<>"" then
							call concatena_stringa(recupera_taglia," ","["&rs_misure(48)&"]")
						end if

				case 2	'PANTALONI
					'tipo_prodotto_spettanze="Pantaloni"
					'recupera_taglia="TG:"&rs_misure(9)&" Lunghezza:"&rs_misure(10)&" Vita:"&rs_misure(11)&" Bacino:"&rs_misure(12)&" Cosce:"&rs_misure(18)&" Cavallo:"&rs_misure(19)&" Polpacci:"&rs_misure(20)&" "
					'if rs_misure(9)<>"" then
						'call concatena_stringa(recupera_taglia," ","TG:"&rs_misure(9))
						if rs_misure(41)<>"" then
							call concatena_stringa(recupera_taglia," ","TG: ("&rs_misure(41)&")")
						end if
					'end if
					'if rs_misure(10)<>"" then
						'call concatena_stringa(recupera_taglia," ","Lunghezza:"&rs_misure(10))
						if rs_misure(33)<>"" then
							call concatena_stringa(recupera_taglia," ","Lunghezza: ("&rs_misure(33)&")")
						end if
					'end if
					'if rs_misure(11)<>"" then
						'call concatena_stringa(recupera_taglia," ","Vita:"&rs_misure(11))
						if rs_misure(34)<>"" then
							call concatena_stringa(recupera_taglia," ","Vita: ("&rs_misure(34)&")")
						end if
					'end if
					'if rs_misure(12)<>"" then
						'call concatena_stringa(recupera_taglia," ","Bacino:"&rs_misure(12))
						if rs_misure(35)<>"" then
							call concatena_stringa(recupera_taglia," ","Bacino: ("&rs_misure(35)&")")
						end if
					'end if
					'if rs_misure(18)<>"" then
						'call concatena_stringa(recupera_taglia," ","Cosce:"&rs_misure(18))
						if rs_misure(36)<>"" then
							call concatena_stringa(recupera_taglia," ","Cosce: ("&rs_misure(36)&")")
						end if
					'end if
					'if rs_misure(19)<>"" then
						'call concatena_stringa(recupera_taglia," ","Cavallo:"&rs_misure(19))
						if rs_misure(37)<>"" then
							call concatena_stringa(recupera_taglia," ","Cavallo: ("&rs_misure(37)&")")
						end if
					'end if
					'if rs_misure(20)<>"" then
						'call concatena_stringa(recupera_taglia," ","Polpacci:"&rs_misure(20))
						if rs_misure(38)<>"" then
							call concatena_stringa(recupera_taglia," ","Polpacci: ("&rs_misure(38)&")")
						end if
					'end if
						if rs_misure(47)<>"" then
							call concatena_stringa(recupera_taglia," ","["&rs_misure(47)&"]")
						end if

				case 3
					'tipo_prodotto_spettanze="Camicia"
					'call add2log("misura26:"&rs_misure(26),0)
					recupera_taglia=rs_misure(13)
					if rs_misure(26)<>"" then
						recupera_taglia=recupera_taglia&" ["&rs_misure(26)&"]"
					end if

				case 4
					'tipo_prodotto_spettanze="Scarpa"
					recupera_taglia=rs_misure(14)
						if rs_misure(56)<>"" then
							recupera_taglia=recupera_taglia&" Nota:"&rs_misure(56)
						end if

				case 5
					'tipo_prodotto_spettanze="Berretto"
					recupera_taglia=rs_misure(15)
						if rs_misure(58)<>"" then
							recupera_taglia=recupera_taglia&" Nota:"&rs_misure(58)
						end if
				case 6
					'tipo_prodotto_spettanze="Guanto"
					recupera_taglia=rs_misure(16)
						if rs_misure(59)<>"" then
							recupera_taglia=recupera_taglia&" Nota:"&rs_misure(59)
						end if

				case 7
					'tipo_prodotto_spettanze="Maglieria"
					recupera_taglia=ucase(rs_misure(21))
						if rs_misure(60)<>"" then
							recupera_taglia=recupera_taglia&" Nota:"&rs_misure(60)
						end if
				case 8
					'tipo_prodotto_spettanze="Giacca tecnica"
					if rs_misure(23)<>"" then
						recupera_taglia=rs_misure(23)
					end if
					if rs_misure(50)<>"" then	'Note
						call concatena_stringa(recupera_taglia," ","["&rs_misure(50)&"]")
					end if

				case 9
					'tipo_prodotto_spettanze="Pantaloni tecnici"
					if rs_misure(22)<>"" then
						recupera_taglia=rs_misure(22)
					end if
					if rs_misure(51)<>"" then	'Note
						call concatena_stringa(recupera_taglia," ","["&rs_misure(51)&"]")
					end if

				case 11
					'tipo_prodotto_spettanze="Gonna"
					'recupera_taglia="Lunghezza:"&rs_misure(24)&" Vita:"&rs_misure(11)&" Bacino:"&rs_misure(12)
					if rs_misure(24)<>"" then
						call concatena_stringa(recupera_taglia," ","Lunghezza:"&rs_misure(24))
						if rs_misure(39)<>"" then
							call concatena_stringa(recupera_taglia," ","("&rs_misure(39)&")")
						end if
					end if
					if rs_misure(34)<>"" then
						call concatena_stringa(recupera_taglia," ","Vita:"&rs_misure(34))
					end if
					if rs_misure(35)<>"" then
						call concatena_stringa(recupera_taglia," ","Bacino:"&rs_misure(35))
					end if
					if rs_misure(53)<>"" then	'Note
						call concatena_stringa(recupera_taglia," ","["&rs_misure(53)&"]")
					end if

				case 12
					'tipo_prodotto_spettanze="Cappotto"
					'recupera_taglia="Fondo:"&rs_misure(25)&" Manica:"&rs_misure(5)&" Torace:"&rs_misure(6)&" Vita:"&rs_misure(7)&" Spalle:"&rs_misure(8)
					if rs_misure(25)<>"" then
						call concatena_stringa(recupera_taglia," ","Lunghezza:"&rs_misure(25))
						if rs_misure(32)<>"" then
							call concatena_stringa(recupera_taglia," ","("&rs_misure(32)&")")
						end if
					end if
					if rs_misure(28)<>"" then
						call concatena_stringa(recupera_taglia," ","Manica:"&rs_misure(28))
					end if
					if rs_misure(29)<>"" then
						call concatena_stringa(recupera_taglia," ","Torace:"&rs_misure(29))
					end if
					if rs_misure(30)<>"" then
						call concatena_stringa(recupera_taglia," ","Vita:"&rs_misure(30))
					end if
					if rs_misure(31)<>"" then
						call concatena_stringa(recupera_taglia," ","Spalle:"&rs_misure(31))
					end if
					if rs_misure(52)<>"" then	'Note
						call concatena_stringa(recupera_taglia," ","["&rs_misure(52)&"]")
					end if

				case 13
					'tipo_prodotto_spettanze="Cintura"
					if rs_misure(44)<>""  then
						recupera_taglia=rs_misure(44)
						if rs_misure(62)<>"" then
							recupera_taglia=recupera_taglia&" Nota:"&rs_misure(62)
						end if

					end if
				case 15
					'tipo_prodotto_spettanze="Cinturone"
					if rs_misure(45)<>""  then
						recupera_taglia=rs_misure(45)
						if rs_misure(63)<>"" then
							recupera_taglia=recupera_taglia&" Nota:"&rs_misure(63)
						end if
					end if
				case 18
					'collant
					if rs_misure(54)<>""  then
						recupera_taglia=rs_misure(54)
						if rs_misure(57)<>"" then
							recupera_taglia=recupera_taglia&" Nota:"&rs_misure(57)
						end if

					end if


				case 19
					'polo
					if rs_misure(55)<>""  then
						recupera_taglia=rs_misure(55)
						if rs_misure(61)<>"" then
							recupera_taglia=recupera_taglia&" Nota:"&rs_misure(61)
						end if
					end if


				case else
					recupera_taglia=-1
			end select

			if rs_misure(46)=1 then
				recupera_taglia="<span class=""nuovo_rilievo"">"&recupera_taglia&"</span>"
			end if
			if rs_misure(46)=2 then
				recupera_taglia="<span class=""nuovo_rilievo_bottoni"">"&recupera_taglia&"</span>"
			end if

		else
			recupera_taglia=-2

		end if
	end if
	if isnull(recupera_taglia) then recupera_taglia=""

	'Ritorna -1 se tipo_prodotto non trovato
	'Ritorna -2 se misura dipendente non trovata
end function

function aggiorna_misure_in_ordini(iddip)
	aggiorna_misure_in_ordini=""
    dim rs_misure_in_ordini
    sql="select  ordini_dett_spettanze.id, ordini_dett.iddett, ordini_dett.articolo_ordine, ordini.nord, ordini.idord, prodotti.tipo_taglia from ((ordini_dett_spettanze inner join ordini_dett on ordini_dett_spettanze.iddett = ordini_dett.iddett) inner join ordini on ordini_dett.idord = ordini.idord) inner join prodotti on ordini_dett.idpro = prodotti.idpro where ordini.stato<=5 and ordini_dett_spettanze.iddip="&iddip
    set rs_misure_in_ordini=conn.execute (sql)
    do while not rs_misure_in_ordini.EOF

    	conn.execute ("update ordini_dett_spettanze set taglia_misura='"&replace(recupera_taglia(rs_misure_in_ordini("tipo_taglia"),iddip),"'","''") &"' where id="&rs_misure_in_ordini("id"))
		aggiorna_misure_in_ordini=aggiorna_misure_in_ordini&"<br>Aggiornata misura in [ordine="&rs_misure_in_ordini("idord")&"]"&rs_misure_in_ordini("nord")&"[/ordine] riga: "&rs_misure_in_ordini("iddett")&"  ordini_dett_spettanze.id="&rs_misure_in_ordini("id")

	    rs_misure_in_ordini.MoveNext
    loop

        sql="select  preventivi_dett_spettanze.id, preventivi_dett.iddett, preventivi_dett.articolo_ordine, preventivi.nord, preventivi.idord, prodotti.tipo_taglia from ((preventivi_dett_spettanze inner join preventivi_dett on preventivi_dett_spettanze.iddett = preventivi_dett.iddett) inner join preventivi on preventivi_dett.idord = preventivi.idord) inner join prodotti on preventivi_dett.idpro = prodotti.idpro where preventivi.stato<=5 and preventivi_dett_spettanze.iddip="&iddip
    set rs_misure_in_preventivi=conn.execute (sql)
    do while not rs_misure_in_preventivi.EOF

    	conn.execute ("update preventivi_dett_spettanze set taglia_misura='"&replace(recupera_taglia(rs_misure_in_preventivi("tipo_taglia"),iddip),"'","''") &"' where id="&rs_misure_in_preventivi("id"))
		aggiorna_misure_in_preventivi=aggiorna_misure_in_preventivi&"<br>Aggiornata misura in [ordine="&rs_misure_in_preventivi("idord")&"]"&rs_misure_in_preventivi("nord")&"[/ordine] riga: "&rs_misure_in_preventivi("iddett")&"  preventivi_dett_spettanze.id="&rs_misure_in_preventivi("id")

	    rs_misure_in_preventivi.MoveNext
    loop



end function

function determina_dipendente(iddip)
			set rs=conn.execute ("select dipendenti.cognome, dipendenti.nome, dipendenti.iddip from dipendenti  where dipendenti.iddip="&iddip)
			if not rs.eof then

				'denominazione_v=denominazione(rs("nome"),rs("cognome"),rs("azienda"))
				determina_dipendente=" dipendente iddip="&iddip&" <b>"&rs("cognome")&" "&rs("nome")&"</b>"
			else
				determina_dipendente=" dipendente id:"&iddip&" non trovato"
			end if
			Set rs = Nothing
end function

function mepa_tipo(stato)
	mepa_tipo=""
	if not isnumeric(stato) then
		mepa_tipo=stato
	else
		select case stato
		case -1
			mepa_tipo=4
		case 0
			mepa_tipo="Non specificato"
		case 1
			mepa_tipo="Ordine mepa"
		case 2
			mepa_tipo="Ordine diretto"
		case 3
			mepa_tipo="Trattativa diretta"
		case 4
			mepa_tipo="Rdo"
		end select
	end if
end function



'FUNZIONI PER PDF
sub T_totali_compila( byref T_totali, byref totord, byref sconto_ordine, trattamento_iva_v, aliquota_iva, spese_bancariev,fattura_sp)
	totord=round(totord,2)

	'TOTALE MERCE
	'T_totali.Rows(2).Cells(5).AddText formatcurrency(totord,2), "alignment=1; size=10; Expand=true;"


	if spese_bancariev>0 then
		T_totali.Rows(2).Cells(1).AddText formatcurrency(spese_bancariev,2), "alignment=1; size=8; Expand=true;" 'SPESE BANCARIE
		totord=totord+spese_bancariev
	end if

	'IMPONIBILE
	T_totali.Rows(2).Cells(2).AddText formatcurrency(totord,2), "alignment=1; size=8; Expand=true;"


	'SCONTO ORDINE
	if sconto_ordine>0 then
		totord=totord-sconto_ordine
		T_totali.Rows(4).Cells(1).AddText formatcurrency(sconto_ordine,2), "alignment=1; size=10; Expand=true;"
	end if


	'IVA
	if esenzione_iva(trattamento_iva_v) and trattamento_iva_v<>10 then
		valore_iva=0
		imposta=FormatCurrency(0,2)
		totaleconiva=totord
	else
		imposta=roundup(totord*aliquota_iva/100,2)
		totaleconiva=totord+imposta
		totord=totaleconiva
		imposta=FormatCurrency(imposta,2)
		if trattamento_iva_v=10 or fattura_sp then
			'valore_iva=formatcurrency(totord*aliquota_iva/100,2)&vbcrlf&vbcrlf&"-"&formatcurrency(totord*aliquota_iva/100,2)
			imposta=imposta&vbcrlf&vbcrlf&"OPERAZIONE SOGGETTA A SPLIT PAYMENT"
			totord=imponibile
		end if
	end if
	testo_esenzione=replace(trattamento_iva(trattamento_iva_v),"Normale","")

	T_totali.Rows(2).Cells(4).AddText imposta, "alignment=1; size=8; Expand=true;"
	T_totali.Rows(2).Cells(3).AddText aliquota_iva&"%"&vbcrlf&vbcrlf&testo_esenzione, "alignment=2; size=8; Expand=true;"

	'TOTALE
	totord=totord-sconto_ordine
	T_totali.Rows(4).Cells(5).AddText Formatcurrency(totaleconiva,2), "alignment=1; size=10; Expand=true;" 'TOTALE FATTURA


end sub


function aggiungi_consegna(iduser,d_azienda,d_indirizzo,d_citta,d_cap)


	'cerco negli indirizzi
	ricerca=replace(d_azienda&d_indirizzo&d_citta,"'","''")
	sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&ricerca&"' and iduser="&iduser
	'response.write "<br>"&sql&"<br>"
	set rs2=conn.execute(sql)
	if rs2.eof then
		set rs_consegna=Server.CreateObject("ADODB.Recordset")
		rs_consegna.open "utenti_consegna",conn,3,3
		rs_consegna.addnew
		rs_consegna("iduser")=iduser
		rs_consegna("d_azienda")=d_azienda
		rs_consegna("d_indirizzo")=d_indirizzo
		rs_consegna("d_citta")=d_citta
		'rs_consegna("d_regione")=request.form("d_regione")
		'rs_consegna("d_provincia")=request.form("d_provincia")
		rs_consegna("d_cap")=d_cap
		rs_consegna.update
		aggiungi_consegna=Get_last_id("utenti_consegna")
		rs_consegna.close
		set rs_consegna = Nothing
	else
		aggiungi_consegna=rs2("id")
	end if
	set rs2 = Nothing
end function
function dati_consegna(id)
	if id>0 then
		set rs_consegna=conn.execute("select * from utenti_consegna where id="&id)
		if rs_consegna.eof then
			dati_consegna="Dati consegna non trovati"
		else
			dati_consegna=""
			if rs_consegna("d_azienda")<>"" then
				dati_consegna="<b>"&rs_consegna("d_azienda")&"</b><br>"
			end if
			dati_consegna=dati_consegna&rs_consegna("d_Indirizzo")&"<br>"
			if rs_consegna("d_cap")<>"" then
				dati_consegna=dati_consegna&rs_consegna("d_cap")&" - "
			end if
			dati_consegna=dati_consegna&rs_consegna("d_citta")
			if rs_consegna("d_provincia")<>"" then
				dati_consegna=dati_consegna&" ("&rs_consegna("d_provincia")&")"
			end if
		end if
	else
		dati_consegna="IDEM"
	end if
end function

sub tab_dati_consegna(idconsegna,iduser)

	if false then
	'if idconsegna>0 then
		set rs_consegna=conn.execute("select * from utenti_consegna where id="&idconsegna)
		if rs_consegna.eof then
			dati_consegna="Dati consegna non trovati"
		else

			d_azienda= rs_consegna("d_azienda")
			d_Indirizzo=rs_consegna("d_Indirizzo")
			d_cap= rs_consegna("d_cap")
			d_citta=rs_consegna("d_citta")
			d_provincia=rs_consegna("d_provincia")
		end if

	end if

	%>
			<div id="dati-consegna">
			<strong>Azienda</strong>:<br>
		    <input name="d_azienda" type="text" value="<%=d_azienda%>" style="width:98%;">
		    <br>
		    <strong>Indirizzo</strong>:<br>
		    <input name="d_indirizzo" type="text" value="<%=d_indirizzo%>"  style="width:98%;" class="richiesto">
		    <br>
		    <strong>Cap</strong>:<br>
		    <input name="d_cap" type="text" value="<%=d_cap%>"  style="width:98%;">
		    <br>
		    <strong>Citt&agrave;</strong>:<br>
		    <input name="d_citta" type="text" value="<%=d_citta%>"  style="width:98%;">
		    <br>
		    <strong>Provincia</strong>:<br>
		    <input name="d_provincia" type="text"  style="width:98%;" value="<%=d_provincia%>" maxlength="2">
		    <%
			if idconsegna="" then idconsegna=0
			if iduser<>"" then
				set rs_consegna=conn.execute("select utenti_consegna.* from utenti_consegna where iduser="&iduser)
				if not rs_consegna.eof then
					response.write "Scegli indirizzo esistente<br><ul id=""elenco_consegna"">"
					response.write "<li><input type=""radio"" name=""consegna"" value="""" "
					if idconsegna=0 then response.write " checked "
					response.write ">&nbsp;Usa indirizzo di fatturazione o nuovi dati consegna</li>"
					do while not rs_consegna.EOF
						response.write "<li><input type=""radio"" name=""consegna"" value="""&rs_consegna("id")&""""
						if idconsegna=clng(rs_consegna("id")) then response.write " checked "
						response.write ">&nbsp;"&rs_consegna("d_azienda")&", "&rs_consegna("d_indirizzo")&", "&rs_consegna("d_citta")&"</li>"
					rs_consegna.MoveNext
					loop
				    response.write "</ul>"
				end if
			end if
		    %>
		</div>
	<%
end sub
sub tab_dati_ddt(idord,iduser)
	mostra_ddt=true
	if idddt>0 then
		set rs_ddt=conn.execute("select ddt.*, ordini.data, ordini.nord, ordini.verde from ddt inner join ordini on ddt.idord = ordini.idord where ddt.idord = "&idord)
		causale=rs_ddt("causale")
		porto=rs_ddt("porto")
		imballo=rs_ddt("imballo")
		colli=rs_ddt("colli")
		peso=rs_ddt("peso")
		incaricato=rs_ddt("incaricato")
		dimensione=rs_ddt("dimensione")
		idddt=rs_ddt("idddt")
		nddt=formatdatetime(rs_ddt("nord"),2)
		data=formatdatetime(rs_ddt("data"),2)
		da_fatturare=rs_ddt("da_fatturare")
		set rs_ddt = Nothing
	else
		data=date()
		colli="1"
		peso="1"
		da_fatturare=1
	end if
	if utente_andrea then response.write "idddt:"&idddt
	%>

	<div id="dati-ddt" style="<%=display_none%> border: 0px; padding-bottom:0px;">
		<input type="hidden" name="dati-ddt" value="dati-ddt">
		<input type="hidden" name="idddt" value="<%=idddt%>">
		<input type="hidden" name="tipofattura" value="">
		<input type="hidden" name="nord" value="<%=nord%>">
		<input type="hidden" name="parziale" id="parziale" value="<%=parziale%>">
			<table width="100%" border="0" cellpadding="" cellspacing="0" class="">
			<%if month(date())=1 then %>
			<tr>
				<td>
					Anno
				</td>
			<td>
			<%
				anno=year(date())
				annoprecedente=anno-1
			%>
			<input type="radio" name="anno" value="" checked><%=anno%>		<input type="radio" name="anno" value="<%=annoprecedente%>"><%=annoprecedente%>
				</td>
			</tr>
			<%end if %>
	            <tr>
	              <td valign="top">Data</td>
	              <td>
					<input type="text" name="data_documento" value="<%=data%>">
	                </td>
	            </tr>
	            <tr>
					<td valign="top">Causale</td>
					<td><select name="causale" class="richiesto" id="causale">
						<%Set ccausale_ddt = new cl_causale_ddt
		                  ccausale_ddt.stampa_option(causale)
		                  set ccausale_ddt = Nothing
					%>
					</select>
	            </tr>

	            <tr>
	              <td valign="top">Da fatturare</td>
	              <td>
		              <input type="checkbox" name="da_fatturare" id="da_fatturare" value="1" <%if da_fatturare=1 then response.write " checked"%>>

	                </td>
	            </tr>

	            <tr>
	              <td valign="top">Porto</td>
	              <td><select name="porto" id="porto" class="richiesto">
	                  <%
						for n=1 to porto_ddt(0)
							response.write "<option value='"&n&"'"
							if porto=n then response.write " selected style='color:red;'"
							response.write ">" & porto_ddt(n)&"</option>"
						next
						%>
	                </select>
	                </td>
	            </tr>
	            <tr>
	              <td valign="top">Imballo</td>
	              <td><select id="imballo" name="imballo" class="richiesto">
	                  <%
					for n=1 to imballo_ddt(0)
						response.write "<option value='"&n&"'"
						if imballo=n then response.write " selected style='color:red;'"
						response.write ">" & imballo_ddt(n)&"</option>"
					next
						%>
	                </select></td>
	            </tr>
	            <tr>
	              <td valign="top">Numero colli</td>
	              <td>
		              <input type="text" id="colli" name="colli" size="8" maxlength="3" value="<%=colli%>" class="richiesto">
		              </td>
	            </tr>
	            <tr>
	              <td valign="top">Peso</td>
	              <td>
	                <input type="text" id="peso" name="peso" size="8" maxlength="5" value="<%=peso%>" class="richiesto"></td>
	            </tr>
	            <tr>
	              <td valign="top">Dimensione</td>
	              <td>
	                <input type="text" id="dimensione" name="dimensione" size="50" maxlength="50" value="<%=dimensione%>"></td>
	            </tr>
	            <tr>
	              <td valign="top">Annotazioni</td>
	              <td><input type="text" name="annotazioni" size="50"  value=""></td>
	            </tr>
	            <tr>
	              <td valign="top">Incaricato del trasporto <%=incaricato%></td>
	              <td><select name="incaricato" class="richiesto">
	                  <%
						for n=1 to incaricato_ddt(0)
						response.write "<option value='"&n&"'"
						if incaricato=n then response.write "selected"
						response.write ">" & incaricato_ddt(n)&"</option>"
						next
						%>
	                </select>
	               </td>
	            </tr>
	            <%
		            set rs_user=conn.execute("select * from utenti where iduser="&iduser)
		            vettore=rs_user("note_trasporto_2")
		            set rs_user = Nothing
				%>
				 <tr>
	              <td valign="top">Vettore</td>
	              <td>
	                <input type="text" name="vettore" size="50" value="<%=vettore%>"></td>
	            </tr>
	            <tr>
	              <td valign="top">Data inizio trasporto</td>
	              <td><input type="text" name="data_inizio" size="12" maxlength="10" value="<%= formatDateTime(date(), vbShortDate)%>"></td>
	            </tr>


		</table>
	</div>
	<script>
		$(function() {
			$("#causale").change(function(){
				var fattura=$( "#causale option:selected" ).attr("data-fattura");
				console.log("fattura:"+fattura);
				$("#da_fatturare").prop('checked', fattura=="True");

			})
		});



		</script>



	<%
end sub



function aggiungi_intestazione(iduser,azienda,cognome,nome,cf,piva,indirizzo,citta,cap,provincia)

	azienda=trim(azienda)
	cognome=trim(cognome)
	nome=trim(nome)
	cf=trim(cf)
	piva=trim(piva)
	indirizzo=trim(indirizzo)
	citta=trim(citta)



	ricerca=replace(azienda&cognome&nome&cf&piva&indirizzo&citta,"'","''")

	'cerco negli indirizzi
	sql="select id from utenti_intestazioni where concat(azienda,cognome,nome,cf,piva,indirizzo,citta) = '"&ricerca&"' and iduser="&iduser
	sql2=sql
	call add2log(sql2,0)
	'response.write "<br>"&sql&"<br>"
	set rs2=conn.execute(sql)

	if rs2.eof then
		'cerco regione
		codice_regione=0
		set rs_regione=conn.execute("select codice_regione from elenco_province where Sigla_automobilistica='"&provincia&"'")
		if not rs_regione.eof then
			codice_regione=rs_regione("codice_regione")
		end if
		set rsrs_regione=Nothing
		set rs_consegna=Server.CreateObject("ADODB.Recordset")
		rs_consegna.open "utenti_intestazioni",conn,3,3
		rs_consegna.addnew
		rs_consegna("iduser")=iduser
		rs_consegna("azienda")=azienda
		rs_consegna("cognome")=cognome
		rs_consegna("nome")=nome
		rs_consegna("cf")=cf
		rs_consegna("piva")=piva
		rs_consegna("indirizzo")=indirizzo
		rs_consegna("citta")=citta
		rs_consegna("cap")=cap
		rs_consegna("provincia")=provincia
		if codice_regione>0 then
			rs_consegna("regione")=codice_regione
		end if
		rs_consegna.update
		aggiungi_intestazione=Get_last_id("utenti_intestazioni")
		rs_consegna.close
		set rs_consegna = Nothing
	else
		aggiungi_intestazione=rs2("id")
	end if
	set rs2 = Nothing
end function

sub  dati_fatturazione_consegna(rs)
		%>
          <tr>
            <td width="50%" bgcolor="#E5E5E5">Dati fatturazione</td>
            <td width="50%" bgcolor="#E5E5E5">Dati consegna</td>
          </tr>
          <tr >
            <td width="50%" valign="top" style="border-bottom:1px solid;">
	            <%=dati_fatturazione(rs("idintestazione"))%><br>
              Pagamento: <%=metodo_pagamento(rs("pagamento"),"descrizione")%>
              </td>
            <td width="50%" valign="top" style="border-bottom:1px solid;">
				<%=dati_consegna(rs("idconsegna"))%>
                <br />Metodo spedizione: <%=tipo_trasporto(rs("tipo_trasporto"))%>

              </td>
          </tr>
		<%
end sub


function dati_fatturazione(id)
	if id>0 then
		set rs_consegna=conn.execute("select * from utenti_intestazioni where id="&id)
		if rs_consegna.eof then
			dati_fatturazione="Dati fatturazione non trovati"
		else
			dati_fatturazione=""
			if utente_andrea then dati_fatturazione="("&id&") "
			dati_fatturazione=dati_fatturazione&"<b>"&denominazione(rs_consegna("nome"),rs_consegna("cognome"),rs_consegna("azienda"))&"</b><br>"
			if rs_consegna("Indirizzo")<>"" then
				dati_fatturazione=dati_fatturazione&rs_consegna("Indirizzo")&"<br>"
			end if
			if rs_consegna("cap")<>"" then
				dati_fatturazione=dati_fatturazione&rs_consegna("cap")&" - "
			end if
			dati_fatturazione=dati_fatturazione&rs_consegna("citta")
			if rs_consegna("provincia")<>"" then
				dati_fatturazione=dati_fatturazione&" ("&rs_consegna("provincia")&")"
			end if
			dati_fatturazione=dati_fatturazione&"<br>Codice Fiscale: "&rs_consegna("cf")&"<br>Partita Iva: "&rs_consegna("piva")
		end if
	else
		dati_fatturazione="id:"&"0"
	end if
end function
function dati_fatturazione_pdf(id)
	dati_fatturazione_pdf=""
	if id>0 then
		set rs_consegna=conn.execute("select * from utenti_intestazioni where id="&id)
		if rs_consegna.eof then
			dati_fatturazione_pdf="Dati fatturazione non trovati"
		else
			dati_fatturazione_pdf="<b>"&ucase(denominazione(rs_consegna("nome"),rs_consegna("cognome"),  rs_consegna("azienda")))&"</b><br>"
			dati_fatturazione_pdf=dati_fatturazione_pdf&rs_consegna("Indirizzo")&" - "
			if rs_consegna("cap")<>"" then
				dati_fatturazione_pdf=dati_fatturazione_pdf&rs_consegna("cap")&" - "
			end if
			dati_fatturazione_pdf=dati_fatturazione_pdf&rs_consegna("citta")
			if rs_consegna("provincia")<>"" then
				dati_fatturazione_pdf=dati_fatturazione_pdf&" ("&rs_consegna("provincia")&")"
			end if
			dati_fatturazione_pdf=dati_fatturazione_pdf&"<br>Codice Fiscale: "&rs_consegna("cf")&" - Partita Iva: "&rs_consegna("piva")
		end if
	else
		dati_fatturazione_pdf=""
	end if
end function
function get_fattura_pa(b_0_1,b_0_2)
	if b_0_1=1 then
		get_fattura_pa="PA"
	elseif b_0_2=1 then
		get_fattura_pa="CE"
	else
		get_fattura_pa=""
	end if
end function
function get_fattura_sp(b_0_1)
	if b_0_1=1 then
		get_fattura_sp="SP"
	else
		get_fattura_sp=""
	end if
end function



function anno_se_passato(anno)
	if anno<year(date()) then
		anno_se_passato="/"&anno
	else
		anno=""
	end if
end function

function get_sql_ordine_fattura(tipo_fattura,idfat)
			if tipo_fattura="ddt" then
				get_sql_ordine_fattura="select ordini_fatture.idfat, ordini_fatture.idord, ordini.idord as idddt, ordini.nord as nddt, ordini.data as dataddt, ordini.tipo_documento,ordini.trasporto as trasporto_ddt, ddt.tipo_ddt, ddt.sub_idord, o2.idord, o2.nord, o2.trasporto, o2.data as dataordine,  o2.sconto_ordine, o2.iva_ordine,o2.trattamento_iva_ordine, ddt.annotazioni from ordini_fatture inner join ordini on ordini_fatture.idord =ordini.idord left join ddt on ordini_fatture.idord = ddt.idord  left join ordini o2 on ddt.sub_idord = o2.idord where idfat="&idfat&" order by ddt.sub_idord"
			else
				get_sql_ordine_fattura="select ordini_fatture.idfat, ordini_fatture.idord, ordini.nord as idddt, ordini.nord as nddt, ordini.data as dataddt, ordini.tipo_documento, ddt.tipo_ddt, ddt.sub_idord, o2.idord, o2.nord, o2.data as dataordine, o2.trasporto, o2.sconto_ordine, o2.iva_ordine, o2.trattamento_iva_ordine, ddt.annotazioni from ordini_fatture inner join ordini on ordini_fatture.idord =ordini.idord left join ddt on ordini_fatture.idord = ddt.idord  left join ordini o2 on ordini_fatture.idord = o2.idord  where idfat="&idfat&" order by ddt.sub_idord"
			end if
end function







function get_codice_destinatario(CodiceDestinatario,FatturaPACodiceDestinatario)
	call add2log("CodiceDestinatario:"&CodiceDestinatario&"FatturaPACodiceDestinatario"&FatturaPACodiceDestinatario,0)
	              if not isnull(CodiceDestinatario) and CodiceDestinatario<>"" then
		              get_codice_destinatario=CodiceDestinatario
	          call add2log("A",0)
	          elseif not isnull(FatturaPACodiceDestinatario) and FatturaPACodiceDestinatario<>"" then
	          call add2log("B",0)
		          	arr=split(FatturaPACodiceDestinatario,vbcrlf)
		          	'if IsArray(arr) then
			          	tmp=arr(0)
			        'else
				    '    tmp=FatturaPACodiceDestinatario
				    'end if
					tmp2=split(tmp,":")
					'if IsArray(tmp2) then
						get_codice_destinatario=tmp2(0)
					'else

					'	get_codice_destinatario=tmp
					'end if

	          else
	          call add2log("C",0)
		          get_codice_destinatario=""

              end if
end function
sub dati_fatturazione_chkpivacf(id,iduser)
				if ordine.campo("stato_estero")<>1 then
				chkpiva=ControllaPIVA(trim(ordine.campo("piva")))
				chkcf=controllacf(trim(ordine.campo("cf")))
				if ordine.campo("cf")=ordine.campo("piva") then
					if chkcf="" or chkpiva="" then
						chkcf=""
						chkpiva=""
					end if

				end if
				end if

end sub
function anno_se_diverso(byval anno)
	anno_se_diverso=""
	if anno<>year(date()) then
		anno_se_diverso=" / "&anno
	end if
end function
function banca_appoggio(byval idbanca)
	dim array_tmp(2)
		array_tmp(0)=""
		Array_tmp(1)=""

		if idbanca="" or isnull(idbanca) or idbanca="0"  then
			set rsb=conn.execute("select * from banche where predefinita=1")
		else
			set rsb=conn.execute("select * from banche where idbanca="&idbanca)
		end if
		if not rsb.eof then
			array_tmp(0)=rsb("denominazione")
			Array_tmp(1)=rsb("iban")
		end if
		banca_appoggio=array_tmp
end function

%>
