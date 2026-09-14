<!--#include virtual="/md5-2.asp"-->

<%
'Sintassi
dim permessi_utente
permessi_utente=array("A1","B1")






Class ClasseUtente
	Private m_log_classe
	Private m_log
	Private rs_utente
	Private p_oper
	Private modalita_registrazione
	Private registrato_da_iduser
	Private iduser
	Private p_errore
	Private m_secondario
	Dim modificheRS
	
	
	
	
    Public Default Function Init(parameters)
	    dim oper
	    m_log_classe="Init<br>"
	    m_log=false
	    p_errore=""
	    Set Init = Me
		if isarray(parameters) then
			n_parametri=UBound(parameters)+1
			p_oper=lcase(parameters(0))
			iduser=lcase(parameters(1))

		else
			p_oper=parameters
		end if

        if n_parametri>2 then
			iduser=parameters(1)
			da_dove=parameters(2)
		end if
		
		set rs_utente=Server.CreateObject("ADODB.Recordset")
		
	    select case p_oper
	    	case "add"
	    	'	set utente= (new Classeutente)("add")

	    		modalita_registrazione=1	'Da Admin
	    		registrato_da_iduser=sessioniduser
				call creanuovorecord()
				call aggiornarecord()
				
	    	case "add_da_user"
	    		modalita_registrazione=2	'Da modulo registrazione
	    		registrato_da_iduser=0
				call creanuovorecord()
				call aggiungidatiform()
				
	    	case "add_da_dialog"
	    	'	set utente= (new Classeutente)("add")

	    		modalita_registrazione=1	'Da Admin
	    		registrato_da_iduser=sessioniduser
				call creanuovorecord()
				call aggiungidatidialog()

			case "update"
	    	'	set utente= (new Classeutente)(array("update",iduser))

				call aprirecord()
				call aggiornarecord()
				
	    	case "update_da_user"
	    	'	set utente= (new Classeutente)(array("update_da_user",iduser))

	    	
				call aprirecord()
				call aggiornarecordform()
			
	    	case "delete"
	    	
				call aprirecord()
				call elimina_utente()
				
			case "new_password"
		    	'	set utente= (new Classeutente)(array("new_password",iduser,"password_md5"))

		        call add_new_password(parameters(2))
				
				
			
		end select        

    End Function
    
    Private Sub Class_Terminate
	    if m_log then call add2log(m_log_classe,0)
		Set rs_utente = Nothing
	End Sub
	
		''case -1
		''modalitaRegistrazione=4
		''case 1
		''modalitaRegistrazione="Da Admin"
		''case 2
		''modalitaRegistrazione="Da modulo registrazione"
		''case 3
		''modalitaRegistrazione="Da completamento ordine"
		''case 4
		''modalitaRegistrazione="Importato"
		''case else
		''modalitaRegistrazione="Da modulo registrazione"
	
    Private Sub aprirecord()
		rs_utente.open "select * from utenti where iduser="&iduser,conn,1,3
		set modificheRS= (new ClasseModificheRS)("add")
		modifichers.leggi(rs_utente)	
		m_secondario=rs_utente("secondario")
	End Sub
	
    Private Sub creanuovorecord()
	    m_secondario=0
		rs_utente.open "select * from utenti",conn,3,3
		rs_utente.addnew
		set modificheRS= (new ClasseModificheRS)(p_oper)
		modifichers.leggi(rs_utente)	
		rs_utente("data")=now()
		rs_utente("registrato_da_iduser")=registrato_da_iduser
		rs_utente("modalita_registrazione")=modalita_registrazione
		call add2log("idagente:"&session("idagente"),0)
		if session("idagente")<>"" then
			rs_utente("idagente")=session("idagente")
		end if
		num_fatture=0
		
	End Sub

	Private sub aggiornarecord()
		
		if request.form("tipo")="P" then
			azienda=""
		else
			azienda=AllFirstUp(trim(request.form("azienda")))
		end if
		
		rs_utente("email")=trim(request.form("email"))
		rs_utente("cognome")=AllFirstUp(trim(request.form("cognome")))
		rs_utente("nome")=AllFirstUp(trim(request.form("nome")))
		rs_utente("ruolo")=AllFirstUp(trim(request.form("ruolo")))
		'rs_utente("Indirizzo")=trim(request.form("Indirizzo"))
		'rs_utente("citta")=allfirstup(trim(request.form("citta")))
		'rs_utente("cap")=trim(request.form("cap"))
		'rs_utente("provincia")=trim(request.form("provincia"))
		'rs_utente("regione")=trim(request.form("regione"))
		rs_utente("telefono")=trim(request.form("telefono"))
		rs_utente("fax")=trim(request.form("fax"))
		rs_utente("cellulare")=trim(request.form("cellulare"))
		'rs_utente("piva")=trim(request.form("piva"))
		'rs_utente("cf")=ucase(trim(request.form("cf")))
		'rs_utente("d_Azienda")=trim(request.form("d_azienda"))
		'rs_utente("d_Indirizzo")=trim(request.form("d_Indirizzo"))
		'rs_utente("d_citta")=trim(request.form("d_citta"))
		'rs_utente("d_cap")=trim(request.form("d_cap"))
		'rs_utente("d_provincia")=ucase(trim(request.form("d_provincia")))
		rs_utente("note_su_utente")=trim(request.form("note_su_utente"))
		if m_secondario=0 then
			rs_utente("pag_accordato")=request.form("pag_accordato")
			rs_utente("banca_appoggio")=ucase(trim(request.form("banca_appoggio")))
			rs_utente("iban")=ucase(trim(request.form("iban")))
			rs_utente("tipo_trasporto_2")=request.form("tipo_trasporto_2")
			rs_utente("note_trasporto_2")=trim(request.form("note_trasporto_2"))
			
			if request.form("spese_0")="si" then
				rs_utente("spese_0")=1
			else
				rs_utente("spese_0")=0
			end if
			'Permesso per agente
			'if true then
					'if session("idagente")="" then

				rs_utente("idagente")=request.form("idagente")
				'end if
			'end if
			rs_utente("trattamento_iva")=request.form("trattamento_iva")
			rs_utente("idbanca")=request.form("idbanca")
			rs_utente("sito_suggerito")=cerca_sito(trim(request.form("email")))
			rs_utente("voto")=request.form("voto")
			rs_utente("tipologia")=replace(request.form("tipologia")," ","")
			if request.form("fornitore")="si" or num_fatture>0 then
				rs_utente("fornitore")=1
			else
				rs_utente("fornitore")=0
			end if
		end if
		if request.form("vedi_prezzi")="si" then
			rs_utente("vedi_prezzi")=1
		else
			rs_utente("vedi_prezzi")=0
		end if
		
		if request.form("non_eliminare")="si" then
			rs_utente("non_eliminare")=1
		else
			rs_utente("non_eliminare")=0
		end if
		if request.form("fattura_sp")="si" then
			rs_utente("fattura_sp")=1
		else
			rs_utente("fattura_sp")=0
		end if
		
		
		if request.form("selezmailing")="si" then
			rs_utente("selezmailing")=1
		else
			rs_utente("selezmailing")=0
		end if
		
		
		permessi_txt=""
		for n=0 to UBound(permessi_utente)
			if request.form(permessi_utente(n))="si" then
				call concatena_stringa( permessi_txt,",",permessi_utente(n)&"=1")
			end if
		next
		rs_utente("permessi_utente")=permessi_txt
		
		bFornitore=rs_utente("fornitore")
		rsnome=AllFirstUp(trim(request.form("nome")))
		rscognome=AllFirstUp(trim(request.form("cognome")))
		modifiche_rs=modificheRS.confronta(rs_utente)
		rs_utente.update
		
		if p_oper="add" then
			iduser=get_last_id("utenti")
		end if
		rs_utente.close
		if m_secondario=0 then
			idintestazione= aggiungi_intestazione(iduser,azienda,rscognome,rsnome,ucase(trim(request.form("cf"))),trim(request.form("piva")),trim(request.form("Indirizzo")),allfirstup(trim(request.form("citta"))),trim(request.form("cap")),trim(request.form("provincia")))
			conn.execute ("UPDATE utenti set idintestazione="&idintestazione&" where iduser="&iduser)
		end if
		'if oper="update" then 'Confronto recordset dopo modifiche
			
		'if modifiche_rs="" then modifiche_rs="Nessuna modifica"
		'end if
		if oper="add" then
			txt= "Aggiunto"
		else
			txt= "Modifica"
		end if
		Set rs_utente = Nothing
		
		if bFornitore then
			
			Set rs_for = Server.CreateObject("ADODB.Recordset")
			rs_for.Open "select utenti_fornitori.* from utenti_fornitori where iduser="&iduser, conn, 3, 3
			if rs_for.eof then
				rs_for.addnew
				rs_for("iduser")=iduser
			end if
			'memorizzo situazione iniziale recordset successivo
			modifichers.leggi(rs_for)
			'Modifico recordset
			if request.form("importo_minimo")="" then rs_for("importo_minimo")=0 else rs_for("importo_minimo")=aggiusta_decimale(request.form("importo_minimo"),"asp")
			modifiche_rs=modificheRs.confronta(rs_for)
			rs_for.update
			rs_for.close
			set rs_for=nothing
		end if
		if m_secondario=0 then
			Set rs_cliente = Server.CreateObject("ADODB.Recordset")
			rs_cliente.Open "select utenti_clienti.* from utenti_clienti where iduser="&iduser, conn, 1, 3
			if rs_cliente.eof then
				rs_cliente.addnew
				rs_cliente("iduser")=iduser
			end if
			
			modificheRS.leggi(rs_cliente)
			if trim(request.form("FatturaPACodiceDestinatario"))="" then
				rs_cliente("FatturaPACodiceDestinatario")=""
			else
				codicedestinatario=""
				tmp=split(trim(request.form("FatturaPACodiceDestinatario")),vbcrlf)
				for n=0 to ubound(tmp)
					
					if instr(tmp(n),":")>0 then 'Ci sono i 2 punti
						tmp2=split(tmp(n),":")
						riga=ucase(trim(tmp2(0)))&":"&trim(tmp2(1))
					else 						'Non ci sono i 2 punti
						riga=ucase(tmp(n))
					end if
					if codicedestinatario<>"" then codicedestinatario=codicedestinatario&vbcrlf
					codicedestinatario=codicedestinatario&riga
				next
				rs_cliente("FatturaPACodiceDestinatario")=codicedestinatario
			end if
			rs_cliente("email_inoltro_fattura")=request.form("email_inoltro_fattura")
			rs_cliente("tipo")=request.form("tipo")
			if request.form("stato_estero")="si" then
				rs_cliente("stato_estero")=1
			else
				rs_cliente("stato_estero")=0
			end if
			
			'Confronto recordset successivo dopo modifico
			modifiche_rs=modificheRS.confronta(rs_cliente)
			rs_cliente.update
			rs_cliente.close
		end if
		set rs_cliente=nothing
		valore_add2log=2

		if modifiche_rs="" then
			modifiche_rs="Nessuna modificas"
			valore_add2log=1
		end if

		add2log txt&" [utente="&iduser&"]"&denominazione(rsnome,rscognome,azienda)&"[/utente] eseguita da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&modifiche_rs,valore_add2log
	End Sub
	
	Private sub aggiungidatiform()
	
	
		'Dati form utente
		secondario=0
		nome=AllFirstUp(trim(request.form("nome")))
		cognome=AllFirstUp(trim(request.form("cognome")))
		azienda=AllFirstUp(trim(request.form("azienda")))
		cf=trim(request.form("cf"))
		if request.form("tipo_form")="pa" then
			'Cerco partita iva
			sql="select * from utenti_intestazioni i inner join utenti u on i.id = u.idintestazione where cf='"&cf&"'"
			set rs_intestazioni=conn.execute(sql)
			if not rs_intestazioni.eof then
				rs_utente("idintestazione")=rs_intestazioni("id")
				secondario=rs_intestazioni("iduser")
				rs_utente("secondario")=secondario
			end if
			call add2log(sql&" secondario:"&secondario,0)
		end if
		
		rs_utente("email")=lcase(trim(request.form("email")))
		rs_utente("1email")=lcase(trim(request.form("email")))
		rs_utente("telefono")=trim(request.form("telefono"))
		rs_utente("cognome")=cognome
		rs_utente("nome")=nome
		rs_utente("ruolo")=FirstUp(trim(request.form("ruolo")))
		if session("lingua")<>"" then rs_utente("lingua")=replace(session("lingua"),"_","")
		if request.form("mailing")="Si" then rs_utente("mailing")=1 else rs_utente("mailing")=0 end if
		
		'Altri dati
		rs_utente("permessi_utente")=permessi_utente_default()
		rs_utente("HTTP_USER_AGENT")=Request.ServerVariables("HTTP_USER_AGENT")&"<br>IP:"&Request.ServerVariables("REMOTE_ADDR")
		modifiche_rs=modificheRS.confronta(rs_utente)
		rs_utente.update
		
		iduser=get_last_id("utenti")
		rs_utente.close
		if secondario=0 then
			'Aggiungo intestazione
			idintestazione= aggiungi_intestazione(iduser,azienda,cognome,nome,cf,"","","","","")
			'idintestazione=aggiungi_intestazione(iduser,azienda,cognome,nome,cf,piva,indirizzo,citta,cap,provincia)
			conn.execute ("UPDATE utenti set idintestazione="&idintestazione&" where iduser="&iduser)
			txt_secondario="<b> PRINCIPALE </b>"
		else
			txt_secondario="<b> SECONDARIO </b>"
		end if
		call add2log ("Registrazione [utente="&iduser&"]"&denominazione(nome,cognome,azienda)&"[/utente]"&txt_secondario&vbcrlf&queryeform(),2)
		logga "reg",iduser
	end sub
	
	Private sub aggiungidatidialog()
		
		'Dati form utente
		nome=trim(request.form("nome"))
		cognome=trim(request.form("cognome"))
		azienda=trim(request.form("azienda"))
		cf=trim(request.form("cf"))
		piva=trim(request.form("piva"))
		indirizzo=trim(request.form("indirizzo"))
		cap=trim(request.form("cap"))
		citta=trim(request.form("citta"))
		provincia=trim(request.form("provincia"))
		
		'rs_utente("email")=lcase(trim(request.form("email")))
		'rs_utente("1email")=lcase(trim(request.form("email")))
		'rs_utente("telefono")=trim(request.form("telefono"))
		
		'Altri dati
		modifiche_rs=modificheRS.confronta(rs_utente)
		rs_utente.update
		
		iduser=get_last_id("utenti")
		rs_utente.close
		'Aggiungo intestazione (iduser,azienda,cognome,nome,cf,piva,indirizzo,citta,cap,provincia)
		idintestazione= aggiungi_intestazione(iduser,azienda,cognome,nome,cf,piva,indirizzo,citta,cap,provincia)
		conn.execute ("UPDATE utenti set idintestazione="&idintestazione&" where iduser="&iduser)
		
		call add2log ("Registrazione [utente="&iduser&"]"&denominazione(nome,cognome,azienda)&"[/utente]"&vbcrlf&queryeform(),2)
		logga "reg",iduser
	end sub
	Private function elimina_utente_strong()
	Cancella=true
	'controllo ordini non eliminati
    lngCount=clng(conn.execute("select count(*) FROM ordini where eliminato=0 and iduser="& iduser )(0))
    if lngCount>0 then
		cancella=false
		errore="Ci sono elementi negli ordini"
	end if
	'controllo preventivi non eliminati	
    lngCount=clng(conn.execute("select count(*) FROM preventivi where eliminato=0 and iduser="& iduser )(0))
    if lngCount>0 then
		cancella=false
		errore="Ci sono elementi nei preventivi"
	end if
	if cancella=true then
		'elimina carrello
		sql="DELETE  FROM carrello WHERE iduser=" & iduser 
		conn.execute(sql)
		'elimina preferiti
		sql="DELETE  FROM preferiti WHERE iduser=" & iduser
		conn.execute(sql)
		'elimina note su utente
		sql="DELETE  FROM note_su_utenti WHERE iduser=" & iduser
		conn.execute(sql)
		'elimina dettaglio_ordini_originale e dettaglio ordini
		set rs_ordini=conn.execute("select * from ordini where iduser=" & iduser)
		do while not rs_ordini.eof
			sql="DELETE  FROM ordini_dett WHERE idord=" & rs_ordini("idord")
			conn.execute(sql)
			sql="DELETE  FROM ordini_dett_originale WHERE idord=" & rs_ordini("idord")
			conn.execute(sql)
			rs_ordini.MoveNext
		loop
		rs_ordini.Close
		'elimina ordini
		sql="DELETE  FROM ordini WHERE iduser=" & iduser
		conn.execute(sql)
		'elimina preventivi e dettaglio preventivi
		set rs_ordini=conn.execute("select * from preventivi where iduser=" & iduser)
		do while not rs_ordini.eof
			sql="DELETE  FROM preventivi_dett WHERE idord=" & rs_ordini("idord")
			conn.execute(sql)
			rs_ordini.MoveNext
		loop
		rs_ordini.Close
		set rs_ordini=nothing
		'elimina preventivi
		sql="DELETE  FROM preventivi WHERE iduser=" & iduser
		conn.execute(sql)
		'elimina utenti_fornitori
		sql="DELETE  FROM utenti_fornitori WHERE iduser=" & iduser
		conn.execute(sql)
		'elimina utenti_clienti
		sql="DELETE  FROM utenti_clienti WHERE iduser=" & iduser
		conn.execute(sql)
		'elimina utente
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "select * FROM utenti where iduser="& iduser , conn, 3,2
		n=rs.fields.count-1
		Set Flds = rs.Fields
		'response.write "n:"&n
		'memorizzo i nomi dei campi e i valori attuali
		txt_add2log=""
		for i=0 to n
			'response.write "i:"&i
			if rs(i)<>"" then txt_add2log=txt_add2log&RS.Fields(i).Name&":"&rs(i)&vbcrlf
		next
		add2log "Eliminato [utente="&iduser&"]"&nominativo(rs("nome"),rs("cognome"),rs("azienda"))&"[/utente] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&txt_add2log,3
		rs.delete
		rs.close
		response.redirect questofile
		oper="list"
		
		Set rs = Nothing
	end if
	
		
		
		
	End function
	
	
	Public function elimina_utente(iduser)
	dim eliminabile, rs_utente, txt
	elimina_utente=""
	eliminabile=""
	txt=""
	secondario=conn.execute("select secondario from utenti where iduser="&iduser)(0)
	'Controllo se eliminabile per relazioni in altre tabelle
	sql="select count(*) from utenti_dipendenti where iduser="&iduser
	n=clng(conn.execute(sql)(0))
	if n>0 then
		call concatena_stringa(eliminabile,"<br>","Sono presenti "&n&" dipendenti")
	end if
	sql="select count(*) from ordini where eliminato=0 and iduser="&iduser
	n=clng(conn.execute(sql)(0))
	if n>0 then
		
		
		call concatena_stringa(eliminabile,"<br>","Sono presenti "&n&" ordini")
	end if
	sql="select count(*) from preventivi where iduser="&iduser
	n=clng(conn.execute(sql)(0))
	if n>0 then
		call concatena_stringa(eliminabile,"<br>","Sono presenti "&n&" preventivi")
	end if
	sql="select count(*) from ordini_fornitori where iduser="&iduser
	n=clng(conn.execute(sql)(0))
	if n>0 then
		call concatena_stringa(eliminabile,"<br>","Sono presenti "&n&" ordini fornitori")
	end if
	sql="select count(*) from fatture_for where idfor="&iduser
	n=clng(conn.execute(sql)(0))
	if n>0 then
		call concatena_stringa(eliminabile,"<br>","Sono presenti "&n&" fatture fornitore")
	end if
	sql="select count(*) from prodotti where idfor="&iduser
	n=clng(conn.execute(sql)(0))
	if n>0 then
		call concatena_stringa(eliminabile,"<br>","Sono presenti "&n&" prodotti forniti")
	end if
		
		
	set rs_utente=conn.execute ("select * from utenti where iduser="&iduser)
	if request.form("elimina_utente")="" then
		'Controllo se eliminabile per relazioni impostazioni utente
		if rs_utente("fornitore")=1 then
			call concatena_stringa(eliminabile,"<br>","L'utente &egrave; un fornitore")
		end if
		if rs_utente("non_eliminare")=1 then
			call concatena_stringa(eliminabile,"<br>","L'utente &egrave; impostato non eliminabile")
		end if
	end if
	
	if eliminabile="" then
		'recupero denominazione
		set rs_i=conn.execute("select * from utenti_intestazioni inner join utenti on utenti_intestazioni.id = utenti.idintestazione where utenti.iduser="&iduser)
		if rs_i.eof then
			denominazione_v="Utente senza intestazione"
		else
			denominazione_v=denominazione(rs_i("nome").value,rs_i("cognome").value,rs_i("azienda").value)
		end if
			
		set rs_i = Nothing
		sql="DELETE  from carrello  WHERE iduser=" & iduser
		conn.execute(sql)
		sql="DELETE  from preferiti  WHERE iduser=" & iduser
		conn.execute(sql)
	
		'Cancello record in tabelle correlate
		sql="DELETE FROM utenti_clienti WHERE iduser = " & iduser
		conn.execute(sql)
		sql="DELETE FROM utenti_fornitori WHERE iduser = " & iduser
		conn.execute(sql)
		if secondario=0 then
			sql="DELETE FROM utenti_intestazioni WHERE iduser = " & iduser
			conn.execute(sql)
			sql="DELETE FROM utenti_consegna WHERE iduser = " & iduser
			conn.execute(sql)
		end if
		sql="DELETE FROM utenti_dipendenti WHERE iduser = " & iduser
		conn.execute(sql)

		
		'elimina dettaglio_ordini_originale e dettaglio ordini
		set rs_ordini=conn.execute("select * from ordini where iduser=" & iduser)
		do while not rs_ordini.eof
			sql="DELETE  FROM ordini_dett WHERE idord=" & rs_ordini("idord")
			conn.execute(sql)
			sql="DELETE  FROM ordini_dett_originale WHERE idord=" & rs_ordini("idord")
			conn.execute(sql)
			rs_ordini.MoveNext
		loop
		rs_ordini.Close
		'elimina ordini
		sql="DELETE  FROM ordini WHERE iduser=" & iduser
		conn.execute(sql)
		
		
		txt=txt&"Elimino utente "
		n=rs_utente.fields.count-1
		for i=0 to n
			if (rs_utente(i)<>"") and (instr(rs_utente.Fields(i).Name,"md5")=0) then txt=txt&rs_utente.Fields(i).Name&"="&rs_utente(i)&"| "
		next
		call concatena_stringa( elimina_utente,"<br>",txt)
		
		sql="DELETE FROM utenti WHERE iduser = " & iduser
		conn.execute(sql)
	else
		
		elimina_utente="Impossibile eliminare utente id:"& iduser&" "&denominazione_v&" per:<br>"&  eliminabile
		
	end if
	set rs_utente = Nothing
	call add2log ("Eliminato [utente="&iduser&"]"&denominazione_v&"[/utente] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&elimina_utente,3)
	
	

		
		
		
		
	End function
	
	
	Private function aggiornarecordform()
		if rs_utente("secondario")=0 then
			'Campi per utenti NON SECONDARI
			if request.form("piva")<>"" or request.form("cf")<>"" then
				sql_controllo="( "
				if request.form("cf")<>"" then controllo_cf=" utenti_intestazioni.cf='"& request.form("cf")&"'"
				if request.form("piva")<>"" then controllo_piva=" utenti_intestazioni.piva='"& request.form("piva")&"'"
				if controllo_cf<>"" and controllo_piva<>"" then controllo_or=" or "
				sql_controllo=sql_controllo&controllo_cf&controllo_or&controllo_piva
				sql_controllo=sql_controllo&")" 
				sql_controllo="select * from utenti_intestazioni inner join utenti on utenti_intestazioni.id=utenti.idintestazione where "&sql_controllo
				sql_controllo=sql_controllo&" and utenti_intestazioni.iduser<>"&iduser
				set rs=conn.execute(sql_controllo)
				if not rs.eof then error="PIVA o CF gi&agrave; utilizzato da un altro utente."
				rs.close
				set rs=nothing
			end if
			if error="" then
				if request.form("piva_editabile")="SI" then
					cf=ucase(trim(request.form("cf")))
					piva=ucase(trim(request.form("piva")))
				else
					'Recupero dati precedenti	
					set rs=conn.execute ("select cf,piva from utenti_intestazioni where id="&rs_utente("idintestazione"))
					cf=rs("cf")
					piva=rs("piva")
					set rs = nothing
				end if
			idintestazione= aggiungi_intestazione(iduser,AllFirstUp(trim(request.form("azienda"))),AllFirstUp(trim(request.form("cognome"))),AllFirstUp(trim(request.form("nome"))),cf,piva,trim(request.form("Indirizzo")),allfirstup(trim(request.form("citta"))),trim(request.form("cap")),trim(request.form("provincia")))
				rs_utente("idintestazione")=idintestazione
		
				
			end if
		end if
		if error="" then
		
			rs_utente("cognome")=AllFirstUp(trim(request.form("cognome")))
			rs_utente("nome")=AllFirstUp(trim(request.form("nome")))
			rs_utente("ruolo")=trim(request.form("ruolo"))
			rs_utente("telefono")=trim(request.form("telefono"))
			rs_utente("lingua")=replace(session("lingua"),"_","")
			'rs_utente("cellulare")=request.form("cellulare")
			modifiche_rs=modificheRS.confronta(rs_utente)
			rs_utente.update
			rs_utente.close
			set rs_utente=nothing
			
		
		
			add2log "L'utente [utente="&iduser&"]"&session("nominativo")&"[/utente] ha modificato i propri dati"&vbcrlf&modifiche_rs,1
		end if

		
	end function
	
	Private sub add_new_password(passwordInChiaro)
			dim h
			set h = new MD5

			dim temp_passw
			set rs_password= Server.CreateObject("ADODB.Recordset")
			rs_password.Open "select * from password_reset where iduser="&iduser, conn, 1, 3
			if rs_password.eof then
				rs_password.addnew
				rs_password("iduser")=iduser
			end if
			rs_password("ora_invio")=now()
			rs_password("newpassword")=h.hash(passwordInChiaro)
			rs_password.update
			rs_password.close
			set rs_password = nothing
			call add2log("Impostato password iduser:"&iduser,0)
	end sub
	
	
	Public Property Get get_iduser()
		get_iduser=iduser
	end property
	Public Property Get errore()
		errore=p_errore
	end property

	
End Class

function permessi_utente_default()
	permessi_utente_default="A1=1,B1=1"
end function

%>