<%
	
Class ClasseMagazzino
	Private c_idord
	Private c_idfor
	Private c_operazione
	Private c_txt_log
	Private idmag_no_indietro

    Private sub Class_Initialize
	    c_txt_log="Init<br>"
	    c_idord="NULL"
	    c_idfor="NULL"
	End Sub



    Private Sub Class_Terminate
	    
	End Sub
	
	
	Public sub scarica_ordine(idord)
	    c_txt_log=c_txt_log&"scarica_ordine("&idord&")<br>"
		c_idord=idord
		c_operazione=2
		
		sql="select ordini_dett.quantita as quantita_movimento, magazzino.quantita_magazzino, magazzino.db_ven,magazzino.db_acq,magazzino.idmag FROM magazzino INNER JOIN ordini_dett ON (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro) where ordini_dett.idpro>0 and ordini_dett.idord="&c_idord
		on error resume next
		set rs_magazzino=conn.execute (sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
		'response.end
		call scorri_magazzino(rs_magazzino,2,0)
		response.write c_txt_log
		'call add2log(c_txt_log,0)
	end sub


    Private sub scorri_magazzino(rs_magazzino,operazione,quantita_massima)
	    if operazione<>5 then
		    idmag_no_indietro=0
	    end if
	    c_txt_log=c_txt_log&"<b>scorri_magazzino(rs_magazzino, operazione:"&operazione&")</b><br>"
	    
	    'Elenco campi rs_magazzino
	    	'idmag
	    	'quantita_magazzino
	    	'quantita_movimento
	    	'db_ven
	    	'db_acq
	    
	    	'operazione:
			'1 Rettifica giacienza
			'2 Scarido magazzino per ordine
			'3 Carico magazzino
			'4 Carico magazzino per reso DDT
			'5 Scarico magazzino per ordine (DB)
			'6 Carico magazzino da ordine fornitore
			
	    
	    dim quantita_movimento, quantita_magazzino
	    if not rs_magazzino.eof then
		    do while not rs_magazzino.EOF
		    	quantita_magazzino=cdbl(rs_magazzino("quantita_magazzino"))
		    	quantita_movimento=cdbl(rs_magazzino("quantita_movimento"))
			    txt_prima="idmag:"&rs_magazzino("idmag")&"(quantita_magazzino:"&quantita_magazzino&", quantita_movimento:"&quantita_movimento&")"
	
		    	'Gestisco le quantita
		    	select case operazione
		    		case 2		'Scarico
		    			if quantita_movimento>quantita_magazzino then
			    			'Non posso scaricare più di quello che è a magazzino
			    			quantita_movimento=quantita_magazzino		    			
			    		end if
						'Inverto il segno per lo scarico
			    		quantita_movimento=-quantita_movimento
			    		quantita_magazzino=quantita_magazzino+quantita_movimento
		    		case 4,6	'Carico
			    		quantita_magazzino=quantita_magazzino+quantita_movimento
		    		case 1		'Imposto
			    		quantita_magazzino=quantita_movimento
			    	case 5		'Scarico i prodotti di acquisto e calcolo quantita massima
						quantita_scarico=quantita_movimento*rs_magazzino("quantita")
						'quantita_movimento
					    c_txt_log=c_txt_log&">>>>Trovato prodotto di acquisto (idmag:"&rs_magazzino("idmag")&", quantita_magazzino:"&rs_magazzino("quantita_magazzino")&")"
						
						if quantita_scarico>quantita_magazzino then
							quantita_scarico=quantita_magazzino
						end if
						
						'Inverto il segno per lo scarico
						'quantita_scarico=-quantita_scarico
						quantita_magazzino=quantita_magazzino+quantita_scarico
						
					
						quantita_tmp=quantita_magazzino/rs_magazzino("quantita")
						if quantita_massima>quantita_tmp or isnull(quantita_massima) then quantita_massima=quantita_tmp
					    c_txt_log=c_txt_log&"quantita_massima:"&quantita_massima&")<br>"
			   
		    		case else	
		    	End Select
			    c_txt_log=c_txt_log&txt_prima&" dopo (quantita_magazzino:"&quantita_magazzino&", quantita_movimento:"&quantita_movimento&")<br>"
		    	
		    	if rs_magazzino("db_ven")=1 then
			    	idmag_no_indietro=rs_magazzino("idmag")
				    c_txt_log=c_txt_log&">>>> <b>imposto</b> no_indietro prima di entrare:"& idmag_no_indietro&")<br>"
					call scorri_db_prodotti_acquisto(rs_magazzino("idmag"),quantita_movimento)
				else
					'Se non ha componenti scarico
					call aggiorna_magazzino(rs_magazzino("idmag"),quantita_magazzino )
					'if operazione<>5 then
						call inserisci_movimento(rs_magazzino("idmag"),quantita_movimento,operazione,idord,idfor)
					'end if
				end if
				c_txt_log=c_txt_log&"Elenco campi rs_magazzino:"
				For Each fld in rs_magazzino.fields
					c_txt_log=c_txt_log&fld.Name&","
				next

				on error resume next
				
				
				
			    c_txt_log=c_txt_log&"<b>Verifico idmag:"&rs_magazzino("idmag")&"</b> (db_acq:"&rs_magazzino("db_acq")&", idmag_no_indietro<>idmag_ven:"&idmag_no_indietro&"<>"&rs_magazzino("idmag_ven")&")<br>"
				if err.number<>0 then
					response.write c_txt_log&sql&"<br>"
				end if
				    
				    
				if rs_magazzino("db_acq")=1 and idmag_no_indietro<>clng(rs_magazzino("idmag_ven")) then
					quantita_massima=null
					call scorri_db_prodotti_vendita(rs_magazzino("idmag"),quantita_magazzino,quantita_massima)
				end if
		    
			    rs_magazzino.movenext
		    loop
		else
			'In rs_magazzino non ci sono record
		    c_txt_log=c_txt_log&"--In rs_magazzino non ci sono record<br>"
		end if
    end sub
    
    
    Private sub scorri_db_prodotti_acquisto(idmag_ven,quantita)
	    c_txt_log=c_txt_log&"scorri_db_prodotti_acquisto(idmag_ven:"&idmag_ven&", quantita:"&quantita&")<br>"
	    c_txt_log=c_txt_log&">>>> imposto no_indietro:"& idmag_no_indietro&")<br>"
	    dim quantita_magazzino, quantita_massima, quantita_scarico, quantita_tmp
		quantita_massima=null
	    
	    'Scorro i componenti di acquisto
    	sql="select distinta_base.*, "&quantita &" as quantita_movimento, magazzino.quantita_magazzino, magazzino.db_ven, magazzino.db_acq, magazzino.idmag FROM magazzino inner JOIN distinta_base ON distinta_base.idmag_acq = magazzino.idmag where distinta_base.idmag_ven="&idmag_ven
    	
    	
    	
		on error resume next
    	set rs_db = conn.execute ( sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
		operazione=5	'Scarico magazzino per ordine (DB)
    	call scorri_magazzino(rs_db,operazione,quantita_massima)
	    c_txt_log=c_txt_log&"Aggiorno magazzino del prodotto venduto(idmag_ven:"&idmag_ven&", quantita_massima:"&quantita_massima&")<br>"
	
		'Aggiorno magazzino del prodotto venduto
		call aggiorna_magazzino(idmag_ven,quantita_massima)    
    End Sub
    
    
    Private sub scorri_db_prodotti_vendita(idmag_acq, quantita_magazzino)
	    c_txt_log=c_txt_log&"scorri_db_prodotti_vendita(idmag_acq:"&idmag_acq&", quantita_magazzino:"&quantita_magazzino&")<br>"
		dim quantita_tmp, quantita_massima
	    quantita_massima=NULL
	    sql="select magazzino.idmag, distinta_base.*,  magazzino.quantita_magazzino,"&quantita_magazzino&" as quantita_movimento, magazzino.db_acq, magazzino.db_ven FROM magazzino INNER JOIN distinta_base ON distinta_base.idmag_ven =magazzino.idmag  where idmag_acq="&idmag_acq	    
	    'Scorro i componenti di vendita
		on error resume next
    	set rs_db = conn.execute ( sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
			response.end
		end if
		on error goto 0
	    
		operazione=5	'Scarico magazzino per ordine (DB)
    	call scorri_magazzino(rs_db,operazione,0)
    
    End Sub
    
    
	    
    Private sub aggiorna_magazzino(idmag,quantita)
	    dim sql, num
	    c_txt_log=c_txt_log&"<b>aggiorna_magazzino</b>(idmag:"&idmag&", quantita:"&quantita&")<br>"
		sql="UPDATE magazzino SET magazzino.quantita_magazzino = "&quantita&",data_aggiornamento=now() WHERE (((magazzino.idmag)="&idmag&"));"
		on error resume next
		conn.execute sql,num	
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
    end sub
    Private sub inserisci_movimento(idmag,quantita,operazione,idord,idfor)
	    c_txt_log=c_txt_log&"<b>inserisci_movimento</b>(idmag:"&idmag&", quantita:"&quantita&", operazione:"&operazione&", idord:"&idord&", idfor:"&idfor&")<br>"
	    dim sql, num
	    sql="INSERT INTO magazzino_movimenti (data,idmag,quantita,causale,iduser,idord,idfor) values (now(),"&idmag&","&metti_punto(quantita)&","&operazione&","&sessionIDUser&","&c_idord&","&c_idfor&");"
		on error resume next
		conn.execute sql,num	
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
    end sub


End Class
	
	
	sub aggiorna_magazzino_db(idpro_acq,idvara_liv1,idvarb_liv1,quantita)
	dim idpro,idvara,idvarb
	session("mag_db")=session("mag_db")&"<br>Prodotto acquistato:"&idpro_acq&"-"&idvara_liv1&"-"&idvarb_liv1
	'idpro ,idvara ,idvarb del prodotto acquistato
	'Risalgo la db e cerco i prodotti venduti
	'set rs_db_vend=conn.execute ("select distinta_base.* from distinta_base where idpro_acq="&idpro_acq&" and idvara_acq="&idvara_liv1&" and idvarb_acq="&idvarb_liv1)
	
	
set rs_db_vend=conn.execute ("select distinta_base.*, magazzino.db_acq FROM magazzino INNER JOIN distinta_base ON (magazzino.idpro = distinta_base.idpro_ven) AND (magazzino.idvara = distinta_base.idvara_ven) AND (magazzino.idvarb = distinta_base.idvarb_ven) where idpro_acq="&idpro_acq&" and idvara_acq="&idvara_liv1&" and idvarb_acq="&idvarb_liv1)
	
	
	if not rs_db_vend.eof then
		do while not rs_db_vend.eof 
		idpro=rs_db_vend("idpro_ven")
		idvarA=rs_db_vend("idvara_ven")
		idvarb=rs_db_vend("idvarb_ven")
		session("mag_db")=session("mag_db")&"<br>In db trovato prodotto venduto:"&idpro&"-"&idvara&"-"&idvarb&" SCANSIONO LA DB<br>"
		'Entro nella db e scansiono i prodotti di acquisto del prodotto di vendita
		dim quantita_tmp
		sql="select magazzino.idpro,magazzino.idvara,magazzino.idvarb,magazzino.quantita_magazzino, distinta_base.quantita, distinta_base.idpro_ven, distinta_base.idvara_ven, distinta_base.idvarb_ven ,magazzino.db_ven, magazzino.db_acq FROM distinta_base INNER JOIN magazzino ON (distinta_base.idpro_acq = magazzino.idpro) AND (distinta_base.idvarb_acq = magazzino.idvarb) AND (distinta_base.idvara_acq = magazzino.idvara) where distinta_base.idpro_ven="&idpro&" and idvara_ven="&idvarA&" and idvarb_ven="&idvarB
		'response.write sql
		set rs_db_acq=conn.execute(sql)
		do while not rs_db_acq.EOF
			session("mag_db")=session("mag_db")&"<br>In db trovato prodotto acquistato:"&rs_db_acq("idpro")&"-"&rs_db_acq("idvara")&"-"&rs_db_acq("idvarb")&" quantita a magazzino:"&rs_db_acq("quantita_magazzino")&"<br>"
				
		
				quantita_magazzino=rs_db_acq("quantita_magazzino")
				if isnull(quantita_magazzino) then
					quantita_magazzino=0
					quantita_tmp=0
				else
					quantita_tmp=quantita_magazzino/rs_db_acq("quantita")
				end if
				session("mag_db")=session("mag_db")&"coefficente utilizzo:"&rs_db_acq("quantita")&" quantita_tmp:"&quantita_tmp&"<br>"
				
				session("mag_db")=session("mag_db")&"quantita_massima prima:"&quantita_massima&"<br>"
				if quantita_massima>quantita_tmp or isnull(quantita_massima) then quantita_massima=quantita_tmp
				session("mag_db")=session("mag_db")&"quantita_massima dopo:"&quantita_massima&"<br>"
				
				
		rs_db_acq.MoveNext
		Loop
		set rs_db_acq=nothing
		'Trovato quantita massima per prodotto di vendita
		session("mag_db")=session("mag_db")&"<br>Imposto quantita calcolata:"&idpro&"-"&idvara&"-"&idvarb&" quantita:"&quantita_massima&"<br>"
		quantita_massima=replace(quantita_massima,",",".")
		sql="UPDATE magazzino SET magazzino.quantita_magazzino = "&quantita_massima&",data_aggiornamento=now() WHERE idpro="&idpro&" and idvara="&idvara&" and idvarb="&idvarb
		'session("mag_db")=session("mag_db")&"Query aggiornamento magazzino:"&sql&"<br>"
		conn.execute(sql	)
		session("mag_db")=session("mag_db")&"<br>Aggiornato magazzino prodotto di vendita:"&idpro&"-"&idvara&"-"&idvara&" quantità massima:"&quantita_massima
		
		
		
		
		
		if rs_db_vend("db_acq") then
			session("mag_db")=session("mag_db")&"<br>Richiamo ricorsivo aggiorna_magazzino_db quantita_magazzino:"&quantita_magazzino
			call aggiorna_magazzino_db(rs_db_vend("idpro_ven"),rs_db_vend("idvara_ven"),rs_db_vend("idvarb_ven"),quantita_magazzino)
		end if

		
		
		
		
		
		
		
		
		
		rs_db_vend.MoveNext
		Loop
	end if
	session("mag_db")=session("mag_db")&"<br>--------"

end sub

	
	
	
	
	
	
	
	
	
	sub scarica_distintabase(idpro,idvarA,idvarB,quantita,idmag)
	
	dim quantita_massima,quantita_tmp
	set rs_db = Server.CreateObject("ADODB.Recordset")
	if len(session("scarica_distintabase"))>10000 then
		session("scarica_distintabase")=session("scarica_distintabase")&"Arrestato"
		response.end
	end if
	session("scarica_distintabase")=session("scarica_distintabase")&"Chiamato scarica_distintabase"&now()&" per idord:"&idord&" per prodotto di vendita idpro:"&idpro&" idvara:"&idvarA&" idvarb:"&idvarb&" quantita:"&quantita&" idmag:"&idmag&"<br>"

	sql="select distinta_base.*, magazzino.quantita_magazzino, magazzino.db_ven,magazzino.idmag FROM magazzino RIGHT JOIN distinta_base ON (magazzino.idvarb = distinta_base.idvarb_acq) AND (magazzino.idvara = distinta_base.idvara_acq) AND (magazzino.idpro = distinta_base.idpro_acq) where distinta_base.idpro_ven="&idpro&" and idvara_ven="&idvarA&" and idvarb_ven="&idvarB
	set rs_db = conn.execute ( sql)
	causale=5	'Scarico magazzino per ordine (DB)
	quantita_massima=0
	do while not rs_db.EOF
		session("scarica_distintabase")=session("scarica_distintabase")&"<br>Trovato componente acquisto idpro:"&rs_db("idpro_acq")&" idvarA_acq:"&rs_db("idvarA_acq")&" idvarb_acq:"&rs_db("idvarb_acq")&"<br>"
			
		quantita_magazzino=rs_db("quantita_magazzino")
		if not isnull(quantita_magazzino) then
			quantita_scarico=quantita*rs_db("quantita")
			quantita_magazzino=quantita_magazzino-quantita_scarico
			
			conn.execute ("update magazzino set quantita_magazzino="&replace(quantita_magazzino,",",".")&" where idmag="&rs_db("idmag"))
			'rs_db.Update 
			quantita_tmp=quantita_magazzino/rs_db("quantita")
			if quantita_massima>quantita_tmp or quantita_massima=0 then quantita_massima=quantita_tmp
			quantita=replace(quantita,",",".")
			if rs_db("db_ven")=1 then
				session("scarica_distintabase")=session("scarica_distintabase")&"Chiamato ricorsivo "&now()&" "&rs_db("idpro_ven")&" "&rs_db("idvarA_ven")&" "&rs_db("idvarb_ven")&" "&quantita_scarico&" "&rs_db("idmag")&"<br>"
				call scarica_distintabase(rs_db("idpro_acq"),rs_db("idvarA_acq"),rs_db("idvarb_acq"),quantita_scarico,rs_db("idmag"))
			else
				quantita_scarico=replace (quantita_scarico,",",".")
				session("scarica_distintabase")=session("scarica_distintabase")&"Inserisco movimento magazzino nel prodotto di acquisto<br>"
				sql="INSERT INTO magazzino_movimenti (data,idpro,quantita,idvara,idvarb,causale,iduser,idord) values (now(),"&rs_db("idpro_acq")&",-"&replace(quantita_scarico,",",".")&","&rs_db("idvara_acq")&","&rs_db("idvarb_acq")&","&causale&","&session("iduser")&","&idord&");"
				'session("scarica_distintabase")=session("scarica_distintabase")&sql
				conn.execute sql,num
			end if
		else
			session("scarica_distintabase")=session("scarica_distintabase")&"<b>Salto perchè quantità è NULL</b><br>"

		end if
		rs_db.movenext
	Loop	
	set rs_db=nothing
	
	'Aggiorno magazzino del prodotto venduto
	quantita_massima=replace(quantita_massima,",",".")
	if idmag<>"" then
		conn.execute("UPDATE magazzino SET magazzino.quantita_magazzino = "&quantita_massima&",data_aggiornamento=now() WHERE (((magazzino.idmag)="&idmag&"));"	)
		session("scarica_distintabase")=session("scarica_distintabase")&"Aggiorno quantità massima nel prodotto di vendita quantita:"&quantita_massima&"<br>"
	else
		session("scarica_distintabase")=session("scarica_distintabase")&"<b>Problema idmag nullo</b><br>"
		add2log "idpro:"&idpro&" idvara:"&idvara&" idvarb:"&idvarb&" errore in scarica_distintabase, idmag è nullo",3
	end if
end sub

	
	
	
	
%>