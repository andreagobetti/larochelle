<%
	
Class ClasseMagazzino
	Private c_idord
	Private c_idfor
	Private c_operazione
	Private c_txt_log
	Private sql

    Private sub Class_Initialize
	    c_txt_log="Init<br>"
	    c_idord="NULL"
	    c_idfor="NULL"
	End Sub



    Private Sub Class_Terminate
	    call add2log(c_txt_log,0)
	    'response.write c_txt_log
	End Sub
	
	
	Public sub scarica_ordine(idord)
	    c_txt_log=c_txt_log&"scarica_ordine("&idord&")<br>"
		c_idord=idord
	    c_idfor="NULL"
		c_operazione=2
		
		sql="select ordini_dett.quantita as quantita_movimento,ordini_dett.iddett,ordini_dett.quantita_scalato_magazzino, magazzino.quantita_magazzino, magazzino.db_ven,magazzino.db_acq,magazzino.idmag FROM magazzino INNER JOIN ordini_dett ON (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro) where quantita>quantita_scalato_magazzino and ordini_dett.idpro>0 and ordini_dett.idord="&c_idord
		call scorri_magazzino(sql,c_operazione)
		'call add2log(c_txt_log,0)
	end sub
	
	Public sub carica_ordine_fornitore(idord)
	    c_txt_log=c_txt_log&"carica_ordine_fornitore("&idord&")<br>"
		c_idfor=idord
	    c_idord="NULL"

		c_operazione=3
		
		sql="select ordini_fornitori_dett.quantita as quantita_movimento, magazzino.quantita_magazzino, magazzino.db_ven,magazzino.db_acq,magazzino.idmag FROM magazzino INNER JOIN ordini_fornitori_dett ON (magazzino.idvarb = ordini_fornitori_dett.idvarb) AND (magazzino.idvara = ordini_fornitori_dett.idvara) AND (magazzino.idpro = ordini_fornitori_dett.idpro) where ordini_fornitori_dett.idpro>0 and ordini_fornitori_dett.idord="&c_idfor
		'response.end
		call scorri_magazzino(sql,c_operazione)
		'call add2log(c_txt_log,0)
	end sub

	Public sub aggiorna_da_form()
	    c_txt_log=c_txt_log&"aggiorna_da_form()<br>"
		c_operazione=1
		
		sql="select magazzino.*, 0 as quantita_movimento FROM magazzino where  magazzino.idmag IN ("&request("chg")&")"
				
		call scorri_magazzino(sql,c_operazione)
		'call add2log(c_txt_log,0)
	end sub
	
	Public sub reso_da_ddt(stringa,idord)
		c_idord=idord
	    c_idfor="NULL"
	    c_txt_log=c_txt_log&"reso_da_ddt("&stringa&","&idord&")<br>"
		c_operazione=4
		if stringa<>"" then
			sql="select ordini_dett.quantita as quantita_movimento, magazzino.quantita_magazzino, magazzino.db_ven,magazzino.db_acq,magazzino.idmag FROM magazzino INNER JOIN ordini_dett ON (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro) where ordini_dett.iddett IN ("&stringa&")"

					
			call scorri_magazzino(sql,c_operazione)
			'call add2log(c_txt_log,0)
		end if
	end sub
	
	Public sub riserva_per_ordine(idord)
	    c_txt_log=c_txt_log&"riserva_per_ordine("&idord&")<br>"
		c_idord=idord
	    c_idfor="NULL"
		c_operazione=6
		
		sql="select ordini_dett.quantita as quantita_movimento,ordini_dett.iddett,ordini_dett.quantita_scalato_magazzino, magazzino.quantita_magazzino, magazzino.db_ven,magazzino.db_acq,magazzino.idmag FROM magazzino INNER JOIN ordini_dett ON (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro) where quantita_magazzino>0 and ordini_dett.idpro>0 and ordini_dett.idord="&c_idord
		call scorri_magazzino(sql,c_operazione)
	end sub




    Private sub scorri_magazzino(sql,operazione)
	    dim rs_magazzino, n_record
	    c_txt_log=c_txt_log&"scorri_magazzino(sql:"&sql&", operazione:"&operazione&")<br>"
		n_record=1
	    'Elenco campi rs_magazzino
	    	'idmag
	    	'quantita_magazzino
	    	'quantita_movimento
	    	'db_ven
	    	'db_acq
	    
	    	'operazione:
			'1 Rettifica giacienza
			'2 Scarico magazzino per ordine
			'3 Carico magazzino da ordine fornitore
			'4 Carico magazzino per reso DDT
			'5 Ricalcolo magazzino per Distinta Base
		on error resume next
		set rs_magazzino=conn.execute (sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
	    
	    dim quantita_movimento, quantita_magazzino, aggiorna_giacenza
	    if not rs_magazzino.eof then
		    do while not rs_magazzino.EOF
				idmag=rs_magazzino("idmag")
			    aggiorna_giacenza=false
		    	
		    	quantita_magazzino=cdbl(rs_magazzino("quantita_magazzino"))
		    	quantita_movimento=cdbl(rs_magazzino("quantita_movimento"))
				
			    c_txt_log=c_txt_log&"<hr>scorri_magazzino()>>n_record:"&n_record&", idmag:"&idmag&"(quantita_magazzino:"&quantita_magazzino&", quantita_movimento:"&quantita_movimento&")"
	
		    	'Gestisco le quantita
		    	select case operazione
		    		case 2,6	'Scarico magazzino per ordine
						quantita_scalato_magazzino=cdbl(rs_magazzino("quantita_scalato_magazzino"))
		    			c_txt_log=c_txt_log& "quantita_scalato_magazzino: "&quantita_scalato_magazzino
						quantita_movimento_originale=quantita_movimento
						quantita_movimento=quantita_movimento-quantita_scalato_magazzino
		    			if quantita_movimento>quantita_magazzino and rs_magazzino("db_ven")=0 then  'Se c'è db non rettifico la quantita movimento
			    			'Non posso scaricare più di quello che è a magazzino
			    			quantita_movimento=quantita_magazzino		    			
			    		end if
			    		quantita_movimento=-quantita_movimento
			    		quantita_magazzino=quantita_magazzino+quantita_movimento
					    aggiorna_giacenza=true
			    		
		    		case 3	'Carico magazzino da ordine fornitore
			    		quantita_magazzino=quantita_magazzino+quantita_movimento
					    aggiorna_giacenza=true
		    		case 4	'Carico magazzino per reso DDT
			    		quantita_magazzino=quantita_magazzino+quantita_movimento
					    aggiorna_giacenza=true
					    
		    		case 1		'Rettifica giacienza da form
		    			
						newQ = TRIM( Request( "quantita" & idmag ) )
						IF isNumeric( newQ ) THEN
							newq=aggiusta_decimale (newq,"asp")
							oldQ=RS_magazzino("quantita_magazzino")
							newq=roundup(newq,2)
							if isnull(oldQ) then oldQ=0
							if Cdbl(oldQ)<>Cdbl(newQ) or isnull(RS_magazzino("quantita_magazzino")) then
								'response.write "aggiorna:true"
								aggiorna_giacenza=true
								quantita_magazzino = newQ
								quantita_movimento = newQ
								cosa_aggiorna="quantita_magazzino da "&oldQ&" a "& newq
							else
								sql="UPDATE magazzino SET data_aggiornamento=now() WHERE (((magazzino.idmag)="&idmag&"));"
								on error resume next
								conn.execute sql,num	
								if err.number<>0 then
									response.write c_txt_log&sql&"<br>"
									
								end if
								on error goto 0

							end if
						END IF
						newR = TRIM( Request( "riordino" & idmag ) )
						IF isNumeric( newR ) THEN
							oldR=RS_magazzino("quantita_riordino")
							if isnull(oldR) then oldR=0
							if Cdbl(newR)<>cdbl(oldR) or isnull(RS_magazzino("quantita_riordino")) then
								call aggiorna_riordino(idmag,newR)
								cosa_aggiorna=aggiungi_virgola(cosa_aggiorna,"riordino da "&oldR&" a "&newR)
							end if
						END IF
		    		
		    		case else	
		    		
		    	end select
			    c_txt_log=c_txt_log&" dopo (quantita_magazzino:"&quantita_magazzino&", quantita_movimento:"&quantita_movimento&"), db_ven:"&rs_magazzino("db_ven")&", db_acq:"&rs_magazzino("db_acq")&"aggiorna_giacenza:"&aggiorna_giacenza&"<br>"
		    	
		    	
		    	
		    	if aggiorna_giacenza then
			    	if rs_magazzino("db_ven")=1 then
				    	
						call scorri_db_prodotti_acquisto(idmag,quantita_movimento)
					else
						'Se non ha componenti scarico
						call aggiorna_magazzino(idmag,quantita_magazzino )
						call inserisci_movimento(idmag,quantita_movimento,operazione)
					end if
					if rs_magazzino("db_acq")=1 then
						quantita_massima=null
						call scorri_db_prodotti_vendita(idmag,quantita_magazzino)
					end if
				end if
				if operazione=2 then
					conn.execute ("update ordini_dett set quantita_scalato_magazzino=quantita_scalato_magazzino+"&aggiusta_decimale(quantita_movimento_originale,"sql")&" where iddett="&rs_magazzino("iddett"))
					c_txt_log=c_txt_log&"Aggiorno quantita_scalato_magazzino<br>"
				end if
				c_txt_log=c_txt_log&"scorri_magazzino().movenext<br>"
			    rs_magazzino.movenext
				n_record=n_record+1
		    loop
		else
			'In rs_magazzino non ci sono record
		    c_txt_log=c_txt_log&"--In rs_magazzino non ci sono record, sql:<br>"&sql&"<br>"
		end if
		set rs_magazzino = Nothing
	    c_txt_log=c_txt_log&"scorri_magazzino()<<<br>"
    end sub
    
    
    Private sub scorri_db_prodotti_acquisto(idmag_ven,quantita)
	    c_txt_log=c_txt_log&"scorri_db_prodotti_acquisto(idmag_ven:"&idmag_ven&", quantita:"&quantita&")<br>"
	    
	    dim quantita_magazzino, quantita_massima, quantita_scarico, quantita_tmp
		quantita_massima=0
	    
	    'Scorro i componenti di acquisto
    	sql="select distinta_base.*, magazzino.quantita_magazzino, magazzino.db_ven, magazzino.idmag FROM magazzino inner JOIN distinta_base ON distinta_base.idmag_acq = magazzino.idmag where distinta_base.idmag_ven="&idmag_ven
		on error resume next
    	set rs_db = conn.execute ( sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
    	
		operazione=5	'Scarico magazzino per ordine (DB)
		if not rs_db.eof then
		    c_txt_log=c_txt_log&">>Trovato prodotto di acquisto (idmag_ven:"&rs_db("idmag_ven")&"), entro nella db<br>"
			do while not rs_db.EOF
				
				quantita_magazzino=roundup(rs_db("quantita_magazzino"),3)
				quantita_scarico=quantita*roundup(rs_db("quantita"),3)
			    c_txt_log=c_txt_log&">>>>Trovato prodotto di acquisto (idmag:"&rs_db("idmag")&", quantita_magazzino:"&rs_db("quantita_magazzino")&")<br>"
				c_txt_log=c_txt_log&"<br>scorri_db_prodotti_acquisto()>>Trovato prodotto di acquisto idmag:"&rs_db("idmag")&" giacenza:"&quantita_magazzino&", quantita componente vendita:"&quantita&", coefficiente utilizzo:"&rs_db("quantita")&", quantita_scarico:"&quantita_scarico
				
				quantita_magazzino=quantita_magazzino+quantita_scarico
				c_txt_log=c_txt_log&" ,nuova giacenza:"&quantita_magazzino&"<br>"
				
				call aggiorna_magazzino(rs_db("idmag"),quantita_magazzino)
					
				quantita_tmp=quantita_magazzino/rs_db("quantita")
				if quantita_massima>quantita_tmp or quantita_massima=0 then quantita_massima=quantita_tmp
				
					
					
					if rs_db("db_ven")=1 then
						call scorri_db_prodotti_acquisto(rs_db("idmag"),quantita_scarico)
					else
						call inserisci_movimento(rs_db("idmag"),quantita_scarico,operazione)
					end if
				rs_db.movenext
			Loop	
		else
			'In magazzino indicato db_ven ma senza db
			call add2log("'In magazzino indicato db_ven ma senza db idmag:"&idmag_ven,0)
		end if
		set rs_db=nothing
	
		'Aggiorno magazzino del prodotto venduto
		call aggiorna_magazzino(idmag_ven,quantita_massima)    
    End Sub
    
    
    Private sub scorri_db_prodotti_vendita(idmag_acq, quantita_magazzino)
	    c_txt_log=c_txt_log&"scorri_db_prodotti_vendita(idmag_acq:"&idmag_acq&", quantita:"&quantita_magazzino&")<br>"
		dim quantita_tmp, quantita_massima
	    quantita_massima=NULL
	    sql="select distinta_base.*, magazzino.db_acq FROM magazzino INNER JOIN distinta_base ON distinta_base.idmag_ven =magazzino.idmag  where idmag_acq="&idmag_acq	    
	    'Scorro i componenti di vendita
		on error resume next
    	set rs_db = conn.execute ( sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
	    
    	if not rs_db.eof then
			do while not rs_db.eof 
			idmag=rs_db("idmag_ven")
		    c_txt_log=c_txt_log&">>Trovato prodotto di vendita (idmag_ven:"&rs_db("idmag_ven")&"), entro nella db<br>"
			
			'Entro nella db e scansiono i prodotti di acquisto del prodotto di vendita
			sql="select magazzino.idmag, magazzino.quantita_magazzino, distinta_base.quantita FROM distinta_base INNER JOIN magazzino ON (distinta_base.idmag_acq = magazzino.idmag) where distinta_base.idmag_ven="&rs_db("idmag_ven")
			'response.write sql
			set rs_db_acq=conn.execute(sql)
			do while not rs_db_acq.EOF
			    c_txt_log=c_txt_log&">>>>Trovato prodotto di acquisto (idmag:"&rs_db_acq("idmag")&", quantita_magazzino:"&rs_db_acq("quantita_magazzino")&")<br>"
					quantita_magazzino=rs_db_acq("quantita_magazzino")
					
				    c_txt_log=c_txt_log&"scorri_db_prodotti_vendita()>>Trovato prodotto di acquisto (idmag:"&rs_db_acq("idmag")&", quantita_magazzino:"&quantita_magazzino&")<br>"
					if isnull(quantita_magazzino) then
						quantita_magazzino=0
						quantita_tmp=0
					else
						quantita_tmp=quantita_magazzino/rs_db_acq("quantita")
					end if
					c_txt_log=c_txt_log&"coefficente utilizzo:"&rs_db_acq("quantita")&" quantita_tmp:"&quantita_tmp&"<br>"
					
					c_txt_log=c_txt_log&"quantita_massima prima:"&quantita_massima&"<br>"
					if quantita_massima>quantita_tmp or isnull(quantita_massima) then quantita_massima=quantita_tmp
					c_txt_log=c_txt_log&"quantita_massima dopo:"&quantita_massima&"<br>"
			    c_txt_log=c_txt_log&">>>>Calcolo quantita massima (quantita_massima:"&quantita_massima&")<br>"
					
			rs_db_acq.MoveNext
			Loop
			set rs_db_acq=nothing
			
			'Trovato quantita massima per prodotto di vendita
		    c_txt_log=c_txt_log&">>>>Trovato quantita massima per prodotto di vendita (idmag:"&idmag&", quantita_massima:"&quantita_massima&")<br>"
			
			call aggiorna_magazzino(idmag,quantita_massima)
			
			c_txt_log=c_txt_log&"<br>Aggiornato magazzino prodotto di vendita:"&idmag&" quantità massima:"&quantita_massima
			
			
			if rs_db("db_acq")=1 then
				session("scarica_distintabase")=session("scarica_distintabase")&"<br>Richiamo ricorsivo aggiorna_magazzino_db quantita_magazzino:"&quantita_magazzino
				call scorri_db_prodotti_vendita(rs_db("idmag_ven"),quantita_magazzino)
			end if
	
			rs_db.MoveNext
			Loop
		else
			'In magazzino indicato db_acq ma senza db
			call add2log("In magazzino indicato db_acq ma senza db idmag:"&idmag_acq,0)
		end if

	    
    
    End Sub
    
    
	    
    Private sub aggiorna_magazzino(idmag,quantita)
	    dim sql, num
	    quantita=roundup(quantita,3)
	    c_txt_log=c_txt_log&"aggiorna_magazzino(idmag:"&idmag&", quantita:"&quantita&")<br>"
		sql="UPDATE magazzino SET magazzino.quantita_magazzino = "&aggiusta_decimale(quantita,"sql")&",data_aggiornamento=now() WHERE (((magazzino.idmag)="&idmag&"));"
		on error resume next
		conn.execute sql,num	
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
			
		end if
		on error goto 0
    end sub
    Private sub aggiorna_riordino(idmag,quantita)
	    dim sql, num
	    quantita=roundup(quantita,3)
	    c_txt_log=c_txt_log&"aggiorna_riordino(idmag:"&idmag&", quantita:"&quantita&")<br>"
		sql="UPDATE magazzino SET magazzino.quantita_riordino = "&aggiusta_decimale(quantita,"sql" )&" WHERE (((magazzino.idmag)="&idmag&"));"
		on error resume next
		conn.execute sql,num	
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
    end sub
    Private sub inserisci_movimento(idmag,quantita,operazione)
	    c_txt_log=c_txt_log&"inserisci_movimento(idmag:"&idmag&", quantita:"&quantita&", operazione:"&operazione&", idord:"&c_idord&", idfor:"&c_idfor&")<br>"
	    dim sql, num
	    quantita=roundup(quantita,3)
	    sql="INSERT INTO magazzino_movimenti (data,idmag,quantita,causale,iduser,idord,idfor) values (now(),"&idmag&","&aggiusta_decimale(quantita,"sql")&","&operazione&","&sessionIDUser&","&c_idord&","&c_idfor&");"
		on error resume next
		conn.execute sql,num	
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		on error goto 0
    end sub


End Class
	
	
	sub aggiorna_magazzino_db_old(idpro_acq,idvara_liv1,idvarb_liv1,quantita)
	dim idpro,idvara,idvarb
	session("scarica_distintabase")=session("scarica_distintabase")&"<br>Prodotto acquistato:"&idpro_acq&"-"&idvara_liv1&"-"&idvarb_liv1
	'idpro ,idvara ,idvarb del prodotto acquistato
	'Risalgo la db e cerco i prodotti venduti
	'set rs_db_vend=conn.execute ("select distinta_base.* from distinta_base where idpro_acq="&idpro_acq&" and idvara_acq="&idvara_liv1&" and idvarb_acq="&idvarb_liv1)
	
	
set rs_db_vend=conn.execute ("select distinta_base.*, magazzino.db_acq FROM magazzino INNER JOIN distinta_base ON (magazzino.idpro = distinta_base.idpro_ven) AND (magazzino.idvara = distinta_base.idvara_ven) AND (magazzino.idvarb = distinta_base.idvarb_ven) where idpro_acq="&idpro_acq&" and idvara_acq="&idvara_liv1&" and idvarb_acq="&idvarb_liv1)
	
	
	if not rs_db_vend.eof then
		do while not rs_db_vend.eof 
		idpro=rs_db_vend("idpro_ven")
		idvarA=rs_db_vend("idvara_ven")
		idvarb=rs_db_vend("idvarb_ven")
		session("scarica_distintabase")=session("scarica_distintabase")&"<br>In db trovato prodotto venduto:"&idpro&"-"&idvara&"-"&idvarb&" SCANSIONO LA DB<br>"
		'Entro nella db e scansiono i prodotti di acquisto del prodotto di vendita
		dim quantita_tmp
		sql="select magazzino.idpro,magazzino.idvara,magazzino.idvarb,magazzino.quantita_magazzino, distinta_base.quantita, distinta_base.idpro_ven, distinta_base.idvara_ven, distinta_base.idvarb_ven ,magazzino.db_ven, magazzino.db_acq FROM distinta_base INNER JOIN magazzino ON (distinta_base.idpro_acq = magazzino.idpro) AND (distinta_base.idvarb_acq = magazzino.idvarb) AND (distinta_base.idvara_acq = magazzino.idvara) where distinta_base.idpro_ven="&idpro&" and idvara_ven="&idvarA&" and idvarb_ven="&idvarB
		'response.write sql
		set rs_db_acq=conn.execute(sql)
		do while not rs_db_acq.EOF
			session("scarica_distintabase")=session("scarica_distintabase")&"<br>In db trovato prodotto acquistato:"&rs_db_acq("idpro")&"-"&rs_db_acq("idvara")&"-"&rs_db_acq("idvarb")&" quantita a magazzino:"&rs_db_acq("quantita_magazzino")&"<br>"
				
		
				quantita_magazzino=rs_db_acq("quantita_magazzino")
				if isnull(quantita_magazzino) then
					quantita_magazzino=0
					quantita_tmp=0
				else
					quantita_tmp=quantita_magazzino/rs_db_acq("quantita")
				end if
				session("scarica_distintabase")=session("scarica_distintabase")&"coefficente utilizzo:"&rs_db_acq("quantita")&" quantita_tmp:"&quantita_tmp&"<br>"
				
				session("scarica_distintabase")=session("scarica_distintabase")&"quantita_massima prima:"&quantita_massima&"<br>"
				if quantita_massima>quantita_tmp or isnull(quantita_massima) then quantita_massima=quantita_tmp
				session("scarica_distintabase")=session("scarica_distintabase")&"quantita_massima dopo:"&quantita_massima&"<br>"
				
				
		rs_db_acq.MoveNext
		Loop
		set rs_db_acq=nothing
		
		'Trovato quantita massima per prodotto di vendita
		session("scarica_distintabase")=session("scarica_distintabase")&"<br>Imposto quantita calcolata:"&idpro&"-"&idvara&"-"&idvarb&" quantita:"&quantita_massima&"<br>"
		quantita_massima=replace(quantita_massima,",",".")
		sql="UPDATE magazzino SET magazzino.quantita_magazzino = "&quantita_massima&",data_aggiornamento=now() WHERE idpro="&idpro&" and idvara="&idvara&" and idvarb="&idvarb
		'session("scarica_distintabase")=session("scarica_distintabase")&"Query aggiornamento magazzino:"&sql&"<br>"
		conn.execute(sql	)
		session("scarica_distintabase")=session("scarica_distintabase")&"<br>Aggiornato magazzino prodotto di vendita:"&idpro&"-"&idvara&"-"&idvara&" quantità massima:"&quantita_massima
		
		
		
		
		
		if rs_db_vend("db_acq") then
			session("scarica_distintabase")=session("scarica_distintabase")&"<br>Richiamo ricorsivo aggiorna_magazzino_db quantita_magazzino:"&quantita_magazzino
			call aggiorna_magazzino_db(rs_db_vend("idpro_ven"),rs_db_vend("idvara_ven"),rs_db_vend("idvarb_ven"),quantita_magazzino)
		end if

		
		
		
		
		
		
		
		
		
		rs_db_vend.MoveNext
		Loop
	end if
	session("scarica_distintabase")=session("scarica_distintabase")&"<br>--------"

end sub

	
	
	
	
	
	
	
	
	
	sub scarica_distintabase_old(idpro,idvarA,idvarB,quantita,idmag)
	
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
		session("scarica_distintabase")=session("scarica_distintabase")&"<br>Trovato componente acquisto idmag:"&rs_db("idmag")&"<br>"
			
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
		
		'conn.execute("UPDATE magazzino SET magazzino.quantita_magazzino = "&quantita_massima&",data_aggiornamento=now() WHERE (((magazzino.idmag)="&idmag&"));"	)
		call aggiorna_magazzino(idmag,quantita_massima)
		session("scarica_distintabase")=session("scarica_distintabase")&"Aggiorno quantità massima nel prodotto di vendita quantita:"&quantita_massima&"<br>"
	else
		session("scarica_distintabase")=session("scarica_distintabase")&"<b>Problema idmag nullo</b><br>"
		add2log "idpro:"&idpro&" idvara:"&idvara&" idvarb:"&idvarb&" errore in scarica_distintabase, idmag è nullo",3
	end if
end sub

	
	
	
	sub movimento_magazzino_old(idpro,quantita,idvara,idvarb,causale,idord,idmag)
	'CAUSALE:
	'1 Rettifica giacienza
	'2 Scarido magazzino per ordine
	'3 Carico magazzino
	'4 Carico magazzino per reso DDT
	'5 Scarico magazzino per ordine (DB)
	'6 Carico magazzino da ordine fornitore
	dim quantita_massima
	if causale=1 then '---------------------------RETTIFICA GIACIENZA
		sql="UPDATE magazzino SET magazzino.quantita_magazzino = "&quantita&",data_aggiornamento=now() WHERE (((magazzino.idmag)="&idmag&"));"	
		quantita=replace(quantita,",",".")
		conn.execute "INSERT INTO magazzino_movimenti (data,idpro,quantita,idvara,idvarb,causale,iduser,idord) values (now(),"&idpro&","&quantita&","&idvara&","&idvarb&","&causale&","&session("iduser")&","&idord&");"
		
	elseif causale=2 or causale=6 or causale=4 then '---------------------------SCARICO magazzino PER ORDINE e CARICO magazzino da ordine fornitore
		sql="select ordini_dett.iddett, ordini_dett.idpro, ordini_dett.idvara, ordini_dett.idvarb, ordini_dett.quantita, ordini_dett.reso, magazzino.quantita_magazzino, magazzino.db_ven,magazzino.db_acq,magazzino.idmag FROM magazzino RIGHT JOIN ordini_dett ON (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro) where ordini_dett.idpro>0 and ordini_dett.idord="&idord
		if causale=6 then
			sql=replace(sql,"ordini_dett","ordini_fornitori_dett")
		end if
		Set rs_magazzino = conn.execute(sql)
		do while NOT rs_magazzino.EOF
			if rs_magazzino("db_ven")=0 then
				
				'Non è un componente di vendita quindi non scansiono db
				'response.write " idpro:"&rs_magazzino("idpro")&" isnumeric"&isnumeric(rs_magazzino("quantit?"))&"<br>"
				if  isnumeric(rs_magazzino("quantita_magazzino")) then 'PEZZA per il problema degli articoli con variante non specificata
					quantita=null
					if causale=2 then 'Per scarico ordine verifico la congruenza della quantità magazzino e ordine
						if rs_magazzino("quantita_magazzino")>=rs_magazzino("quantita") then
							quantita=rs_magazzino("quantita")
							quantita_magazzino=rs_magazzino("quantita_magazzino")-rs_magazzino("quantita")
						else
							quantita=rs_magazzino("quantita_magazzino")
							quantita_magazzino=0
						end if
					elseif causale=4 then	
						if rs_magazzino("reso")=0 and lcase(trim(request.form("reso"&rs_magazzino("iddett"))))="si" then
							quantita=rs_magazzino("quantita")
							quantita_magazzino=rs_magazzino("quantita_magazzino")+rs_magazzino("quantita")
						end if
						
					elseif causale=6 then
						quantita=rs_magazzino("quantita")
						quantita_magazzino=rs_magazzino("quantita_magazzino")+rs_magazzino("quantita")
					end if
					if not isnull(quantita) then
						conn.execute ("update magazzino set quantita_magazzino="&aggiusta_decimale( quantita_magazzino,sql) &" WHERE idmag="&rs_magazzino("idmag"))
						session("scarica_distintabase")=session("scarica_distintabase")&"Modificato idmag:"&rs_magazzino("idmag")&" quantita_magazzino prima:"&rs_magazzino("quantita_magazzino")&" quantita_magazzino dopo:"&quantita_magazzino
						if rs_magazzino("db_acq") then
							quantita_massima=null
							call aggiorna_magazzino_db(rs_magazzino("idpro"),rs_magazzino("idvara"),rs_magazzino("idvarb"),rs_magazzino("quantita_magazzino"))
						end if
						if causale=2 then
							quantita=0-quantita
						end if
						quantita=replace(quantita,",",".")
						conn.execute "INSERT INTO magazzino_movimenti (data,idpro,quantita,idvara,idvarb,causale,iduser,idord) values (now(),"&rs_magazzino("idpro")&","&aggiusta_decimale( quantita,"sql") &","&rs_magazzino("idvara")&","&rs_magazzino("idvarb")&","&causale&","&session("iduser")&","&idord&");"	
					end if			
				end if
			else		'Distinta Base
				
				call scarica_distintabase(rs_magazzino("idpro"),rs_magazzino("idvara"),rs_magazzino("idvarb"),rs_magazzino("quantita"),rs_magazzino("idmag"))
				
			end if
			rs_magazzino.MoveNext
		loop
		set rs_magazzino=nothing		
	end if	
end sub





	
%>