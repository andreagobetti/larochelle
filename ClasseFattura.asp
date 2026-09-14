<%
'Sintassi




'Spostare qui operazioni da fare con pulsanti


m_log_txt=""
Class ClasseFattura
	
    Private m_idfat
	Private n_fat
    Private m_iduser
	Private txt_log
		
	Private m_tabella
	private m_modifica_articoli
	Private fattura_Dictionary

	
	dim m_recordset
	
	
    Public Default Function Init(parameters)
	    m_tabella="fatture"
	    m_log_txt="Init<br>"
	    m_log=false
	    m_chk_calcolato=false
	    dim come
	    Set Init = Me
		m_nuovo=false
		if isarray(parameters) then
		    n_parametri=UBound(parameters)+1
		    oper=lcase(parameters(0))
	        
	        if n_parametri>2 then
				m_idfat=parameters(2)
			end if
	
		    select case oper
	
		    	case "apri"
					'set objDocumento = (new ClasseFattura)(array("apri","",idord,"totali"))
					m_log_txt=m_log_txt&", CASE apri,"
		
					m_chk_calcolato=true
			        
					call ApriRecord()
					call ControlloOperazioni()
					
		    	case "crea"
					'set objdocumento= (new ClasseFattura)(array("crea","tipo_fattura:ordine,ddt,fattura","tipo:accompagnatoria","iddocumento","iduser"))
					n=crea_fattura(parameters)
					
				
			end select    
		else
			m_idfat=parameters   
			call ApriRecord()

		end if 
    End Function
    
    Private Sub Class_Terminate
	    if true then call add2log(m_log_txt,0)
		Set fattura_Dictionary = Nothing
	End Sub


    
	Private sub ApriRecord()
		dim rs_ordine
		dim n
		dim flds
		'response.write  "ApriRecord"
		Set fattura_Dictionary=Server.CreateObject("Scripting.Dictionary")
		sql="select * from fatture where idfat="&m_idfat
		
		'on error resume next
		set rs_ordine=conn.execute (sql)
		if err.number<>0 then
			m_log_txt=m_log_txt&"Errore in ApriRecord:"&sql
			response.write m_log_txt&"<br>"
			response.end
		end if
		on error goto 0
		'rs_ordine.Open sql, conn, 1, 3

		if not rs_ordine.eof then
			m_log_txt=m_log_txt&" record trovato"
			m_esiste=true
			
			'memorizzo i nomi dei campi e i valori attuali

			m_n_campi=rs_ordine.fields.count-1
			'response.write "m_n_campi:"&m_n_campi&"<br>"
			Set Flds = rs_ordine.Fields
			
			for i=0 to m_n_campi
				value=rs_ordine(i).value
				if VarType(value)=14 then value=cdbl(value)
				fattura_Dictionary.add lcase(rs_ordine.Fields(i).Name),value
			next
			
			if DateDiff("d","01/01/2015",fattura_Dictionary.item("data")) <0 then 'prima del 01/01/2015
				m_vecchio_conteggio=true
			else
				m_vecchio_conteggio=false
			end if
			
		    
		else
			m_log_txt=m_log_txt&" record non trovato"
			m_esiste=false
		end if
		set rs_ordine=Nothing
		m_log_txt=m_log_txt&" Fine ApriRecord m_iduser:"&m_iduser
    End sub

	
	
    
    sub ControlloOperazioni()
		IF Request.form("aggiorna_ordine") <> "" THEN
			n=calcolatotalemerce(true)
		end if

	    if request.form("ricalcola")<>"" then
			n=calcolatotalemerce(false)
		end if
		
		if request("elimina_riga_ordine")<>"" then
			call elimina_riga_ordine()
		end if
		
		
		'-------------------------------------------
		'Penso che non servono più
		if request("aggiungisopra")<>"" then
			call AggiungiRiga(-1)
		end if
		if request("aggiungisotto")<>"" then
			call AggiungiRiga(1)
		end if
		'------------------------------------------
		
	end sub
    
	private function crea_fattura(array_param)
		'tipo_fattura= "fattura","ddt","ordine"
		'tipo="accompagnatoria"
		'set objdocumento= (new ClasseFattura)(array("crea","tipo_fattura:ordine,ddt,fattura","tipo:accompagnatoria","iddocumento","iduser"))
		tipo_fattura=array_param(1)
		tipo=array_param(2)
		idord=array_param(3)
		iduser=array_param(4)
		
		dim fattura_pa,rs_tmp
		call add2log(queryeform(),0)
		Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class
		idord=request.form("idord")
		
		
		if tipo_fattura="dtt" or tipo_fattura="ordine" then
			'verifico se esiste già una fattura
			sql="select ordini_fatture.idord FROM ordini_fatture WHERE (((ordini_fatture.idord)="&idord&"));"
			set rs_tmp=conn.execute(SQL)
			rs_eof=rs_tmp.eof
			set rs_tmp=nothing
			if not rs_eof then 
				'esiste già
				crea_fattura="-1"
				exit function
			end if
		end if
		if tipo_fattura="" and idord<>"" then
			'Non ho tipo_fattura ma ho idord allora determino tipo_fattura
			sql="select ordini.* FROM ordini where ordini.idord="&idord
			Set rs_ordini = conn.execute(sql)
			nord=rs_ordini("nord")
			if rs_ordini("tipo_documento")="ddt" then
				call add2log("Crea fattura per ddt idddt:"&idord&" nddt"&nord,0)
				fatturaper="[ddt="&idord&"]"&nord&"[/ddt]"
				tipo_fattura="ddt"
			else
				fatturaper="[ordine="&idord&"]"&nord&"[/ordine]"
				tipo_fattura="ordine"
			end if
			trattamento_iva_ordine=rs_ordini("trattamento_iva_ordine")
			iduser=rs_ordini("iduser")
			idintestazione=rs_ordini("idintestazione")
			idconsegna=rs_ordini("idconsegna")
			totale_ordine=rs_ordini("totale")
			set rs_ordini = nothing
		end if	
		if tipo_fattura="fattura" then
			idord=0
		end if
		
		fattura_pa=request.form("fattura_pa")
		if request.form("fattura_sp")="si" then
			fattura_sp=1
		else
			fattura_sp=0
		end if
		
		'Determino fattura PA
		if fattura_pa="" then
			if trattamento_iva_ordine=10 and  DateDiff("d","31/03/2015",date())>0 then 'dopo il 31/03/2015 then
				fattura_pa=1
			else
				fattura_pa=0
			end if
			tmp_trattamento_iva=trattamento_iva_ordine
		elseif fattura_pa=1 then
			tmp_trattamento_iva=10
		elseif fattura_sp=1 then
			tmp_trattamento_iva=10

		end if
		call add2log("form fattura_pa:"&request.form("fattura_pa")&queryeform(),0)
			
		'Determino il numero fattura
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
		m_Nfat=max_fatt(fattura_pa,anno)
			
		sql="select * from utenti where iduser="&iduser
		Set rs_utente = Server.CreateObject("ADODB.Recordset")
		rs_utente.Open sql, conn, 3, 3
		if request.form("aggiorna_dati_utente")<>"" then
			rs_utente("pag_accordato")=request.form("pagamento")
			rs_utente("iban")=ucase(request.form("iban"))
			rs_utente("banca_appoggio")=ucase(request.form("banca_appoggio"))
			rs_utente.update
		end if
		spese_0=rs_utente("spese_0")
		idbanca=rs_utente("idbanca")
		rs_utente.close
		set rs_utente = nothing
		
		'cerco banca predefinita
		if idbanca="" or idbanca=0 then idbanca=Application("banca")
		
		Set rs_fattura = Server.CreateObject("ADODB.Recordset")
			
		'Aggiungo fattura
		sql="select fatture.* FROM fatture"
		rs_fattura.Open sql, conn, 1, 3
		rs_fattura.addnew
		rs_fattura("nfat")=m_Nfat
		rs_fattura("data")=data
		rs_fattura("tipo_fattura")=tipo_fattura
		rs_fattura("anno")=year(data)
		rs_fattura("iduser")=iduser
		rs_fattura("stato_fattura")=1

		if idintestazione="" then
			nome=AllFirstUp(trim(request.form("nome")))
			cognome=AllFirstUp(trim(request.form("cognome")))
			azienda=AllFirstUp(trim(request.form("azienda")))
		
		
			idintestazione= aggiungi_intestazione(iduser,azienda,cognome,nome,ucase(trim(request.form("cf"))),trim(request.form("piva")),trim(request.form("Indirizzo")),allfirstup(trim(request.form("citta"))),trim(request.form("cap")),trim(request.form("provincia")))
		end if

		rs_fattura("idintestazione")=idintestazione

			
			
		'Creo dati_consegna
		if idconsegna="" then
			idconsegna=0
			if request.form("d_indirizzo")<>"" then
				idconsegna=aggiungi_consegna(iduser,request.form("d_azienda"),request.form("d_indirizzo"),request.form("d_citta"),request.form("d_cap"))
			end if
			if request.form("consegna")<>"" then
				idconsegna=request.form("consegna")
			end if
		end if
		rs_fattura("idconsegna")=idconsegna

			
		rs_fattura("idord")=idord
		
		rs_fattura("annotazioni")=request.form("annotazioni")
		
		if request.form("data_inizio")<>"" then
			rs_fattura("data_inizio")=request.form("data_inizio")
		end if
		rs_fattura("incaricato")=request.form("incaricato")
		rs_fattura("pagamento")=request.form("pagamento")
		if request.form("CodiceDestinatario")<>"" then
			rs_fattura("CodiceDestinatario")=request.form("CodiceDestinatario")
		end if
		
		'Spese bancarie
		if spese_0 or isnull(rs_fattura("pagamento")) or cdbl(totale_ordine)<0 then
			spese_bancarie_v=0
		elseif request.form("spese_bancarie")<>"" then
			spese_bancarie_v=aggiusta_decimale(request.form("spese_bancarie"),"asp")
		else
			spese_bancarie_v=metodipagamento.spese_bancarie(rs_fattura("pagamento"))
		end if
		rs_fattura("spese_bancarie")=spese_bancarie_v
			
		rs_fattura("banca_appoggio")=request.form("banca_appoggio")
		rs_fattura("iban")=ucase(request.form("iban"))
		rs_fattura("creata_da")=session("nominativo")
		rs_fattura("eliminato")=0
		rs_fattura("fine_mese")=0
		rs_fattura("aliquota_iva")=iva(date())
		if tmp_trattamento_iva="" then
			tmp_trattamento_iva=request.form("trattamento_iva")
		end if
		rs_fattura("trattamento_iva")=tmp_trattamento_iva
		rs_fattura("idbanca")=idbanca
		if checkbox("accompagnatoria")=1 then 
			set rs_ddt=conn.execute("select ddt.causale from ddt where idord="&idord)
			if not rs_ddt.eof then
				rs_fattura("causale")=rs_ddt("causale")
			else
				rs_fattura("causale")=request.form("causale")
			end if
			set rs_ddt= Nothing
			rs_fattura("porto")=request.form("porto")
			rs_fattura("imballo")=request.form("imballo")
			rs_fattura("colli")=request.form("colli")
			rs_fattura("peso")=request.form("peso")
			rs_fattura("dimensione")=request.form("dimensione")
			rs_fattura("vettore")=request.form("vettore")
		else
			rs_fattura("causale")=1
			rs_fattura("fine_mese")=1
		end if
			
		rs_fattura("pa")=fattura_pa
		pagamento=rs_fattura("pagamento")
		rs_fattura.update
		m_idfat=get_last_id("fatture")
		rs_fattura.Close
		set rs_fattura = nothing
		if tipo_fattura="ordine" or tipo_fattura="ddt" then
		
			sql="insert into ordini_fatture set (idord,idfat) values ("&idord&","&m_idfat&")"
			conn.execute (sql)	
			if spese_bancarie_v>0 then
				conn.execute("update ordini set spese_bancarie_ordine="&aggiusta_decimale(spese_bancarie_v,"sql")&" where idord="&idord)
			end if
			totale_fattura=calcola_totale_fattura(m_idfat)
		else
			totale_fattura=0
		end if
		'creo scadenze
		
		testo="Creata [fattura="&m_idfat&"]"&pre_fattura(data)&m_Nfat&"/"&year(data)&"[/fattura] per "&fatturaper
		
		testo=testo&" di importo "&formatcurrency(totale_fattura,2)
		n_incassi=collega_incassi (0,m_idfat)
		if totale_fattura>0 then 
			n=metodipagamento.genera_scadenze(m_idfat,m_Nfat,idord,data,totale_fattura,pagamento)
		else
			n=0
		end if
		testo=testo&", eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&n_incassi&" incassi precedenti"&vbcrlf&n&" scadenze incasso"&vbcrlf&"Idord: "&IDORD
		add2log testo,2
		crea_fattura=1
	end function
	
	
	Private sub elimina_riga_ordine()
			txt_log=""
			iddett=request.querystring("iddett")
			
			sql="select * from "&m_tabella&"_dett where iddett="&iddett
			set rs=conn.Execute(sql)
			articolo_ordine=rs("articolo_ordine")
			quantita=rs("quantita")
			subtotale=cdbl(rs("totale_riga"))
			if m_tabella="ordini" then
				v_numeri_di_serie=rs("numeri_di_serie")
			else
				v_numeri_di_serie=0
			end if
			Set rs = Nothing
			
			call EliminaRiga(iddett,v_numeri_di_serie)
			fattura_Dictionary.Item("totale_merce")=fattura_Dictionary.Item("totale_merce")-subtotale
			
			n=CalcolaTotale()
			'call concatena_stringa(txt_log,"<br>","Eliminato "&iddett&" riga <b>"&articolo_ordine&"</b> quantita: "&quantita)
			'sql="UPDATE "&m_tabella&" SET comunicazioni_precedenti= CONCAT(comunicazioni_precedenti,'<hr>"& now()&" <b>"&session("nominativo")&":</b><br>"&replace(testo_log,"'","''")&"') WHERE idord = "&m_idfat&";"
			'conn.execute (sql)
			
			add2log "Modificato "&m_cosa_log(m_tabella,m_idfat,m_nord)&vbcrlf&"Eliminato riga "&iddett&" <b>"&articolo_ordine&"</b> quantita: "&quantita&"totale_merce: "&fattura_Dictionary.item("totale_merce")&vbcrlf&"Versione finale:"&fattura_Dictionary.item("versione"),2
	end sub
	
	
	Private sub EliminaRiga(iddett,v_numeri_di_serie)
		call da_calcolare()
		if cint(v_numeri_di_serie)>0 then
			set rs=conn.execute("select seriale from numeri_seriali where iddett="&iddett)
			if not rs.eof then
				call concatena_stringa(txt_log,"<br>","Liberato numero di serie: "&rs("seriale"))
				conn.execute ("UPDATE numeri_seriali set idord= NULL, iddett=NULL where iddett="&iddett)
			end if
			set rs = Nothing
		end if
		'conn.execute ("delete from "&m_tabella&"_dett_note where iddett="&iddett)
		'conn.execute ("delete from "&m_tabella&"_dett_spettanze where iddett="&iddett)
		conn.execute ("delete from "&m_tabella&"_dett where iddett="&iddett)
		
	end sub
	
	Private sub AggiungiRiga(dove)
		iddett=request.querystring("iddett")
		sql="select * from "&m_tabella&"_dett where iddett="&iddett
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		posizione=rs("ordine")
		rs.addnew
		rs("idord")=idord
		rs("ordine")=posizione+dove
		rs("quantita")=-1
		'rs("codice_ordine")=null
		rs("articolo_ordine")=""
		rs("modificato")=true
		rs.update
		response.write rs("iddett")
		rs.close
		Set rs = Nothing
		call add2log("Aggiungiriga in classe ordine utilizzata",-1)
	end sub

    
	Public sub elenco_campi
		dim txt,txt_tmp
		if m_esiste then
			
			txt= "<b>Elenco campi</b><br>"
			For Each Key in fattura_Dictionary
			 txt_tmp=Key & "=" & fattura_Dictionary.item(Key) & "<br>"
			 txt=txt&txt_tmp
			Next
			call add2log(txt,0)
		else
			response.write "<b>Record idord="&m_idfat&" nella tabella "&m_tabella&" non trovato</b><br>"
		end if
	end sub
	
	
	Public function CalcolaTotaleMerce(da_form)
	    m_log_txt=m_log_txt&"<br>CalcolaTotaleMerce"
		'on error goto 0
		dim subtotale
		dim imposta
		dim aliquota_iva
		dim versione
		dim ordine_modificato
		dim txt_modifiche
		dim gravita_log
		ordine_modificato=false
		txt_modifiche=""
		gravita_log=2
		versione=fattura_Dictionary.item("versione")
		subtotale=0
		if false then
		'if da_form then
			if cint(request.form("versione"))<>versione then
				call add2log ("<b>Versione "&m_cosa_log(m_tabella,m_idfat,m_nord)&" non corrispondente (versione recordset: "&versione&" versione form: "&request.form("versione")&"<br>Aggiornamento "& cosa&" non eseguito." ,3)
				session("report")="Mentre era aperto in questa pagina del browser questo "&cosa&" &egrave; stato modificato in un altra pagina del browser o da un altro utente. <br>L'aggiornamento dei dati non &egrave; stato eseguito ed &egrave; stata ricaricata la versione pi&ugrave; recente."
				if utente_andrea then 
					session("report")=session("report")&"<br><b>Versione "&m_cosa_log(m_tabella,m_idfat,m_nord)&" non corrispondente (versione recordset: "&versione&" versione form: "&request.form("versione")&"<br>"
				end if
				exit function				
			end if
		end if

		sql = "select * FROM "&m_tabella&"_dett WHERE idfat=" & m_idfat
		Set rs_dettaglio = Server.CreateObject("ADODB.Recordset")
		rs_dettaglio.Open sql, conn, 3, 3
		do while not rs_dettaglio.eof
			if da_form then
				session("txt_aggiorna_ordine")=session("txt_aggiorna_ordine")&"<br>Analizzo iddett:"&rs_dettaglio( "iddett" )
				if Request( "iddett" & rs_dettaglio( "iddett" ) )<>"" then
					txt_modificato=""
					txt_riga=rs_dettaglio("articolo_ordine")
					row_deleted=false
					newQ = TRIM( Request( "quantita" & rs_dettaglio( "iddett" ) ) )
					
					IF newQ = "0"  THEN
						txt_modificato=txt_modificato&"Eliminato "&rs_dettaglio( "iddett" )&" riga <b>"&rs_dettaglio("articolo_ordine")&"</b> quantita: "&rs_dettaglio("quantita")
						row_deleted=true
						
						
						'Devo eliminare la riga ma non qui
						'rs_dettaglio.Delete
						ordine_modificato=true
						'conn.execute ("delete from "&m_tabella&"_dett where iddett="&rs_dettaglio( "iddett" ))
						if m_tabella="ordini" then
							v_numeri_di_serie=rs_dettaglio("numeri_di_serie")
						else
							v_numeri_di_serie=0
						end if
												
						call EliminaRiga(rs_dettaglio( "iddett" ),v_numeri_di_serie)
					ELSE
						
						newP = aggiusta_decimale(TRIM( Request( "prezzo" & rs_dettaglio( "iddett" ) ) ),"asp")
						newA = TRIM( Request( "articolo" & rs_dettaglio( "iddett" ) ) )
						newsconto=aggiusta_decimale (TRIM( Request( "sconto_prodotto" & rs_dettaglio( "iddett" ) ) ),"asp")
						
						conteggio_riga=conteggio_riga+1
						newA=Replace(newA,"""","&quot;")
						IF  len(newA)>0 THEN
							if rs_dettaglio("articolo_ordine")<>newA then
								txt_modificato=txt_modificato&" in "&newA
								rs_dettaglio("articolo_ordine")=newA
								modificato=true
							end if
						END IF
						IF isNumeric( newQ ) THEN
							newQ = aggiusta_decimale(newQ,"asp")
							if Cdbl(rs_dettaglio("quantita"))<>Cdbl(newQ) then
								txt_modificato=txt_modificato&", quantita da "&rs_dettaglio("quantita")&" a "&newQ
								rs_dettaglio("quantita") = newQ
								modificato=true
							end if
						END IF
						IF isNumeric( newsconto ) THEN
							newsconto=Cdbl(newsconto)
							if cdbl(rs_dettaglio("sconto_prodotto"))<>newsconto then
								txt_modificato=txt_modificato&", sconto da "&rs_dettaglio("sconto_prodotto")&" a "&newsconto
								rs_dettaglio("sconto_prodotto")=newsconto
								modificato=true
							end if
						END IF
						IF isNumeric( newp ) THEN
							newP=Cdbl(newP)
							if cdbl(rs_dettaglio("prezzo"))<>newP or isnull(rs_dettaglio("prezzo")) then
								txt_modificato=txt_modificato&", prezzo da "&rs_dettaglio("prezzo")&" a "&newP
								rs_dettaglio("prezzo")=newP
								modificato=true
							end if
						END IF
						IF  len(newA)>0 THEN
							if rs_dettaglio("articolo_ordine")<>newA then
								rs_dettaglio("articolo_ordine")=newA
								modificato=true
							end if
						END IF
						if tabella="ordini" then
							if Request( "selcs" & rs_dettaglio( "iddett" ) ) ="1" then
								rs_dettaglio("flag1_ordine")=true
								'txt_modificato=txt_modificato&", ordinato"
								'modificato=true
							else
								rs_dettaglio("flag1_ordine")=false
							end if
						end if
	
						if modificato then
							rs_dettaglio("modificato")=true
							ordine_modificato=true
							'rs_dettaglio.update
						end if
						'totale_riga_tmp=totale_riga(rs_dettaglio("prezzo"),rs_dettaglio("sconto_prodotto"),rs_dettaglio("quantita"))
						'totale_merce=totale_merce+totale_riga_tmp
						'rs_dettaglio("totale_riga")=totale_riga_tmp
					end if
	
				else
					txt_modificato=txt_modificato&"<br><b>Non trovata corrispondenza nel form per iddett:"& rs_dettaglio( "iddett" ) &"</b>"
					queryeform_txt="si"	
					gravita_log=3
				END IF
				if txt_modificato<>"" then
					if row_deleted=true then
						txt_modifiche=txt_modifiche&"<br>"&txt_modificato
					else
						txt_modifiche=txt_modifiche&"<br>"&"Modificato riga "&rs_dettaglio( "iddett" )&" <b>"&txt_riga&"</b>"&txt_modificato
					end if
				end if
				if err.number<>0 then
					call add2log("Errore aggiorna_ordine("&idord&", "&tabella&", "&da_form&") pos1 aggiornamento iddett:"&rs_dettaglio( "iddett" )&"<br>"&err.description,0)
					response.end
				end if		
				Session.Contents.remove("txt_aggiorna_ordine")
			
			end if 'da_form
			
			'Calcolo totale riga
			if not row_deleted then
	            if m_vecchio_conteggio then
		            prezzo=cdbl(rs_dettaglio("prezzo"))
		            sconto_prodotto=cdbl(rs_dettaglio("sconto_prodotto"))
		            m_totale_riga=round((prezzo-(sconto_prodotto/100)*prezzo)*rs_dettaglio("quantita"),2)
		        else
					m_totale_riga=totale_riga(rs_dettaglio("prezzo"),rs_dettaglio("sconto_prodotto"),rs_dettaglio("quantita"))
				end if
				if m_tabella<>"ordini_fornitori" then totcosto=totcosto+cdbl(rs_dettaglio("costo_ordine"))*cdbl(rs_dettaglio("quantita"))

				rs_dettaglio("totale_riga")=m_totale_riga
				rs_dettaglio.update
				subtotale=subtotale+m_totale_riga
			end if			
			rs_dettaglio.movenext
		loop
		sql_update=""
		if da_form then
			newval=aggiusta_decimale(TRIM( Request( "sconto_ordine"  ) ),"asp")
			IF isNumeric( newval ) THEN
				newval=roundup(Cdbl(newval),2)
				if fattura_Dictionary.item("sconto_ordine")<>newval then
					sql_update=sql_update&", sconto_ordine="&aggiusta_decimale(newval,"sql")
					txt_modifiche=txt_modifiche&"<br>"&"Modificato sconto ordine da "&fattura_Dictionary.item("sconto_ordine")&" a "&newval

					fattura_Dictionary.item("sconto_ordine")=newval
					ordine_modificato=true
				END IF
			end if
			newval=aggiusta_decimale(TRIM( Request( "trasporto"  ) ),"asp")
			IF isNumeric( newval ) THEN
				newval=roundup(Cdbl(newval),2)
				if fattura_Dictionary.item("trasporto")<>newval then
					sql_update=sql_update&", trasporto="&aggiusta_decimale(newval,"sql")
					txt_modifiche=txt_modifiche&"<br>"&"Modificato trasporto da "&fattura_Dictionary.item("trasporto")&" a "&newval
					fattura_Dictionary.item("trasporto")=newval
					response.write "trsporto modificato"
					ordine_modificato=true
				END IF
			end if
		end if
		'patch per evitare che totale_merce sia vuoto
		fattura_Dictionary.Item("totale_merce")=fattura_Dictionary.Item("totale_merce")+0


		
		subtotale=roundup(subtotale,2) '''''''''''''Verificare da dove arrivano i decimali
		if fattura_Dictionary.Item("totale_merce")=subtotale then
			m_totale_merce_ok=true
		else
			ordine_modificato=true
			fattura_Dictionary.Item("totale_merce") = subtotale
			m_totale_merce_ok=false
		end if 
		
		
		if ordine_modificato then
			sql="UPDATE "&m_tabella&" SET totale_merce = "&aggiusta_decimale(subtotale,"sql")&sql_update
			sql=sql&" WHERE idfat = "&m_idfat&";"
			fattura_Dictionary.item("versione")=fattura_Dictionary.item("versione")+1
			'response.write sql
			conn.execute(sql)
			'add2log "aggiorno da CalcolaTotaleMerce m_totale_merce_ok:"&m_totale_merce_ok&" txt_modifiche:"&txt_modifiche,0
			
		end if
		
		if ordine_modificato then
			add2log "Modificato "&m_cosa_log(m_tabella,m_idfat,m_nord)&vbcrlf&txt_modifiche&vbcrlf&"totale_merce:"&subtotale&vbcrlf&"Versione finale:"&versione&vbcrlf&"function CalcolaTotaleMerce",gravita_log			
		end if

		CalcolaTotaleMerce=subtotale
		n=CalcolaTotale()
			
	end function

	Public Property Get totale_merce_ok()
		totale_merce_ok=m_totale_merce_ok
	end property
	
	
	Public function CalcolaTotale()
		'on error goto 0
		dim log_txt
		log_txt=""
	    m_log_txt=m_log_txt&"<br>CalcolaTotale"
	    m_log_txt=m_log_txt&" totale_merce prima:"&fattura_Dictionary.item("totale_merce")
		dim subtotale
		dim imposta
		dim aliquota_iva

		subtotale=fattura_Dictionary.item("totale_merce")+fattura_Dictionary.item("spese_bancarie")
		log_txt="totale_merce:"&fattura_Dictionary.item("totale_merce")&" trasporto:"&fattura_Dictionary.item("trasporto")&" spese_bancarie:"&fattura_Dictionary.item("spese_bancarie")&"<br>"
		
		if esenzione_iva(fattura_Dictionary.item("trattamento_iva")) then
			m_casoiva=1
			imposta=0
			if fattura_Dictionary.item("trattamento_iva")=10 then
				m_casoiva=2
				imposta=roundup(subtotale*fattura_Dictionary.item("aliquota_iva")/100,2)
			end if
		else
			m_casoiva=3
			imposta=roundup(subtotale*fattura_Dictionary.item("aliquota_iva")/100,2)
			subtotale=subtotale+imposta
		end if
		log_txt=log_txt&"trattamento_iva:"&fattura_Dictionary.item("trattamento_iva")&" imposta:"&imposta&" casoiva:"&m_casoiva&"<br>"
		subtotale=roundup(subtotale,2) '''''''''''''Verificare da dove arrivano i decimali
		
		
		if fattura_Dictionary.item("totale_fattura")=subtotale then
			m_totale_ok=true
		else
			
			m_totale_ok=false
		end if

		sql="UPDATE "&m_tabella&" SET imposta = "&aggiusta_decimale( imposta,"sql")&", totale_fattura = "&aggiusta_decimale(subtotale,"sql")&",totale_merce="&aggiusta_decimale(fattura_Dictionary.item("totale_merce"),"sql")&"  WHERE idfat = "&m_idfat&";"
		'response.write sql
		
		conn.execute(sql)
		fattura_Dictionary.item("versione")=fattura_Dictionary.item("versione")+1
		fattura_Dictionary.item("totale_fattura")=subtotale
		fattura_Dictionary.item("imposta")=imposta
	    m_log_txt=m_log_txt&",totale aggiornato:"&fattura_Dictionary.item("totale_fattura")
		CalcolaTotale=subtotale
		log_txt=log_txt&"subtotale:"&subtotale
		
	    'call add2log(m_log_txt,0)
	    
		call add2log("CalcolaTotale "&m_tabella&"="&m_idfat&" totale:"&subtotale&vbcrlf&log_txt,0)

	end function
'******************************************************************************************************************************************************************
'******************************************************************************************************************************************************************





	Public sub elenco_articoli()
		'on error goto 0
		m_manca_numeri_di_serie=false
		m_numeri_di_serie=false
		
		'spettanze=0
		
	    m_log_txt=m_log_txt&",elenco_articoli"
	    'if m_tabella="ordini" then
		'    if fattura_Dictionary.item("stato")>3 then
		'		mostra_giacenze=false
		'	end if
		'elseif m_tabella="preventivi" then
		'	m_modifica_articoli=true
		'elseif m_tabella="ordini_fornitori" then
			'spettanze=fattura_Dictionary.item("spettanze")
		'end if
		if fattura_Dictionary.item("stato_fattura")>3 then
			m_modifica_articoli=false
		else
			m_modifica_articoli=true
		end if
		%>
		
		
		
		<tbody id="tab_articoli_body">
		<%'CICLO SU ARTICOLI
        sql="select fatture_dett.*, varianti_a.codicevara ,varianti_a.variante_a, varianti_b.variante_b FRom fatture_dett LEFT JoIN varianti_b oN fatture_dett.idvarb = varianti_b.IDvarb LEFT JoIN varianti_a oN fatture_dett.idvara = varianti_a.IDvara  left join prodotti on  fatture_dett.idpro = prodotti.idpro where idfat="&m_idfat &" order by ordine,iddett"
        
        
        
        
        
        
		set rs=conn.execute(sql)
		
		tot_costo=0
		tot_no_costo=0
		tot_vendita=0
		
		do while not rs.eof
			selcs=""
			txt_ordinato=""
			txt_fornitori=""
			txt_ordinato2=""
			tr_ordini=false
			tr_fornitori=""
			rowspan=""
			quantita_ordinato=0
			stile_cella=""
			txt_spettanze=""
			if rs("codice_ordine")<>"" then
				riga_libera=false
			else
				riga_libera=true
			end if
			
			if m_tabella<>"ordini_fornitori" then
				costo_ordine=converti_typevar14(rs("costo_ordine"))
				if rs("codice_ordine")<>"" and costo_ordine>0 then
					tot_costo=tot_costo+costo_ordine*cdbl(rs("quantita"))
					tot_vendita=tot_vendita+cdbl(rs("totale_riga"))
				else
					anomalia_costo=anomalia_costo&"<br><span class='red'><b>"&rs("codice_ordine")&"</b> costo zero</span>"

				end if
			end if
		
		
		
			if m_tabella="ordini" then
				quantita_ordinato=clng(rs("quantita_ordinato"))
				if rs("flag1_ordine") then selcs="1"
				if rs("flag1_ordine") then txt_ordinato2="Ordinato"
				
				if clng(rs("quantita_ordinato"))>0 then
					set rs_ordinato=conn.execute("select ordini_fornitori.idord, ordini_fornitori.nord, ordini_fornitori.stato, fatture_dett_ordini_fornitori_dett.quantita from fatture_dett_ordini_fornitori_dett inner join ordini_fornitori_dett on fatture_dett_ordini_fornitori_dett.iddett_fornitori = ordini_fornitori_dett.iddett inner join ordini_fornitori on ordini_fornitori_dett.idord = ordini_fornitori.idord where iddett_ordini="&rs("iddett"))
					do while not rs_ordinato.eof
						if txt_ordinato<>"" then txt_ordinato=txt_ordinato&", "
						txt_ordinato=txt_ordinato& "Ordine fornitore "&rs_ordinato("nord")
						rs_ordinato.movenext
					loop
					txt_ordinato=rs("quantita_ordinato")&" in ordine: "&txt_ordinato
					tr_ordini=true
				end if
			end if

		
			if m_tabella="ordini_fornitori" then
			
				sql="select ordini.idord, ordini.nord, ordini.stato, fatture_dett_ordini_fornitori_dett.quantita from fatture_dett_ordini_fornitori_dett inner join fatture_dett on fatture_dett_ordini_fornitori_dett.iddett_ordini = fatture_dett.iddett inner join ordini on fatture_dett.idord = ordini.idord where iddett_fornitori="&rs("iddett")
				set rs_ordinato=conn.execute(sql)
				do while not rs_ordinato.eof
					if tr_fornitori<>"" then tr_fornitori=tr_fornitori&", "
					tr_fornitori=tr_fornitori&rs_ordinato("quantita")& " per Ordine "&rs_ordinato("nord")
					rs_ordinato.movenext
				loop
			end if
		
		
		
		
			if txt_ordinato<>"" then
				rowspan="2"
				tr_ordini=true
			end if 
			if tr_fornitori<>"" then
				rowspan="2"
			end if 
			'Anticipo spettanze per determinare stile cella
			'Elenco spettanze
			if mod_larochelle  and cint(rs("quantita"))>0 and m_spettanze=1 then
				'Cerco nelle spettanze
				txt_spettanze = dettaglio_spettanze( m_tabella, RS("iddett"), rs("quantita"),false,stile_cella)
			end if

			%>
	          <tr valign="middle" id="id_<%=rs("iddett")%>" style="border-top: 1px solid gray; padding-bottom:10px; padding-top:10px;" class="<%=stile_cella%>">
			      <td align="center" rowspan="<%=rowspan%>">
	              <%if m_modifica_articoli then
	              %>
	              
	              <input type="hidden" name="selcs<%=rs("iddett")%>" id="selcs_<%=rs("iddett")%>" class="selezione" value="<%=selcs%>">
	              <span class="div_icona"><span class="ui-icon ui-icon-triangle-1-s ui-corner-all left_menu bg-gray" id="<%=rs("iddett")%>" data-ordine="<%=rs("ordine")%>"></span></span>
	              <%end if %>
	              <%if fattura_Dictionary.item("stato")<=3 and m_sposta_articoli then %>
	              <span class="div_icona"><span class="ui-icon ui-icon-arrowthick-2-n-s ui-corner-all handle bg-gray" ></span></span>
	              <%end if %>
	              <%if m_modifica_articoli then %>
	              <span class="div_icona" id="sel_<%=rs("iddett")%>"><%=txt_ordinato2%></span>
	              
	              <%end if%>
	              <%
		              if m_tabella="ordini" then
		           if cint(rs("numeri_di_serie"))>0 then %>
	              <span class="div_icona"><span class="ui-icon ui-icon-grip-dotted-horizontal ui-corner-all handle bg-gray" ></span></span>
	              <%
		              end if
		              end if %>
              <%
			txt="product.asp?idpro=" & RS("idpro")
		%>
              <b><a href="<%=txt%>" title="Clicca per le opzioni" class="menu_idpro"  data-idpro="<%=RS("idpro")%>"><%=codice_articolo_e_variante(  rs("codice_ordine"),rs("codicevara"))%></a></b>
            </td>
            <td >
            <%if rs("codice_ordine")<>"" or  m_modifica_articoli=false then%>
				<%=rs("articolo_ordine")%>
			<%else %>
				<textarea type="text" name="articolo<%=rs("iddett")%>" rows="1" class="chkchge autogrow"><%=rs("articolo_ordine")%></textarea>
			<%end if%>
            <%
	            	if rs("varianti_ordine")<>"" then response.write "<br>"&rs("varianti_ordine")
					if m_tabella="ordini" then
						if cint(rs("numeri_di_serie"))>0 then
						
							'Cerco nei seriali
							txt_seriale=""
							conteggio_seriali=0
							seriali_occorrenti=clng(rs("quantita")*rs("numeri_di_serie"))
							set seriali=conn.execute ("select numeri_seriali.* from numeri_seriali where iddett="&RS("iddett"))
							do while not seriali.eof 
								if txt_seriale<>"" then txt_seriale=txt_seriale&", "
								txt_seriale=txt_seriale&seriali("seriale")
								conteggio_seriali=conteggio_seriali+1				
								seriali.movenext
							loop
							if txt_seriale<>"" then 
								response.write "<br>Numeri di serie: "&txt_seriale
								m_numeri_di_serie=true
							end if
							if conteggio_seriali < seriali_occorrenti then
								response.write "<br><span class=""rosso"">"&seriali_occorrenti-conteggio_seriali &" numeri di serie mancanti</span>"
								m_manca_numeri_di_serie=true						
							end if
						
						end if
					end if

					response.write txt_spettanze
					'if m_tabella="ordini" then
						'if rs("nota")<>"" then
						'	response.write "<br>Nota:<span id=""notad"&RS("iddett")&""" class=""notaarticolo"">"&rs("nota")&"</span>"
						'end if
					'end if

					%></td>
            <td align="right" ><%
                if not m_modifica_articoli then %>
              <%=formatcurrency(rs("prezzo"),2)%>
              <%else %>
              <input type="text" name="prezzo<%=rs("iddett")%>" size="8" value="<%=FormatNumber(rs("prezzo"),2,,,0)%>" style="text-align:right;" class="chkchge">
              <%end if%></td>
            <%
				stile=""
				if m_mostra_giacenze and  rs("codice_ordine")<>"" then
					if m_tabella="ordini" then
						quantita=clng(rs("quantita"))
						
						if isnull(rs("quantita_magazzino") ) then
							stile=" color:blue;"
						else
							quantita_magazzino=clng(rs("quantita_magazzino"))
							
							if quantita>(quantita_magazzino+quantita_ordinato) then
								stile=" color:#ff0000;"
							elseif quantita>quantita_magazzino then
								stile=" color:#ff9900;"	'Arancione (Parte in ordine)
							else
								stile=" color:#66c466;" 'Verde (tutto a magazzino)
							end if
						end if
					end if
					giacenza="/"&CancellettoSeNull(rs("quantita_magazzino"))
					if quantita_ordinato>0 then giacenza=giacenza&"+"&quantita_ordinato
				else
					giacenza=""
				end if
				%>
            <td align="center" valign="middle" style="<%=stile%>">
	            
              <%if not m_modifica_articoli then %>
              <%=rs("quantita")%>
              <%else 
              if (m_spettanze=1 and cint(rs("quantita"))>0 and mod_larochelle) then 
              %>
				<%=rs("quantita")%>
              <%
              else
              %>
              <input type="text" name="quantita<%=rs("iddett")%>" size="3" value="<%=rs("quantita")%>" style="text-align:right;">
			  <%end if %>
  			  <input type="hidden" name="iddett<%=rs("iddett")%>" value="SI">
              <%end if%><%=giacenza%>
              </td>
            <td align="center" valign="middle" >
              <%if not m_modifica_articoli then %>
              <%=FormatNumber(rs("sconto_prodotto"),2)%>
              <%else %>
              <input type="text" name="sconto_prodotto<%=rs("iddett")%>" size="4" value="<%=FormatNumber(rs("sconto_prodotto"),2)%>" style="text-align:right;">
              <%end if%></td>
            <td align="right" valign="middle" nowrap >&#8364; <span class="prezzo_totale"><%=formatnumber(totale_riga(rs("prezzo"),rs("sconto_prodotto"),rs("quantita")),2)%></span></td>
          </tr>
		  <%if tr_ordini then %>
		  <tr <%if rs("modificato")=true then response.write ("bgcolor='#ffff99'")%>>
			<td colspan="5">
			<%=soloio("idddt:"&rs("iddett")&", idmag:"&rs("idmag")&" , quantita_scalato_magazzino:"&rs("quantita_scalato_magazzino")&", quantita_ordinato:"&rs("quantita_ordinato"))%>
			<%=txt_ordinato %>			
			</td>
		</tr>
		  <%end if%>
		  <%if tr_fornitori<>"" then %>
		  <tr <%if rs("modificato")=true then response.write ("bgcolor='#ffff99'")%>>
			<td colspan="5">
			<%
			response.write tr_fornitori
			%>
			</td>
		</tr>
		  <%end if%>
		  
		  
		  
		  
		 
          <%
            if vecchio_conteggio then
	            prezzo=cdbl(rs("prezzo"))
	            
				totord=totord+round((prezzo-(cdbl(rs("sconto_prodotto"))/100)*prezzo)*rs("quantita"),2)
	        else
				totord=totord+totale_riga(rs("prezzo"),rs("sconto_prodotto"),rs("quantita"))
			end if
            rs.movenext
			loop
			set rs = nothing
			%>
          </tbody>
          
          
          
		<%
			
			'http://it.wikihow.com/Calcolare-il-Margine-di-Contribuzione
			if tot_costo>0 then
				if tot_vendita=fattura_Dictionary.item("totale_merce") then
					m_txt_margine="Margine = &#8364;"&formatnumber(tot_vendita-tot_costo,2)&" Ricarico:"&formatnumber(calcola_ricarico(tot_vendita,tot_costo),0)&"%"
				else
					m_txt_margine="Margine calcolabile su &#8364;"&formatnumber(tot_vendita,2)&" = &#8364;"&formatnumber(tot_vendita-tot_costo,2)&" Ricarico:"&formatnumber(calcola_ricarico(tot_vendita,tot_costo),0)&"%"
				end if
			else
				m_txt_margine="Margine e ricarico non calcolabili, mancano prezzi di acquisto"
			end if
			
			if m_manca_numeri_di_serie then %>
			<script>
				$(function () {
					$(".stato2").addClass("ui-state-disabled");
					$("#aggiungi_seriali").show();
				});
			</script>
			
			
			
			<%end if 
		
	end sub	
'******************************************************************************************************************************************************************
'******************************************************************************************************************************************************************
	Public sub elenco_totali()
		'on error goto 0
		
	    m_log_txt=m_log_txt&",elenco_totali"
		%>
	<tbody id="tab_articoli_totali">

		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Totale merce: &#8364; <%=FormatNumber(fattura_Dictionary.item("totale_merce"),2)%>
			</td>
		</tr>
		<%if m_txt_margine<>"" then %>
		<tr valign="middle" align="right" class="mieidettagli">
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				<%=m_txt_margine%>
			</td>
		</tr>
		<%end if %>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Spese di trasporto (<%=tipo_trasporto(fattura_Dictionary.item("tipo_trasporto"))%>): &#8364;
				<%if not m_modifica_articoli then %>
					<%=FormatNumber(fattura_Dictionary.item("trasporto"),2)%>
				<%else %>
					<input type="text" name="trasporto" size="7" value="<%=FormatNumber(fattura_Dictionary.item("trasporto"),2)%>" style="text-align:right;" id="trasporto" class="i_text">
				<%end if%>
			</td>
		</tr>
		<%
		'totale_imponibile=fattura_Dictionary.item("totale_merce")+fattura_Dictionary.item("trasporto")
		totale_imponibile=fattura_Dictionary.item("totale_merce")'+fattura_Dictionary.item("trasporto")
		%>
		<%if fattura_Dictionary.item("spese_bancarie")>0 then
			totale_imponibile=totale_imponibile+fattura_Dictionary.item("spese_bancarie")
			%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Spese bancarie: &#8364; <span id="spese_bancarie"><%=FormatNumber(fattura_Dictionary.item("spese_bancarie"),2)%></span>
			</td>
		</tr>		  
		<%end if %>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Totale imponibile: &#8364; <span id="iva_esclusa"><%=FormatNumber(totale_imponibile,2)%></span>
			</td>
		</tr>
		<%
				
				
		'Calcolo IVA
		testo_esenzione=" "&replace(trattamento_iva(fattura_Dictionary.item("trattamento_iva")),"Normale","")
		if esenzione_iva(fattura_Dictionary.item("trattamento_iva")) then
			caso_iva=1
			valore_iva=0
			if fattura_Dictionary.item("trattamento_iva")=10 then
				caso_iva=2
				imposta=fattura_Dictionary.item("aliquota_iva")
				valore_iva=roundup(totale_imponibile*fattura_Dictionary.item("aliquota_iva")/100,2)
				testo_esenzione=""
			end if
		else
			caso_iva=3
			imposta=fattura_Dictionary.item("aliquota_iva")
			valore_iva=roundup(totale_imponibile*fattura_Dictionary.item("aliquota_iva")/100,2)
			totale_imponibile=totale_imponibile+valore_iva
		end if
		%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Imposta IVA <span id="iva"><%=fattura_Dictionary.item("aliquota_iva")%></span>%<%=testo_esenzione%>: &#8364; <span id="valore_iva"><%=FormatNumber(fattura_Dictionary.item("imposta"),2)%></span>
			</td>
		</tr>
		<%if fattura_Dictionary.item("trattamento_iva")=10 then%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				<%=trattamento_iva(fattura_Dictionary.item("trattamento_iva"))%>: &#8364; <span id="valore_iva">-<%=FormatNumber(fattura_Dictionary.item("imposta"),2)%></span>
			</td>
		</tr>
		<%end if%>


		<%if fattura_Dictionary.item("trattamento_iva")=10 then%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Totale con iva: &#8364; <%=FormatNumber(valore_iva+fattura_Dictionary.item("totale_fattura"),2)%>
			</td>
		</tr>
		<%end if%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><input type="hidden" name="versione" value="<%=fattura_Dictionary.item("versione")%>"><b>Totale: &#8364; <span id="totale_fattura"><%=FormatNumber(fattura_Dictionary.item("totale_fattura"),2)%></span></b></td>
		</tr>
		<%if utente_andrea  then%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><b>Versione: <%=fattura_Dictionary.item("versione")%></b></td>
		</tr>
		<%end if%>
	</tbody>
	<script>
		if (typeof doAutogrow == 'function') { 
				doAutogrow(); 
				console.log("autogrow");
		  }else
		  {
				console.log("noautogrow");
		}
	</script>
	
	
	
		<%
	end sub
	
	public sub tabella_stato()
	%>
	<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
	<tr>
		<td colspan="2" class="ui-widget-header">Stato: </td>
	</tr>
	<tr>
		<td width="20%" align="center" valign="bottom" ><input type="hidden" name="stato" value="" id="stato"/>
			<%=fattura_Dictionary.item("data_stato1")%><br>
			<button type="submit" name="" class="pulsanti ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only<%=disabilitato%>" value="<%=1%>" <%if fattura_Dictionary.item("stato_fattura")=1 then%> style="background:#66c466;" <%end if%> onClick="stato.value=<%=1%>; return;" >In modifica</button>
		</td>
	<td width="20%" align="center" valign="bottom" class="">
		<%=fattura_Dictionary.item("data_stato2")%><br>
		<button type="submit" name="" class="pulsanti ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only<%=disabilitato%>" value="<%=6%>" <%if fattura_Dictionary.item("stato_fattura")=6 then%> style="background:#66c466;" <%end if%> onClick="stato.value=<%=6%>; return;" >Chiusa</button></td>
	</tr>
	</table>

	
	
	
	<%
	end sub
		
'******************************************************************************************************************************************************************
'******************************************************************************************************************************************************************



	Public sub modifica_fattura()
		Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class
'		set modificheRS= (new ClasseModificheRS)(oper)

	
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql= "select * from fatture where idfat=" & m_idfat
		rs.Open sql, conn, 1, 3
		'memorizzo situazione iniziale
'		modifichers.leggi(rs)	
		tipo_fattura=rs("tipo_fattura")
		'response.write "CAUSALE:"&trim(request.form("causale"))
		'if request.form("data")<>"" then
			
			'call add2log(,0)
		'	if cdate(request.form("data"))>=cdate(request.form("data_minima")) and cdate(request.form("data"))<=rs("data") then
				rs("data")=request.form("data")
				rs("anno")=year(request.form("data"))
		'	end if
		'end if
		if request.form("new_nfat")<>"" then
			rs("nfat")=request.form("new_nfat")
		end if
		rs("causale")=cint(trim(request.form("causale")))
		if request.form("accompagnatoria")="si" then
			fine_mese=0
		else
			fine_mese=1
		end if
		rs("fine_mese")=fine_mese
		
	
		if fine_mese=0 then
			rs("peso")=trim(request.form("peso"))
			rs("dimensione")=trim(request.form("dimensione"))
			rs("imballo")=trim(request.form("imballo"))
			rs("porto")=trim(request.form("porto"))
		end if
		rs("annotazioni")=trim(request.form("annotazioni"))
		rs("pagamento")=trim(request.form("pagamento"))
		rs("banca_appoggio")=trim(request.form("banca_appoggio"))
		rs("iban")=trim(request.form("iban"))
		rs("incaricato")=request.form("incaricato")
		rs("spese_bancarie")=aggiusta_decimale(trim(request.form("spese_bancarie")),"asp")
		if request.form("iva_0")="si" then
			rs("iva_0")=true
			else
			rs("iva_0")=false
		end if
		if request.form("trattamento_iva")<>"" then
			rs("trattamento_iva")=request.form("trattamento_iva")
		end if
		rs("idbanca")=trim(request.form("idbanca"))
		if request.form("CodiceDestinatario")="" then
			rs("CodiceDestinatario")=NULL
		else
			rs("CodiceDestinatario")=request.form("CodiceDestinatario")
		end if
		rs.update
'		modifiche_rs=modificheRS.confronta(rs)	
		set modificheRS = Nothing
		if tipo_fattura="fattura" then
			call CalcolaTotale()
			
		else
			totale_fattura=calcola_totale_fattura(m_idfat)
		end if
		
			
		if tipo_fattura="fattura" then
			n_incassi=0
		else
			n_incassi=collega_incassi (idord,m_idfat)
		end if
		
		
		set rs_incassi=conn.execute ("select Sum(incassi.importo) AS SommaDiimporto FROM incassi WHERE (((incassi.idfat)="&m_idfat&"));")
		if not IsNull(rs_incassi("SommaDiimporto")) then
			SommaDiimporto=cdbl(rs_incassi("SommaDiimporto"))
		else
			SommaDiimporto=0
		end if
		
		add2log "Modificati dati [fattura="&m_idfat&"]"&rs("nfat")&"/"&year(rs("data"))&"[/fattura] eseguita da [admin="&session("iduser")&"]"&session("nominativo")&"[/admin],importo "&totale_fattura&", creato "&n&" scadenze riba, associato "&n_incassi&" incassi"&vbcrlf&modifiche_rs,2
		rs.close
		set rs = nothing
		oper="view"
	end sub





	
	
	
	
	
	public function aggiungi(aggiungo)
		dim rs
		dim ordine
		dim totale_riga_tmp
		dim testo_log

		'on error goto 0
	    m_log_txt=m_log_txt&",aggiungi"
			
			'call da_calcolare() 'Imposto calcolato=0
			
			'Calcolo ordine
			ordine=cint(request("ordine"))
			if ordine=-1 then
				SQL="select Max("&m_tabella&"_dett.ordine) AS ordine FROM "&m_tabella&"_dett WHERE idfat="&m_idfat&";"
				set rs=conn.execute(SQL)
				ordine=RS("ordine")
				rs.close
				Set rs = Nothing
				if isnull(ordine) then ordine=0
				ordine=ordine+2
			else
				ordine=ordine+1
			end if
		    m_log_txt=m_log_txt&",totale_merce prima:"&fattura_Dictionary.item("totale_merce")
			
			'Aggiungo articolo
			if aggiungo="articolo" then
				'on error resume next
				idpro=request("idpro")
				idvara=request("idvara")
				idvarb=request("idvarb")
				sconto=request("sconto")
				iddip=request("iddip")
				quantita=request("quantita")
				prezzo=request("prezzo")
				'call add2log("aggiungi-prezzo"&prezzo,0)
				
				totale_riga_tmp=aggiungi_articolo(idpro,idvara,idvarb,sconto,quantita,iddip,ordine, prezzo, testo_log)
				
				
			elseif aggiungo="vocelibera" then
				'on error resume next
				sql="select * from "&m_tabella&"_dett "
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open sql, conn, 3, 3
				rs.addnew
				rs("articolo_ordine")=Replace(request.Form("articolo_ordine"),"""","&quot;")
				
				if request.Form("quantita")<>"" then
					rs("quantita")=request.Form("quantita")
				else
					rs("quantita")=1
				end if
				if request.Form("prezzo")="" then
					rs("prezzo")=0
				else
					rs("prezzo")=aggiusta_decimale(request.Form("prezzo"),"asp")
				end if
				rs("idord")=idord
				rs("modificato")=1
				rs("ordine")=ordine
				rs("um")=request.Form("um")
				totale_riga_tmp=totale_riga(rs("prezzo"),rs("sconto_prodotto"),rs("quantita"))
				rs("totale_riga")=totale_riga_tmp
				if m_tabella<>"ordini_fornitori" then rs("costo_ordine")=0
				articolo_ordine=rs("articolo_ordine")
				quantita=rs("quantita")
				rs.update
				iddett=get_last_id(m_tabella&"_dett")

				testo_log="Aggiunto riga "&iddett&" libera "&articolo_ordine&" quantita: "&quantita&" prezzo totale: "&totale_riga_tmp
				rs.close
				Set rs = Nothing
				'Aggiungo spettanze su riga libera
				iddip=request("iddip")
				quantita=cdbl(quantita)
				if quantita>0 then
					call aggiungi_spettanza(iddett,iddip,0,"", quantita, testo_log)
				end if
				
				
				'add2log "Ordine aggiornato",0
				esito="Voce libera " & request("articolo_ordine") & " aggiunta."
				if Err.Number <> 0 then 
					txt_errore="Err.Description:"&Err.Description& "<br><br>"
					add2log "Errore in aggiungi voce libera" & txt_errore,0
				end if
			   'On Error GoTo 0
			
			else
				call add2log("Non so cosa aggiungere"&vbcrlf&queryeform(),3)
			end if
			
			
			
			call riordina_dettaglio (idord,m_tabella)
			
			fattura_Dictionary.item("totale_merce")=fattura_Dictionary.item("totale_merce")+totale_riga_tmp
		    m_log_txt=m_log_txt&",totale_merce dopo:"&fattura_Dictionary.item("totale_merce")
			n=CalcolaTotaleMerce(false)
						'sql="UPDATE "&m_tabella&" SET comunicazioni_precedenti= CONCAT(comunicazioni_precedenti,'<hr>"& now()&" <b>"&session("nominativo")&":</b><br>"&replace(testo_log,"'","''")&"') WHERE idord = "&m_idfat&";"
			'conn.execute (sql)
			
			call add2log ("Modificato "&m_cosa_log(m_tabella,m_idfat,m_nord)&vbcrlf&testo_log&vbcrlf&"totale_merce: "&fattura_Dictionary.item("totale_merce")&vbcrlf&"Versione finale:"&fattura_Dictionary.item("versione"),2)
	end function	'aggiungi(aggiungo)
		
		
		
	private function aggiungi_articolo(idpro,idvara,idvarb,sconto,quantita,iddip,ordine,prezzo, testo_log) 'Restituisce l'importo
				dim raggruppa
				raggruppa=false	'Flag raggruppamento articoli nell'ordine
				dim tipo_taglia
				tipo_taglia=null
				dim sql_prodotti
				sql_prodotti=""
				
				
				sqlString = "select iddett,prezzo FROM "&m_tabella&"_dett WHERE idfat=" & m_idfat& " and idpro=" & idpro  
				if idvara<>0 then
					sqlString=sqlString & " AND idvara=" & idvara 
					sql_prodotti=sql_prodotti &" AND varianti_a.IDvara="& idvara
				end if
				if idvarb<>0 then
					sqlString=sqlString & "  AND idvarb=" & idvarb 
					sql_prodotti=sql_prodotti &" AND varianti_b.IDvarb="& idvarb
				end if
				
				if mod_dip=true then
					raggruppa=true

					if prezzo<>"" then
						
						sqlString= sqlString & " and prezzo="& aggiusta_decimale( prezzo,"sql")
					end if
					
				end if
				
				'response.write sqlString
				SET RS = Conn.Execute( sqlString )
				iddett=0
				if not rs.eof then
					iddett=rs("iddett")
					testo_esiste="Articolo gi&agrave; presente nell'ordine."
					prezzo_esiste=rs("prezzo")
				end if
				rs.close
				
				
				IF raggruppa=false or iddett=0  THEN 'IF RS.EOF THEN  'MODIFICA PER RAGGRUPPAMENTO ARTICOLI
					sql_prodotti="select prodotti.codice, prodotti.articolo, prodotti.prezzo, prodotti.costo, prodotti.sconto, prodotti.promozione, prodotti.prodata, prodotti.variante1, prodotti.variante2, prodotti.aggiungi_a_ordine, prodotti.richiedi_seriale, prodotti.um, varianti_a.variante_a, varianti_a.prezzo_ve_va, varianti_a.prezzo_ac_va, varianti_b.variante_b, prodotti.tipo_taglia FROM (prodotti LEFT JOIN varianti_a ON prodotti.IDpro = varianti_a.idpro) LEFT JOIN varianti_b ON prodotti.IDpro = varianti_b.idpro WHERE prodotti.IDpro= "&idpro&sql_prodotti
					set rs_pro=conn.execute(sql_prodotti)
					tipo_taglia=rs_pro("tipo_taglia")
					
					sql="select * from "&m_tabella&"_dett"
					rs.Open sql, conn, 3, 3
					rs.addnew
					'Imposto campi articolo
					rs("idfat")=m_idfat
					rs("idpro")=idpro
					rs("idvara")=idvara
					rs("idvarb")=idvarb
					
					
					
					'Imposto prezzi e costo
					prezzo_ve_va=converti_typevar14(rs_pro("prezzo_ve_va"))
					prezzo_ac_va=converti_typevar14(rs_pro("prezzo_ac_va"))
					
					if prezzo_ve_va>0 then prezzodb=rs_pro("prezzo_ve_va") else prezzodb=rs_pro("prezzo")
					prezzodb=cdbl(prezzodb)

					if isnull(prezzo_ve_va) or prezzo_ve_va=0 then
						passo="a"	'Togliere
						prezzo_ac=rs_pro("costo")
					elseif prezzo_ac_va>0 then
						passo="b"	'Togliere
						prezzo_ac=rs_pro("prezzo_ac_va")
					else
						passo="c"	'Togliere
						prezzo_ac=rs_pro("costo")
					end if
					if m_tabella<>"ordini_fornitori" then
						rs("costo_ordine")=prezzo_ac
					end if
					if m_tabella="ordini_fornitori" then
						prezzodb=prezzo_ac
						passo2="a"	'Togliere
					elseif rs_pro("promozione")=1 and (rs_pro("prodata")>date() or isnull(rs_pro("prodata"))) then 
						passo2="b"	'Togliere
						prezzodb=roundup(prezzodb*(1-rs_pro("sconto")/100),2)
					else
						passo2="c"	'Togliere
						'prezzo=rs_carrello("prezzo")
					end if				
					if isnull(prezzo) or prezzo="" then
						rs("prezzo")=prezzodb
					else
						rs("prezzo")=aggiusta_decimale( prezzo,"asp")
					end if
					
					if m_tabella="ordini" then
						rs("numeri_di_serie")=conteggia_numeri_di_serie(idpro,idvara,idvarb,0)
					end if

					
					
					
					rs("um")=rs_pro("um") 
					rs("quantita")=quantita 
					rs("modificato")=true
					rs("codice_ordine")=rs_pro("codice")
					rs("articolo_ordine")=Replace(rs_pro("articolo"),"""","&quot;")
					rs("variante1_ordine")=rs_pro("variante1")
					rs("variante2_ordine")=rs_pro("variante2")
					rs("var1")=rs_pro("variante_a")
					rs("var2")=rs_pro("variante_b")
					rs("varianti_ordine")=codifica_varianti_ordine(rs_pro("variante_a"), rs_pro("variante1"), rs_pro("variante_b"), rs_pro("variante2"))
					rs("ordine")=ordine
					totale_riga_tmp=totale_riga(rs("prezzo"),rs("sconto_prodotto"),rs("quantita"))
					rs("totale_riga")=totale_riga_tmp
					articolo_ordine=rs("articolo_ordine")
					if sconto<>"" then
						rs("sconto_prodotto")=aggiusta_decimale(sconto,"asp")
					else
						rs("sconto_prodotto")=0
					end if
					
					
					
					
					rs.update
					iddett=get_last_id(m_tabella&"_dett")
					
					
					esito="Articolo " & articolo_ordine&" agiunto."
					aggiungi_a_ordine=rs_pro("aggiungi_a_ordine")
					rs_pro.close
					set rs_pro=Nothing
					
					if aggiungi_a_ordine>0 and m_tabella="ordini" then
						voci_libere=0
						sql="select aggiungi_a_ordine_dett.* from aggiungi_a_ordine_dett where idgruppo="&aggiungi_a_ordine
						set rs_aggiungi=conn.execute(sql)
						if rs_aggiungi.eof then
							testo_log="<b>Voci libere per gruppo "&aggiungi_a_ordine&" non trovate</b>"
						else
							for n=1 to cint(quantita)
								
								do while not rs_aggiungi.eof
									voci_libere=voci_libere+1
									rs.addnew
									rs("idord")=idord
									'rs("idpro")=""
									rs("prezzo")=0
									'rs("um")=request("um") 
									rs("quantita")=rs_aggiungi("quantita") 
									rs("modificato")=true
									rs("sconto_prodotto")=0
									rs("codice_ordine")=""
									rs("articolo_ordine")=rs_aggiungi("nome_voce")
									rs("ordine")=ordine
									rs("totale_riga")=0
									rs.update
									rs_aggiungi.movenext
								loop
								rs_aggiungi.movefirst
								n=n+1
							next
						end if
						if voci_libere>0 then testo_log=", aggiunto "&voci_libere&" voci a ordine"
					end if
					testo_log="Aggiunto riga "&iddett&" articolo "&articolo_ordine&" quantita: "&quantita&" prezzo totale: "&totale_riga_tmp&testo_log
				else 'IF raggruppa=true  'MODIFICA PER RAGGRUPPAMENTO ARTICOLI
					sql="select * from "&m_tabella&"_dett where iddett="&iddett
					rs.Open sql, conn, 3, 3
					rs("quantita")=quantita+cdbl(rs("quantita"))
					testo_log=testo_log&"Aggiungo a riga esistente "&iddett&" "&rs("articolo_ordine")&" quantita "&quantita
					rs.update
				END IF 'IF raggruppa=false  'MODIFICA PER RAGGRUPPAMENTO ARTICOLI
				rs.close
				
				
				
				'Aggiungi_spettanza
				call aggiungi_spettanza(iddett,iddip,idpro,tipo_taglia, quantita, testo_log)
				'------------------------
				
				
				if Err.Number <> 0 then 
					txt_errore="Err.Description:"&Err.Description& "<br><br>"
					add2log "Errore in aggiungi articolo a ordine fase 1" & txt_errore,0
				end if
				'on error goto 0
				'on error resume next
				Set rs = Nothing
				'add2log "Ordine aggiornato",0
				if Err.Number <> 0 then 
					txt_errore="Err.Description:"&Err.Description& "<br><br>"
					add2log "Errore in aggiungi articolo a ordine fase 2" & txt_errore,0
				end if
			   'On Error GoTo 0

		
		
		aggiungi_articolo=totale_riga_tmp
		
	end function
	private sub aggiungi_spettanza(iddett,iddip,idpro,tipo_taglia, quantita, byref testo_log)
		
				'Aggiungo riferimento dipendente (spettanze)
				if (m_tabella="ordini" or m_tabella="preventivi") and mod_dip then
					'if isnumeric(iddip) then
						'fattura_Dictionary.item("spettanze")=1
						if not isnumeric(iddip) then
							iddip=0
						end if
						if iddip>0 and idpro>0 then
							if isnull(tipo_taglia)  then
								tipo_taglia=conn.execute("select tipo_taglia from prodotti where idpro="&idpro)(0)
							end if
							testotaglia=recupera_taglia(tipo_taglia,iddip)
						end if
						if vartype(testotaglia)=2 then
							testotaglia=""
						end if
						dim rs_spettanze
						Set rs_spettanze = Server.CreateObject("ADODB.Recordset")
						sql="select "&m_tabella&"_dett_spettanze.* from "&m_tabella&"_dett_spettanze where iddett="&iddett &" and iddip="&iddip
						rs_spettanze.Open sql, conn,3,3
						if rs_spettanze.eof then
							rs_spettanze.addnew
							rs_spettanze("iddett")=iddett
							rs_spettanze("iddip")=iddip
							rs_spettanze("quantita")=quantita
							rs_spettanze("taglia_misura")=testotaglia
							
						else
							rs_spettanze("quantita")=rs_spettanze("quantita")+quantita	
						end if
						rs_spettanze.update
						rs_spettanze.Close
						set rs_spettanze = Nothing
						
						if fattura_Dictionary.item("spettanze")=0 then
							fattura_Dictionary.item("spettanze")=1
						    conn.execute ("update ordini set spettanze=1 where idord="&m_idfat)
					    end if
					    testo_log=testo_log&" per dipendente "&dipendente_log(iddip)
					'end if
				end if
	end sub
		
	
		
	private sub copia_da(tabella,idord)
				set rs_da=conn.execute ("select * from "&tabella&" where idord="&idord)
				Set rs_a = Server.CreateObject("ADODB.Recordset")
				
				rs_a.Open "select * from "&m_tabella, conn, 3, 3
				rs_a.addnew
				n=rs_da.fields.count-1
				nord_da=rs_da("nord")
				
				m_cosa_log_da=m_cosa_log(tabella,idord,nord_da)

				add2log "step1",1
				
				'on error resume next
				'on error resume next
				for i=1 to n
					'response.write rs_da.fields(i).name&":"&converti_typevar14( rs_da(rs_da.fields(i).name))&"<br>"
					rs_a(rs_da.fields(i).name)=converti_typevar14( rs_da(rs_da.fields(i).name))
					if err.number<>0 then
						call add2log( "Errore nella copia da "&tabella&":"&rs_da.fields(i).name & " a "&m_tabella&":"&rs_da.fields(i).name,0)
						Err.Clear
					end if
				next
				on error goto 0
				'Sovrascrivo i campi che servono
				rs_a("comunicazioni_precedenti")=now()&" <b>"&session("nominativo")&"</b> creato da "&tabella_singolare(tabella)&" "&nord_da&"<hr />"
				if mod_larochelle then
					rs_a("stato")=1
				else
					rs_a("stato")=0
				end if
				rs_a("data")=now()
				rs_a("aliquota_iva")=iva(date())
				anno=year(date())
				m_nord=max_ord("ordini",anno)
				rs_a("anno")=anno
				rs_a("nord")=m_nord
				rs_a("versione")=0
				rs_a("calcolato")=0
				if m_tabella="ordini" then
					rs_a("tipo_documento")="ordine"
					if mod_larochelle then
						rs_a("spettanze")=1
					end if
				end if
				
				rs_a.update
				m_idfat=Get_last_id(m_tabella)
				call add2log("step2 m_idfat:"&m_idfat,0)
				rs_a.close
				set rs_a=nothing
				
				
				call copia_articoli_spettanze(idord, tabella)
				
				'ApriRecord("totali")
		        'm_cosa_log=replace(tabella_singolare(m_tabella)," ","_")
		        'm_cosa_log="["&m_cosa_log&"="&m_idfat&"]"&m_nord&"[/"&m_cosa_log&"]"

				add2log "Creato "&m_cosa_log(m_tabella,m_idfat,m_nord)&" da "&m_cosa_log_da&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]",2
				call ApriRecord("totali")
				call CalcolaTotaleMerce(false)
				'CalcolaTotale()
	end sub
	
	public sub copia_articoli_spettanze(da_idord, da_tabella)
				Set rs_dett_pre = Server.CreateObject("ADODB.Recordset")
				Set rs_dett_ord = Server.CreateObject("ADODB.Recordset")
				sql1="select "&da_tabella&"_dett.* from ("&da_tabella&"_dett left join prodotti on "&da_tabella&"_dett.idpro= prodotti.idpro ) where idord="&da_idord&";"
				rs_dett_pre.Open sql1, conn, 3, 3
				
				
				
				sql2="select * from "&m_tabella&"_dett;"
				rs_dett_ord.Open sql2, conn, 3, 3
				'call add2log("sql rs_dett_pre:"&sql1&" sql  rs_dett_ord:"&sql2,0)
				n=rs_dett_pre.fields.count-1
				do while not rs_dett_pre.eof
				rs_dett_ord.addnew
				for i=1 to n
					'on error resume next
					rs_dett_ord(rs_dett_pre.fields(i).name)=converti_typevar14(rs_dett_pre(rs_dett_pre.fields(i).name))
					if err.number<>0 then
						'response.write "errore nella copia del campo:"&	rs_dett_pre.fields(i).name
						call add2log("errore nella copia del campo:"&	rs_dett_pre.fields(i).name &" con valore "&rs_dett_pre(rs_dett_pre.fields(i).name)&"sql rs_dett_pre:"&sql1&" sql  rs_dett_ord:"&sql2,0)
						Err.Clear
					end if
				next
				on error goto 0
				rs_dett_ord("idord")=m_idfat
				rs_dett_ord.update
				if mod_larochelle and m_tabella="ordini" then
					'copio le spettanze
					iddett_da=rs_dett_pre("iddett")
					iddett_a=get_last_id(m_tabella&"_dett")
					
					
					sql="INSERT INTO "&m_tabella&"_dett_spettanze (iddett, iddip, quantita, note, taglia_misura)  SELECT "&iddett_a&", iddip, quantita, note, taglia_misura FROM "&da_tabella&"_dett_spettanze WHERE iddett = "&iddett_da
					conn.execute(sql)
					
					
					
					
					
					
					
				end if
				
				
				
				rs_dett_pre.movenext
				loop
				rs_dett_pre.close
				rs_dett_ord.close
				set rs_dett_pre=nothing
				set rs_dett_ord = nothing

		
		
	end sub
	
	private sub riordina_dettaglio(idord,tabella)
		sql="select * from "&tabella&"_dett where idfat="&m_idfat&" order by ordine, iddett"
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

	
	
	
	
	private sub sposta_copia_articoli(operazione,nuovo_ordine)
		dim v_elenco_articoli
		m_log_txt=m_log_txt&"sposta_copia_articoli(operazione:"&operazione&")"&" stringa_iddett:"&stringa_iddett&"<br>"

		'stringa_iddett=request.form("stringa_iddett")
		
		sql = "select * FROM "&m_tabella&"_dett WHERE iddett in(" & stringa_iddett &") and idord="&m_idfat
		if operazione="D" then
			sql = "select * FROM "&m_tabella&"_dett WHERE idord="&m_idfat
		end if
		
		sql_new = "select * FROM "&m_tabella&"_dett"
		Set rs = Server.CreateObject("ADODB.Recordset")
		
		set rs_new= Server.CreateObject("ADODB.Recordset")
		'on error resume next
		rs.Open sql, conn, 3, 3
		if err.number<>0 then
			m_log_txt=m_log_txt&"Errore in sposta_copia_articoli:"&sql
			response.write m_log_txt&"<br>"
			
		end if
		m_log_txt=m_log_txt&"Trovati "&rs.fields.count-1&" campi in rs<br>"

		n=rs.fields.count-1
		'dim campo(50)
		dim campo()
		redim campo(n+2)
		on error goto 0
		rs_new.Open sql_new, conn, 3, 3
		'elenco_articoli=""
		WHILE NOT RS.EOF
		IF operazione="S" THEN
			rs("idord")=nuovo_ordine
			rs.update
			v_elenco_articoli=v_elenco_articoli&rs("codice_ordine")&" "
		END IF
		IF operazione="C" or operazione="D" THEN 'Copia
			n=rs.fields.count-1
			for i=1 to n
				campo(i)=rs(i)
			next
			rs_new.addnew
			for i=1 to n
				rs_new(i)=campo(i)
			next
			rs_new("idord")=nuovo_ordine
			rs_new.update
			v_elenco_articoli=v_elenco_articoli&rs("codice_ordine")&" "
		END IF
		RS.MoveNext
		WEND
		RS.Close
		rs_new.close
		set rs_new=nothing
	
		if operazione="C" or operazione="S" then
			sql="select * from "&m_tabella&" where idord="&m_idfat
			rs.Open sql, conn,3,3
			rs("comunicazioni_precedenti")= rs("comunicazioni_precedenti")&"<br>articoli spostati: "&v_elenco_articoli
			if operazione="S" then
				rs("calcolato")=0
			end if
			rs.Update 
			rs.Close
			sql="select * from "&m_tabella&" where idord="&nuovo_ordine
			rs.Open sql, conn,3,3
			rs("comunicazioni_precedenti")= rs("comunicazioni_precedenti")&"<br>articoli spostati: "&v_elenco_articoli
			rs.Update 
			rs.Close
			set rs=nothing
		end if
		if operazione="S" then
			'Ricalcolo ordine sorgente
			m_log_txt=m_log_txt&">>Ricalcolo ordine sorgente idord:"&m_idfat&"<br>"
			call ApriRecord("totali")
			call CalcolaTotaleMerce(false)
		end if
		'Calcolo ordine destinazione
		m_idfat=nuovo_ordine
		m_log_txt=m_log_txt&">>Ricalcolo ordine duplicato idord:"&m_idfat&"<br>"
		call ApriRecord("totali")
		call CalcolaTotaleMerce(false)
		
	end sub
	
	private function duplica()
		m_log_txt=m_log_txt&"duplica()<br>"
		Set rs_from = Server.CreateObject("ADODB.Recordset")
		Set rs_to = Server.CreateObject("ADODB.Recordset")
		sql="select * from "&m_tabella&" where idord="&m_idfat
		'on error resume next
		rs_from.Open sql, conn,3,3
		if err.number<>0 then
			response.write "errore query:"&sql
		end if
		on error goto 0
		'Prova
		nord_from=rs_from("nord")
		sql="select * from "&m_tabella
		rs_to.Open sql, conn, 3, 3
		rs_to.addnew
	    for i=1 to rs_to.Fields.Count-1
	        'Response.Write(Rs.Fields(i).Name + "<br>")
			rs_to(rs_to.Fields(i).Name)=rs_from(rs_to.Fields(i).Name)
		next
		comunicazioni_precedenti= now()&" <b>"&session("nominativo")&":</b> creato ordine per articoli spostati da "&tabella_singolare(tabella) &" "&nord_from
		rs_to("comunicazioni_precedenti")= comunicazioni_precedenti
		rs_to("creato_da_admin")=sessionIDUser
		rs_to("note_gestore")=""
		nord_to=max_ord(m_tabella,year(date()))
		rs_to("nord")=nord_to
		rs_to("data")=now()
		rs_to("anno")=year(date())
		rs_to("modificato")=1
		rs_to("versione")=1
		rs_to("calcolato")=0
		rs_to.update
		duplica=get_last_id(m_tabella)
		comunicazioni_precedenti= now()&" <b>"&session("nominativo")&":</b> articoli spostati a "&tabella_singolare(tabella)&" "&nord_to
		rs_from("comunicazioni_precedenti")=rs_from("comunicazioni_precedenti")&"<hr />"& comunicazioni_precedenti
		rs_from.update
		rs_to.close
		rs_from.close
		set rs_from=nothing
		set rs_to=nothing
		m_log_txt=m_log_txt&">>duplicato "&m_tabella&" "&m_idfat&" su "&duplica&"<br>"
	end function


	Public sub dati_fatturazione_consegna()
		%>
          <tr>
            <td bgcolor="#E5E5E5">Dati fatturazione</td>
            <td width="50%" bgcolor="#E5E5E5">Dati consegna</td>
          </tr>
          <tr >
            <td width="50%" valign="top" style="border-bottom:1px solid;">
	            <%=dati_fatturazione(fattura_Dictionary("idintestazione"))%><br>
              Pagamento: <%=metodo_pagamento(fattura_Dictionary("tipopagamento"),"descrizione")%>
  				<%if fattura_Dictionary("idagente")>0 then %>
				<br >
				Agente: <%=conn.execute("select nominativo from agenti where idagente="&fattura_Dictionary("idagente"))(0)%>
				
				<%end if %>

              </td>
            <td width="50%" valign="top" style="border-bottom:1px solid;">
				<%=dati_consegna(fattura_Dictionary("idconsegna"))%>
                <br />Metodo spedizione: <%=tipo_trasporto(fattura_Dictionary("tipo_trasporto"))%>
              </td>
          </tr>
		<%
		if fattura_Dictionary("note_spedizione")<>"" then%>
		<tr class="ui-widget-header">	
			<td colspan="2">Note spedizione:</td>
		</tr>
		<tr>
			<td colspan="2"> <%=fattura_Dictionary("note_spedizione")%></td>
		</tr>
		<%
		end if
	end sub
	
	Public SUb dati_mepa()
		
		if fattura_Dictionary.item("cig")<>"" or fattura_Dictionary.item("mepa_tipo")>0 or fattura_Dictionary.item("mepa_data")<>"" then
			
        if fattura_Dictionary.item("cig")<>"" then
	        cig="Codice CIG: "&fattura_Dictionary.item("cig")
	    else
		    cig=""
		end if
        if fattura_Dictionary.item("mepa_tipo")>0 then
	        mepa_tipo_v=mepa_tipo( fattura_Dictionary.item("mepa_tipo"))&": "&fattura_Dictionary.item("mepa_testo")
	    else
		    mepa_tipo_v=""
		end if
        if fattura_Dictionary.item("mepa_data")<>"" then
	        mepa_data_v="Data ordine: "&fattura_Dictionary.item("mepa_data")
	    else
		    mepa_data_v=""
		end if
        
        %>
          <tr>
	          <td colspan="2">
		          <table border="0" cellpadding="0" cellspacing="0" width="100%">
			          <tr>
				          <td width="33%"><%=cig%></td>
				          <td width="33%"><%=mepa_tipo_v%></td>
				          <td width="33%"><%=mepa_data_v%></td>
			          </tr>
		          </table>
          </tr>
          <%
	    end if

		if fattura_Dictionary.item("impegno_spesa")<>"" then%>
          <tr>
	          <td>Impegno di spesa: <%=fattura_Dictionary.item("impegno_spesa")%></td>
	          <td></td>
          </tr>
          <%
	    end if
		
		
	end sub
		
	Private function m_cosa_log(tabella,idord,nord)
	        m_cosa_log=replace(tabella_singolare(tabella)," ","_")
	        m_cosa_log="["&m_cosa_log&"="&idord&"]"&nord&"[/"&m_cosa_log&"]"
	end function
	
	Private sub da_calcolare()
		if fattura_Dictionary.item("calcolato")=1 then
			conn.execute ("UPDATE "&m_tabella&" SET calcolato = 0 WHERE idord = "&m_idfat&";")
			fattura_Dictionary.item("calcolato")=0
		end if
	end sub
	
	Private sub imposta_stato_articoli()	
		'Non più usata
		dim sql
		Set rs_dettaglio = Server.CreateObject("ADODB.Recordset")
		'Cuclo su spettanze ordine
		sql="select fatture_dett.iddett,id from fatture_dett inner join fatture_dett_spettanze on fatture_dett.iddett = fatture_dett_spettanze.iddett where idord="&m_idfat
		set rs= conn.execute(sql)
		do while not rs.EOF
		quantita=request.form("quantitas_"&rs("id"))
			if quantita<>"" then
				
				rs_dettaglio.open "select * from fatture_dett_spettanze_dettaglio where iddett="&rs("iddett")&" and idspettanza="&rs("id")&" and stato="&request.form("stato_articolo"), conn, 3,3
				if rs_dettaglio.eof then
					rs_dettaglio.addnew
				end if
				rs_dettaglio("iddett")=rs("iddett")
				rs_dettaglio("quantita")=quantita
				rs_dettaglio("stato")=request.form("stato_articolo")
				rs_dettaglio("idspettanza")=rs("id")
				rs_dettaglio("data")=now()
				rs_dettaglio.update
				rs_dettaglio.close
				'sql="INSERT INTO fatture_dett_spettanze_dettaglio( iddett, idspettanza,stato, quantita, data) VALUES ("&rs("iddett")&","&rs("id")&","&request.form("stato_articolo")&","&quantita&",now())"
								'call add2log(sql,0)
				'conn.execute(sql)
			end if
			rs.MoveNext
		loop
	end sub
	Private sub aggiorna_quantita_spettanze(idddt)
		dim sql,rettifica
		
		if request.form("rettifica")<>"" then
			rettifica=true
			testo_rollback="Antecedente rettifica quantit&agrave;"
			
		else
			rettifica=false
			testo_rollback="Antecedente modifica quantit&agrave;"
		end if	
		
		call memorizza_rollback(m_idfat,testo_rollback)
		
		
		call add2log("Inizio aggiorna quantità spettanze per idord:"&m_idfat&" idddt:"&idddt,0)	
		iddett=0
		tot_in_consegna=0
		Set rs_dettaglio = Server.CreateObject("ADODB.Recordset")
		'Cuclo su spettanze ordine
		sql="select id,fatture_dett.iddett from fatture_dett inner join fatture_dett_spettanze on fatture_dett.iddett = fatture_dett_spettanze.iddett where idord="&m_idfat
		set rs= conn.execute(sql)
		do while not rs.EOF
			id=rs("id")
			if iddett=0 then 
				iddett=rs("iddett")
			end if
			
			in_ordine=request.form("in_ordine"&id)
			pronto=request.form("pronto"&id)
			in_consegna=request.form("consegna"&id)
			if pronto<>"" or in_ordine<>"" or in_consegna<>"" then
				in_ordine=zerosenull(in_ordine)
				pronto=zerosenull(pronto)
				quantita=cint(conn.execute("select quantita from fatture_dett_spettanze where id="&id)(0))
				
				sql="update fatture_dett_spettanze set in_ordine="&in_ordine&", pronto="&pronto
				if in_consegna<>"" then
					
					
					if cint(in_consegna)>=quantita then
						in_ordine=0
						pronto=0
					end if
					
					if rettifica then
						sql=sql&", consegnato="&in_consegna
					else
						sql=sql&", consegnato=consegnato+"&in_consegna
						if in_consegna>0 then
							tot_in_consegna=tot_in_consegna+in_consegna
						end if
						
						
					end if
				end if
				conn.execute (sql&" where id="&id)
			end if
			
			
			rs.MoveNext
			if idddt>0 then 'solo per ddt
				if not rs.eof then
					if rs("iddett")<>iddett then
						
						
						call aggiungi_ddt_dett_ordini(idddt,iddett,tot_in_consegna)
						
						tot_in_consegna=0
						iddett=0
					end if
					
				else
					call aggiungi_ddt_dett_ordini(idddt,iddett,tot_in_consegna)
					
				end if
			end if
			
			
		loop
		call aggiorna_totali_spettanze()
		call add2log("aggiorna_quantita_spettanze &idord"&m_idfat,0)
	end sub
	
	Public SUb avviso_eliminato()
	if 	ordine.campo("eliminato")=1 then
	%>
 		<div class="ui-widget" id="report_errore">
			<div class="ui-state-error ui-corner-all" style="padding: 0 .7em;"> 
				<p>
					<span class="ui-icon ui-icon-alert" 
						style="float: left; margin-right: .3em;"></span>
					<strong>Questo <%=tabella_singolare(m_tabella)%> &egrave; eliminato</strong> 
				</p>
				<p>
					Verr&agrave; cancellato definitivamente dal database tra <%=180+DateDiff("d",date(),ordine.campo("data"))%> giorni<br>
					<input type="submit" name="recupera" value="Ripristina <%=tabella_singolare(m_tabella)%>">
				</p>
			</div>
		</div>
	<%end if 
	end sub

	
	Private function max_ord(tabella,anno)
		dim rs_o
		SQL="select Max(Nord) AS MAX_Nord FROM "&tabella&" WHERE"
		if m_tipo_documento="ordine" or m_tipo_documento="ddt" or m_tipo_documento="notacredito" or m_tipo_documento="sostituzione" then
			SQL=SQL&" tipo_documento='"&m_tipo_documento&"' and"
		elseif tabella="ordini" then
			SQL=SQL&" tipo_documento='ordine' and"
		end if
		
		sql=sql&" anno="&anno&";"
		'ssql=sql
		'call add2log(ssql,0)
		set rs_o=conn.execute(SQL)
		MAX_Nord=rs_o("MAX_Nord")
		if isnull(MAX_Nord) then MAX_Nord=0
		max_ord=MAX_Nord+1
		set rs_o=nothing
	end function

	
	public sub elenco_articoli_head()
		%>
		<script>
			var aggiungi_a='<%=m_tabella%>';
			var aggiungi_a_id=<%=m_idfat%>;
		</script>
		<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1" id="tab_articoli">
		<thead id="tab_articoli_testa">
			<%if utente_andrea then %>
			<tr>
				<td colspan="6" >m_tabella:<%=m_tabella%> m_idfat:<%=m_idfat%> m_tipo_ddt:<%=m_tipo_ddt%> </td>
			</tr>
			<%end if %>
			<tr>
				<td colspan="6" class="ui-widget-header">Elenco articoli:<input type="button" id="mieidettagli" value="Dettagli"></td>
			</tr>
			<tr>
				<td align="center" class=testotabella style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;"><b>Codice</b></td>
				<td style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Articolo</b></td>
				<td align="center" class=testotabella style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;"><b>Prezzo<br>unitario</b></td>
				<td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Quantit&agrave;</b></td>
				<td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Sconto %</b></td>
				<td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Prezzo<br>totale</b></td>
			</tr>
		</thead>
		<%
	end sub
	Public sub riga_pulsanti(ordina_fornitore)
		if not m_modifica_articoli and not ordina_fornitore then
			exit sub
		end if
		%>
		
		<tr valign="middle" >
            <td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
			<%if ordina_fornitore then %>
				<input type="submit" value="Ordina a fornitore" id="ordina_fornitore" class="oper_articoli">
			<%end if %>

			<%
				if m_modifica_articoli then%><span style="float: right;"><input type="button" value="Aggiungi articolo" id="puls_aggiungi_articolo">
			<input type="submit" name="Aggiorna_ordine" id="Aggiorna_ordine" value="Aggiorna dati"></span>
			<%end if%>
			<input type="hidden" name="idfat" value="<%=m_idfat%>">
            </td>
        </tr>
		<%
	end sub
		
		
		
		

		
'******************************************************************************************************************************************************************
'******************************************************************************************************************************************************************
			

	Public Property Get modifica_articoli()
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		modifica_articoli=m_modifica_articoli
	end property
	Public Property Let modifica_articoli(value)
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		m_modifica_articoli=value
	end property
	Public Property Get idfat()
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		idfat=m_idfat
	end property
	Public Property Get nfat()
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		nfat=m_nfat
	end property
	
	
		
	'-----------------------------------------------------------------------------------------------------------------------------------------------
	'  												FINE CLASSE
	'-----------------------------------------------------------------------------------------------------------------------------------------------
End Class





	
	



%>