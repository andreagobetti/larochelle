<%
'Sintassi

'set ordine= (new ClasseOrdine)(array("ordine","new"))
'set ordine= (new ClasseOrdine)(array("preventivi",752))
'set ordine= new ClasseOrdine



'Spostare qui operazioni da fare con pulsanti


m_log_txt=""
Class ClasseOrdine
	
	
	
	
    Private m_idord
    Private m_nord
    Private m_tabella
    Private m_iduser
	Private m_n_campi
	Private m_casoiva
	Private m_totale_ok
	Private m_nuovo
	Private m_totale_merce_ok
	Private m_nominativo
	Private txt_log
	
	Private m_txt_margine
	
	
	
	Private m_modifiche_rs_ordine
	
	Private m_vecchio_conteggio
	Private m_modifica_articoli
	Private m_sposta_articoli
	Private m_manca_numeri_di_serie
	Private m_numeri_di_serie
	
	dim m_recordset
	Private m_esiste
	
	Private m_log_txt
	Private m_log
	Private m_chk_calcolato
	
	Private m_tipo_documento
	Private m_spettanze
	Private m_mostra_giacenze
	Private m_tipo_ddt
	Private m_foglioinordine
	
	dim rs_ordine
	dim oper
	Private ordine_Dictionary
	
    Public Default Function Init(parameters)
	    
	    m_log_txt="Init<br>"
	    m_log=false
	    m_chk_calcolato=false
	    dim come
	    Set Init = Me
		m_nuovo=false
	    n_parametri=UBound(parameters)+1
	    oper=lcase(parameters(0))
        m_tabella = parameters(1)
        
        if n_parametri>2 then
			m_idord=parameters(2)
		end if

	    select case oper
	    	case "apri"
			'set ordine= (new ClasseOrdine)(array("apri","ordini_fornitori",idord_fornitore))
	    		
	    		
				m_log_txt=m_log_txt&", CASE apri,"
	
				m_tipo_documento=get_tipo_documento(m_tabella)
			    if n_parametri=4 then
				    come=parameters(3)
				else
					come="completo"
				end if
				m_chk_calcolato=true
		        
				call ApriRecord(come)
				call ControlloOperazioni()

				
				
			case "nuovo"
			    m_log_txt=m_log_txt&", CASE nuovo,"
				m_tipo_documento=get_tipo_documento(m_tabella)
				call nuovo(parameters(2),parameters(3))
				
			case "nuovoddt"
			
				m_tipo_documento="ddt"
				m_tabella="ordini"
				m_iduser=parameters(1)
			    m_log_txt=m_log_txt&", CASE nuovoddt,"
			    'Elimino eventuali ordini tipo_documento=ddt senza record ddt per non compromettere numerazione
			    sql="SELECT ordini.idord FROM `ordini` left join ddt on ordini.idord = ddt.idord  where tipo_documento='ddt' and ddt.idddt is null"
			    set rs_tmp=conn.execute(sql)
			    
			    do while not rs_tmp.EOF
			    	call add2log("Elimino ordine tipo_documento ddt idord="&rs_tmp("idord")&" per evitare buco numerazione",3)
			    	conn.execute("delete from ordini where idord="&rs_tmp("idord"))
			    	rs_tmp.movenext
			    loop
				set rs_tmp = nothing
				    			    
				call crea_modifica_ordine("add")
				call crea_ddt()
				idddt=m_idord
				
				if request.form("spettanze")<>"" then
				
					call add2log("aggiorno spettanze",0)
					m_idord=request("idord")
					call aggiorna_quantita_spettanze(idddt)
					m_idord=idddt
					
					call calcolo_totale_merce_ddt(m_idord)
				end if

				
				
				
				
				
				
			case "aggiungi_spettanze"
				'set ordine= (new ClasseOrdine)(array("aggiungi_spettanze",idord,iduser))
			    m_log_txt=m_log_txt&", CASE aggiungi_spettanze,"
				m_tabella="ordini"
				m_idord=parameters(1)
				call ApriRecord("totali")
				call aggiungi_spettanze()
				
			case "copia_da"
				m_log=true
			    m_log_txt=m_log_txt&", CASE copia_da,"
				m_tabella=parameters(3)
				call copia_da(parameters(1),parameters(2))
			case "sposta_articoli"
			'set ordine= (new ClasseOrdine)(array("sposta_articoli",tabella,idord))
				m_log=true
				nuovo_ordine=duplica()
				call sposta_copia_articoli("S",nuovo_ordine)
			    'm_log_txt=m_log_txt&", CASE sposta_articoli,"
			case "copia_articoli"
			'set ordine= (new ClasseOrdine)(array("copia_articoli",tabella,idord))
				m_log=true
				nuovo_ordine=duplica()
				call sposta_copia_articoli("C",nuovo_ordine)
			    'm_log_txt=m_log_txt&", CASE copia_articoli,"
			    
			case "duplica"
			'set ordine= (new ClasseOrdine)(array("duplica",tabella,iduser,duplicaidord)))
				m_log=true
				
				m_nuovo=true
			    m_log_txt=m_log_txt&",nuovo"
				m_tabella=parameters(1)
				m_iduser=parameters(2)
				m_idord=parameters(3)
			    da_m_cosa_log=m_cosa_log(m_tabella,m_idord,conn.execute("select nord from "&m_tabella&" where idord="&m_idord)(0))
				
				call crea_modifica_ordine("add")
				nuovo_ordine=m_idord
				m_idord=parameters(3)
				call sposta_copia_articoli("D",nuovo_ordine)
			    'm_log_txt=m_log_txt&", CASE copia_articoli,"
    			call add2log ("Creato "&m_cosa_log(m_tabella,nuovo_ordine,m_nord)&" duplicato da "&da_m_cosa_log&vbcrlf&txt_modifiche&vbcrlf&"totale_merce_ordine:"&subtotale&vbcrlf&"Versione finale:"&versione&vbcrlf&"function CalcolaTotaleMerce",2)		

			case "crea_sostituzione"
			'set ordine= (new ClasseOrdine)(array("crea_sostituzione",idord))
				m_log=true
				
				m_nuovo=true
			    m_log_txt=m_log_txt&",nuovo"
				da_idord=parameters(1)
				m_tabella="ordini"
				m_tipo_documento="sostituzione"
				
				'Recupero iduser da ordine
				m_iduser=conn.execute("select iduser from ordini where idord="&da_idord)(0)
								
				call crea_modifica_ordine("add")
				sql="insert into ordini_sostituzioni  (idord,idsostituzione) values("&da_idord&","&m_idord&");"
				response.write sql
				conn.execute(sql)

				'call sposta_copia_articoli("D",nuovo_ordine)
			    'm_log_txt=m_log_txt&", CASE copia_articoli,"
    			call add2log ("Creata sostituzione "&m_cosa_log(m_tabella,nuovo_ordine,m_nord)&" duplicato da "&da_m_cosa_log&vbcrlf&txt_modifiche&vbcrlf&"totale_merce_ordine:"&subtotale&vbcrlf&"Versione finale:"&versione&vbcrlf&"function CalcolaTotaleMerce",2)		
	
				
				
	    	case "elimina"
			'set ordine= (new ClasseOrdine)(array("elimina",tabella,idord))
	
				call elimina_ordine()
	    	case "eliminatutto"
			'set ordine= (new ClasseOrdine)(array("elimina",tabella,idord))
	
				call elimina_ordine_tutto()
				
				
				
	    	case "aggiorna_quantita_spettanze"
			'set ordine= (new ClasseOrdine)(array("aggiorna_quantita_spettanze",tabella,idord))
				m_tabella=parameters(1)
				m_idord=parameters(2)
				idddt=0
				call add2log("aggiorna_quantita_spettanze: nparametri:"&n_parametri,0)
				if n_parametri>=4 then
					idddt=parameters(3)
				end if
				
				
				call aggiorna_quantita_spettanze(idddt)
			case "foglioinordine"
			    m_log_txt=m_log_txt&", CASE foglioinordine,"
				m_tipo_documento="ordine"
				m_tabella="ordini"
				m_foglioinordine=true
				
				
				m_iduser=clng(conn.execute("select iduser from utenti_foglio_spettanze where id="&parameters(2))(0))
				secondario=clng(conn.execute("select secondario from utenti where iduser="&m_iduser)(0))
				if secondario>0 then
					m_iduser=secondario
				end if
				
				
				call crea_modifica_ordine("add")
				call ApriRecord("completo")
				call foglioinordine(idfoglio)
			
		end select        

    End Function
    
    Private Sub Class_Terminate
	    if m_log then call add2log(m_log_txt,0)
		Set ordine_Dictionary = Nothing
	End Sub


    Private Function InitTwoParam(parameter1, parameter2)
        m_tabella = parameter1
        Set InitTwoParam = Me

    End Function
    
	
	Public Property Get esiste()
		esiste=m_esiste
	end property
	
	
	Public sub ApriRecord(come)
	    m_log_txt=m_log_txt&"<br>ApriRecord m_idord:"&m_idord&"<br>"
		m_spettanze=0
		dim rs_ordine
		dim n
		dim flds
		'response.write  "ApriRecord"
		Set ordine_Dictionary=Server.CreateObject("Scripting.Dictionary")
		if come="completo" then
			sql=replace("select tabella.*, utenti.trattamento_iva,utenti.email, utenti.telefono, utenti.fax, utenti.cellulare, utenti.tipologia,utenti.note_trasporto_2 FROM tabella INNER JOIN utenti ON tabella.iduser = utenti.iduser where idord="&m_idord,"tabella",m_tabella)
		elseif come="totali" then
			if m_tabella="ordini" then 
				campi_ordini=",tabella.spese_bancarie_ordine, tabella.arrotondamento, tabella.trattamento_iva_ordine, spettanze"
			elseif m_tabella="preventivi" then
				campi_ordini=", tabella.trattamento_iva_ordine, spettanze"
			else	'ordini_fornitori
				campi_ordini=", spettanze"
			end if

			sql=replace("select tabella.idord, tabella.iduser, tabella.nord, tabella.data, tabella.totale_merce_ordine, tabella.trasporto, tabella.calcolato,  tabella.iva_ordine, tabella.sconto_ordine,  tabella.totale, tabella.versione"&campi_ordini&" FROM tabella where idord="&m_idord,"tabella",m_tabella)
		end if
		'response.write sql
		'Set rs_ordine = Server.CreateObject("ADODB.Recordset")
		
		on error resume next
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
				ordine_Dictionary.add lcase(rs_ordine.Fields(i).Name),value
			next
			
			if DateDiff("d","01/01/2015",ordine_Dictionary.item("data")) <0 then 'prima del 01/01/2015
				m_vecchio_conteggio=true
			else
				m_vecchio_conteggio=false
			end if
			
			m_nord=rs_ordine("nord")
	        m_modifica_articoli=true
	        m_sposta_articoli=true
	        m_iduser=rs_ordine("iduser")
	        m_calcolato=rs_ordine("calcolato")
	        m_mostra_giacenze=true
	        if m_tabella="ordini" or tabella="preventivi" then
		        m_spettanze=rs_ordine("spettanze")
	        end if
	        if rs_ordine("calcolato")=0 and m_chk_calcolato=true then
		        call add2log(m_cosa_log(m_tabella,m_idord,m_nord)& " NON CALCOLATO",-1)
		        call CalcolaTotaleMerce(false)
		        'call CalcolaTotale()
		        call add2log(m_cosa_log(m_tabella,m_idord,m_nord)& " RICALCOLATO",-1)
		        
		    end if
		    if cint(rs_ordine("versione"))<=1 and m_tabella="ordini" then
			    'Valorizzo numeri_di_serie
				call conteggia_tutti_numeri_di_serie()	    
			end if
		    
		else
			m_log_txt=m_log_txt&" record non trovato"
			m_esiste=false
		end if
		set rs_ordine=Nothing
		m_log_txt=m_log_txt&" Fine ApriRecord m_iduser:"&m_iduser
    End sub
    
    Private sub ControlloOperazioni()
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
    Private sub elimina_ordine()
	    dim sql,rs
	    
		if m_tabella="ordini_fornitori" then
			sql="select ordini_dett_ordini_fornitori_dett.* from ordini_dett_ordini_fornitori_dett inner join ordini_fornitori_dett on ordini_dett_ordini_fornitori_dett.iddett_fornitori = ordini_fornitori_dett.iddett where idord="&idord
			'response.write sql
			set rs=conn.execute (sql)
			txt_log=""
			do while not rs.EOF
				'response.write "trovato "&rs("iddett_ordini")&" quantita:"&rs("quantita")
				txt_log=txt_log&"Impostato quantita_ordinato=0 per iddett:"&rs("iddett_ordini")&"<br>"
				conn.execute ("update ordini_dett set quantita_ordinato=0 where iddett="&rs("iddett_ordini"))
				conn.execute ("delete from ordini_dett_ordini_fornitori_dett where id="&rs("id"))
				rs.MoveNext
			loop
			
			
		end if
		if m_tabella="ordini" then
			sql="select ordini_dett_ordini_fornitori_dett.* from ordini_dett_ordini_fornitori_dett inner join ordini_dett on ordini_dett_ordini_fornitori_dett.iddett_ordini = ordini_dett.iddett where idord="&idord
			'response.write sql
			set rs=conn.execute (sql)
			txt_log=""
			do while not rs.EOF
				response.write "trovato "&rs("iddett_fornitori")&" quantita:"&rs("quantita")
				txt_log=txt_log&"Impostato quantita_ordinato=0 per iddett:"&rs("iddett_ordini")&"<br>"
				'conn.execute ("update ordini_dett set quantita=quantita where iddett="&rs("iddett_ordini"))
				conn.execute ("delete from ordini_dett_ordini_fornitori_dett where id="&rs("id"))
				rs.MoveNext
			loop			
		end if
		sql="select * FROM "&m_tabella&" WHERE idord="& m_idord&";"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		
		rs("comunicazioni_precedenti")=comunicazioni_precedenti(rs("comunicazioni_precedenti"),now()&" <b>"&session("nominativo")&"</b>: eliminato ordine")
		rs("eliminato")=1
		m_nord=rs("nord")

		rs.update
		txt_nominativo=nominativo(rs("nome"),rs("cognome"),rs("azienda"))
		rs.close
		set rs=Nothing
		
		
		
		call add2log ("Eliminato "&m_cosa_log(m_tabella,m_idord,m_nord)&" di "&txt_nominativo&" da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&txt_log, 2)
		call reset_tabella_ordini(m_tabella)
	end sub
    Private sub elimina_ordine_tutto()
	    dim sql,rs
	    
		sql="DELETE   FROM ordini_dett WHERE idord=" & m_idord
		conn.execute(sql)
		sql="DELETE   FROM incassi WHERE idord=" & m_idord
		conn.execute(sql)
		sql="DELETE   FROM ordini WHERE idord=" & m_idord
		conn.execute(sql)
		
		
		call add2log ("Eliminato "&m_cosa_log(m_tabella,m_idord,m_nord)&" di "&txt_nominativo&" da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&txt_log, 2)
		call reset_tabella_ordini(m_tabella)
	end sub
    
	Public function campo(nome)
		campo=ordine_Dictionary.Item(nome)
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
			ordine_Dictionary.Item("totale_merce_ordine")=ordine_Dictionary.Item("totale_merce_ordine")-subtotale
			
			n=CalcolaTotale()
			call concatena_stringa(txt_log,"<br>","Eliminato "&iddett&" riga <b>"&articolo_ordine&"</b> quantita: "&quantita)
			sql="UPDATE "&m_tabella&" SET comunicazioni_precedenti= CONCAT(comunicazioni_precedenti,'<hr>"& now()&" <b>"&session("nominativo")&":</b><br>"&replace(testo_log,"'","''")&"') WHERE idord = "&m_idord&";"
			conn.execute (sql)
			
			add2log "Modificato "&m_cosa_log(m_tabella,m_idord,m_nord)&vbcrlf&"Eliminato riga "&iddett&" <b>"&articolo_ordine&"</b> quantita: "&quantita&"totale_merce_ordine: "&ordine_Dictionary.item("totale_merce_ordine")&vbcrlf&"Versione finale:"&ordine_Dictionary.item("versione"),2
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
		conn.execute ("delete from "&m_tabella&"_dett_note where iddett="&iddett)
		conn.execute ("delete from "&m_tabella&"_dett_spettanze where iddett="&iddett)
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
			For Each Key in ordine_Dictionary
			 txt_tmp=Key & "=" & ordine_Dictionary.item(Key) & "<br>"
			 txt=txt&txt_tmp
			Next
			call add2log(txt,0)
		else
			response.write "<b>Record idord="&m_idord&" nella tabella "&m_tabella&" non trovato</b><br>"
		end if
	end sub
	
	
	Public function CalcolaTotaleMerce(da_form)
	    m_log_txt=m_log_txt&"<br>CalcolaTotaleMerce"
		on error goto 0
		dim subtotale
		dim imposta
		dim iva_ordine
		dim versione
		dim ordine_modificato
		dim txt_modifiche
		dim gravita_log
		ordine_modificato=false
		txt_modifiche=""
		gravita_log=2
		versione=ordine_Dictionary.item("versione")
		subtotale=0
		if da_form then
			if cint(request.form("versione"))<>versione then
				call add2log ("<b>Versione "&m_cosa_log(m_tabella,m_idord,m_nord)&" non corrispondente (versione recordset: "&versione&" versione form: "&request.form("versione")&"<br>Aggiornamento "& cosa&" non eseguito." ,3)
				session("report")="Mentre era aperto in questa pagina del browser questo "&cosa&" &egrave; stato modificato in un altra pagina del browser o da un altro utente. <br>L'aggiornamento dei dati non &egrave; stato eseguito ed &egrave; stata ricaricata la versione pi&ugrave; recente."
				if utente_andrea then 
					session("report")=session("report")&"<br><b>Versione "&m_cosa_log(m_tabella,m_idord,m_nord)&" non corrispondente (versione recordset: "&versione&" versione form: "&request.form("versione")&"<br>"
				end if
				exit function				
			end if
		end if

		sql = "select * FROM "&m_tabella&"_dett WHERE idord=" & m_idord
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
						'totale_merce_ordine=totale_merce_ordine+totale_riga_tmp
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
				if ordine_Dictionary.item("sconto_ordine")<>newval then
					sql_update=sql_update&", sconto_ordine="&aggiusta_decimale(newval,"sql")
					txt_modifiche=txt_modifiche&"<br>"&"Modificato sconto ordine da "&ordine_Dictionary.item("sconto_ordine")&" a "&newval

					ordine_Dictionary.item("sconto_ordine")=newval
					ordine_modificato=true
				END IF
			end if
			newval=aggiusta_decimale(TRIM( Request( "trasporto"  ) ),"asp")
			IF isNumeric( newval ) THEN
				newval=roundup(Cdbl(newval),2)
				if ordine_Dictionary.item("trasporto")<>newval then
					sql_update=sql_update&", trasporto="&aggiusta_decimale(newval,"sql")
					txt_modifiche=txt_modifiche&"<br>"&"Modificato trasporto da "&ordine_Dictionary.item("trasporto")&" a "&newval
					ordine_Dictionary.item("trasporto")=newval
					response.write "trsporto modificato"
					ordine_modificato=true
				END IF
			end if
		end if
		'patch per evitare che totale_merce_ordine sia vuoto
		ordine_Dictionary.Item("totale_merce_ordine")=ordine_Dictionary.Item("totale_merce_ordine")+0


		
		subtotale=roundup(subtotale,2) '''''''''''''Verificare da dove arrivano i decimali
		if ordine_Dictionary.Item("totale_merce_ordine")=subtotale then
			m_totale_merce_ok=true
		else
			ordine_modificato=true
			ordine_Dictionary.Item("totale_merce_ordine") = subtotale
			m_totale_merce_ok=false
		end if 
		
		
		if ordine_modificato then
			sql="UPDATE "&m_tabella&" SET totale_merce_ordine = "&aggiusta_decimale(subtotale,"sql")&sql_update&", versione=versione+1"
			if txt_modifiche<>"" then 
				sql=sql&", comunicazioni_precedenti= CONCAT(comunicazioni_precedenti,'<hr>"& now()&" <b>"&session("nominativo")&":</b> "&replace(txt_modifiche,"'","''")&"')"
			end if
			sql=sql&" WHERE idord = "&m_idord&";"
			ordine_Dictionary.item("versione")=ordine_Dictionary.item("versione")+1
			'response.write sql
			conn.execute(sql)
			'add2log "aggiorno da CalcolaTotaleMerce m_totale_merce_ok:"&m_totale_merce_ok&" txt_modifiche:"&txt_modifiche,0
			
		end if
		
		if ordine_modificato then
			add2log "Modificato "&m_cosa_log(m_tabella,m_idord,m_nord)&vbcrlf&txt_modifiche&vbcrlf&"totale_merce_ordine:"&subtotale&vbcrlf&"Versione finale:"&versione&vbcrlf&"function CalcolaTotaleMerce",gravita_log			
		end if

		CalcolaTotaleMerce=subtotale
		n=CalcolaTotale()
			
	end function

	Public Property Get totale_merce_ok()
		totale_merce_ok=m_totale_merce_ok
	end property
	
	
	Public function CalcolaTotale()
		on error goto 0
		dim log_txt
		log_txt=""
	    m_log_txt=m_log_txt&"<br>CalcolaTotale"
	    m_log_txt=m_log_txt&" totale_merce_ordine prima:"&ordine_Dictionary.item("totale_merce_ordine")
		dim subtotale
		dim imposta
		dim iva_ordine

		subtotale=ordine_Dictionary.item("totale_merce_ordine")+ordine_Dictionary.item("trasporto")+ordine_Dictionary.item("spese_bancarie_ordine")
		log_txt="totale_merce_ordine:"&ordine_Dictionary.item("totale_merce_ordine")&" trasporto:"&ordine_Dictionary.item("trasporto")&" spese_bancarie_ordine:"&ordine_Dictionary.item("spese_bancarie_ordine")&"<br>"
		
		if esenzione_iva(ordine_Dictionary.item("trattamento_iva_ordine")) then
			m_casoiva=1
			imposta_ordine=0
			if ordine_Dictionary.item("trattamento_iva_ordine")=10 then
				m_casoiva=2
				imposta_ordine=roundup(subtotale*ordine_Dictionary.item("iva_ordine")/100,2)
			end if
		else
			m_casoiva=3
			imposta_ordine=roundup(subtotale*ordine_Dictionary.item("iva_ordine")/100,2)
			subtotale=subtotale+imposta_ordine
		end if
		log_txt=log_txt&"trattamento_iva_ordine:"&ordine_Dictionary.item("trattamento_iva_ordine")&" imposta_ordine:"&imposta_ordine&" casoiva:"&m_casoiva&"<br>"
		subtotale=roundup(subtotale,2) '''''''''''''Verificare da dove arrivano i decimali
		
		if ordine_Dictionary.item("sconto_ordine")>0 then subtotale=subtotale-ordine_Dictionary.item("sconto_ordine")
		subtotale=subtotale+ordine_Dictionary.item("arrotondamento")
		
		if ordine_Dictionary.item("totale")=subtotale then
			m_totale_ok=true
		else
			
			m_totale_ok=false
		end if

		sql="UPDATE "&m_tabella&" SET imposta_ordine = "&aggiusta_decimale( imposta_ordine,"sql")&", totale = "&aggiusta_decimale(subtotale,"sql")&",totale_merce_ordine="&aggiusta_decimale(ordine_Dictionary.item("totale_merce_ordine"),"sql")&",  versione=versione+1, calcolato=1 WHERE idord = "&m_idord&";"
		'response.write sql
		
		conn.execute(sql)
		ordine_Dictionary.item("versione")=ordine_Dictionary.item("versione")+1
		ordine_Dictionary.item("totale")=subtotale
		ordine_Dictionary.item("imposta_ordine")=imposta_ordine
	    m_log_txt=m_log_txt&",totale aggiornato:"&ordine_Dictionary.item("totale")
		CalcolaTotale=subtotale
		log_txt=log_txt&"subtotale:"&subtotale
		
	    'call add2log(m_log_txt,0)
	    
		'Se c'è fattura, ricalcolo fattura
		if m_tabella="ordini" then
			set rs_fat=conn.execute("Select idfat from ordini_fatture where idord="&m_idord)
			if not rs_fat.eof then
				n=calcola_totale_fattura(rs_fat("idfat"))
			end if
			set rs_fat=nothing
		end if
		call add2log("CalcolaTotale "&m_tabella&"="&m_idord&" totale:"&subtotale&vbcrlf&log_txt,0)

	end function
'******************************************************************************************************************************************************************
'******************************************************************************************************************************************************************





	Public sub elenco_articoli()
		on error goto 0
		m_manca_numeri_di_serie=false
		m_numeri_di_serie=false
		
		'spettanze=0
		
	    m_log_txt=m_log_txt&",elenco_articoli"
	    if m_tabella="ordini" then
		    if ordine_Dictionary.item("stato")>3 then
				mostra_giacenze=false
			end if
		elseif m_tabella="preventivi" then
			m_modifica_articoli=true
		elseif m_tabella="ordini_fornitori" then
			'spettanze=ordine_Dictionary.item("spettanze")
		end if
		
		%>
		
		
		
		
		<tbody id="tab_articoli_body">
		<%'CICLO SU ARTICOLI
	        sql="select ordini_dett.*, magazzino.quantita_magazzino,magazzino.idmag, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, magazzino.db_ven, varianti_a.codicevara,varianti_a.variante_a, varianti_b.variante_b, ordini_dett_note.nota FRom (((magazzino RIGHT JoIN (ordini_dett left join ordini_dett_note on ordini_dett.iddett = ordini_dett_note.iddett) oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JoIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara ) left join prodotti on  magazzino.idpro = prodotti.idpro where idord="&m_idord &" order by ordine,iddett"
        
        
        sql=replace(sql,"ordini",m_tabella)
        if m_tipo_ddt="parziale" then
	        sql="select ordini_dett.*,ddt_dett_ordini.quantita, magazzino.quantita_magazzino,magazzino.idmag, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, magazzino.db_ven, varianti_a.codicevara,varianti_a.variante_a, varianti_b.variante_b, ordini_dett_note.nota FRom (((magazzino RIGHT JoIN (ordini_dett left join ordini_dett_note on ordini_dett.iddett = ordini_dett_note.iddett) oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JoIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara ) left join prodotti on  magazzino.idpro = prodotti.idpro inner join ddt_dett_ordini on ordini_dett.iddett = ddt_dett_ordini.iddett where ddt_dett_ordini.idddt="&m_idord &" order by ordine,iddett"
	        'response.write sql
	    end if
        
        
        
        
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
					set rs_ordinato=conn.execute("select ordini_fornitori.idord, ordini_fornitori.nord, ordini_fornitori.stato, ordini_dett_ordini_fornitori_dett.quantita from ordini_dett_ordini_fornitori_dett inner join ordini_fornitori_dett on ordini_dett_ordini_fornitori_dett.iddett_fornitori = ordini_fornitori_dett.iddett inner join ordini_fornitori on ordini_fornitori_dett.idord = ordini_fornitori.idord where iddett_ordini="&rs("iddett"))
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
			
				sql="select ordini.idord, ordini.nord, ordini.stato, ordini_dett_ordini_fornitori_dett.quantita from ordini_dett_ordini_fornitori_dett inner join ordini_dett on ordini_dett_ordini_fornitori_dett.iddett_ordini = ordini_dett.iddett inner join ordini on ordini_dett.idord = ordini.idord where iddett_fornitori="&rs("iddett")
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
	              <%if ordine_Dictionary.item("stato")<=3 and m_sposta_articoli then %>
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
              <b><a href="<%=txt%>" title="Clicca per le opzioni" class="menu_idpro" data-idmag="<%=RS("idmag")%>" data-idpro="<%=RS("idpro")%>"><%=codice_articolo_e_variante(  rs("codice_ordine"),rs("codicevara"))%></a></b>
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
					nota=rs("nota")
						if nota<>"" then
							response.write "<br>Nota:<span id=""notad"&RS("iddett")&""" class=""notaarticolo"">"&replace(nota,chr(10),"<br>")&"</span>"
						end if
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
		  
		  
		  <tr class="mieidettagli" >
		  <td colspan="6">Riga <%=rs("iddett")%>, <%if m_tabella<>"ordini_fornitori" then %>Prezzo di acquisto: <%=formatnumber(rs("costo_ordine"),2)%><%end if %><%if utente_andrea then %><span class="soloio"> idmag:<%=rs("idmag")%> idpro e var: <%=rs("idpro")%>|<%=rs("idvara")%>|<%=rs("idvarb")%><%if m_tabella="ordini" then %> db_ven:<%=rs("db_ven")%>numeri_di_serie:<%=rs("numeri_di_serie")%> quantita_scalato_magazzino: <%=rs("quantita_scalato_magazzino")%><br>tot_costo: <%=tot_costo%> tot_no_costo: <%=tot_no_costo%> tot_vendita: <%=tot_vendita%> in_ordine: <%=rs("in_ordine")%> pronto: <%=rs("pronto")%> consegnato: <%=rs("consegnato")%></span><%end if %>
		   <%end if %>
		  </td>
		  
		  </tr>
		  
		  
		 
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
				if tot_vendita=ordine_Dictionary.item("totale_merce_ordine") then
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
		on error goto 0

	    m_log_txt=m_log_txt&",elenco_totali"
		%>
	<tbody id="tab_articoli_totali">
		<%if utente_andrea then %>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				m_idord: &#8364; <%=m_idord%> ordine_Dictionary.item("imposta_ordine"):<%=ordine_Dictionary.item("imposta_ordine")%>
			</td>
		</tr>
		
		<%end if %>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Totale merce: &#8364; <%=FormatNumber(ordine_Dictionary.item("totale_merce_ordine"),2)%>
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
				Spese di trasporto (<%=tipo_trasporto(ordine_Dictionary.item("tipo_trasporto"))%>): &#8364;
				<%if not m_modifica_articoli then %>
					<%=FormatNumber(ordine_Dictionary.item("trasporto"),2)%>
				<%else %>
					<input type="text" name="trasporto" size="7" value="<%=FormatNumber(ordine_Dictionary.item("trasporto"),2)%>" style="text-align:right;" id="trasporto" class="i_text">
				<%end if%>
			</td>
		</tr>
		<%
		totale_imponibile=ordine_Dictionary.item("totale_merce_ordine")+ordine_Dictionary.item("trasporto")
		%>
		<%if ordine_Dictionary.item("spese_bancarie_ordine")>0 then
			totale_imponibile=totale_imponibile+ordine_Dictionary.item("spese_bancarie_ordine")
			%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Spese bancarie: &#8364; <span id="spese_bancarie"><%=FormatNumber(ordine_Dictionary.item("spese_bancarie_ordine"),2)%></span>
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
		testo_esenzione=" "&replace(trattamento_iva(ordine_Dictionary.item("trattamento_iva_ordine")),"Normale","")
		if esenzione_iva(ordine_Dictionary.item("trattamento_iva_ordine")) then
			caso_iva=1
			valore_iva=0
			if ordine_Dictionary.item("trattamento_iva_ordine")=10 then
				caso_iva=2
				imposta=ordine_Dictionary.item("iva_ordine")
				valore_iva=roundup(totale_imponibile*ordine_Dictionary.item("iva_ordine")/100,2)
				testo_esenzione=""
			end if
		else
			caso_iva=3
			imposta=ordine_Dictionary.item("iva_ordine")
			valore_iva=roundup(totale_imponibile*ordine_Dictionary.item("iva_ordine")/100,2)
			totale_imponibile=totale_imponibile+valore_iva
		end if
		%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Imposta IVA <span id="iva"><%=ordine_Dictionary.item("iva_ordine")%></span>%<%=testo_esenzione%>: &#8364; <span id="valore_iva"><%=FormatNumber(ordine_Dictionary.item("imposta_ordine"),2)%></span>
			</td>
		</tr>
		<%if ordine_Dictionary.item("trattamento_iva_ordine")=10 then%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				<%=trattamento_iva(ordine_Dictionary.item("trattamento_iva_ordine"))%>: &#8364; <span id="valore_iva">-<%=FormatNumber(ordine_Dictionary.item("imposta_ordine"),2)%></span>
			</td>
		</tr>
		<%end if
		if false then
		%>

		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">Sconto: &#8364;
				  <%if not m_modifica_articoli then %>
				  <%=FormatNumber(ordine_Dictionary.item("sconto_ordine"),2)%>
				  <%else %>
				  <input type="text" name="sconto_ordine" size="7" value="<%=FormatNumber(ordine_Dictionary.item("sconto_ordine"),2)%>" style="text-align:right;" id="sconto_ordine">
				  <%end if
				  %>
			</td>
		</tr>

		<%
		end if
		if ordine_Dictionary.item("arrotondamento")<>0 then%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				arrotondamento: &#8364; <%=FormatNumber(ordine_Dictionary.item("arrotondamento"),2)%>
			</td>
		</tr>
		<%end if%>
		<%if ordine_Dictionary.item("trattamento_iva_ordine")=10 then%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
				Totale con iva: &#8364; <%=FormatNumber(valore_iva+ordine_Dictionary.item("totale"),2)%>
			</td>
		</tr>
		<%end if%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><input type="hidden" name="versione" value="<%=ordine_Dictionary.item("versione")%>"><b>Totale: &#8364; <span id="totale"><%=FormatNumber(ordine_Dictionary.item("totale"),2)%></span></b></td>
		</tr>
		<%if utente_andrea  then%>
		<tr valign="middle" align="right" >
			<td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><b>Versione: <%=ordine_Dictionary.item("versione")%></b></td>
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
		
'******************************************************************************************************************************************************************
'******************************************************************************************************************************************************************
	Public sub crea_modifica_ordine(oper)
	    m_log_txt=m_log_txt&",crea_modifica_ordine"

		dim rs_ordine, modifiche_rs_ordine, che_form
		che_form=request.form("cheform")
		modifiche_rs_ordine=""
		Set rs_ordine = Server.CreateObject("ADODB.Recordset")
		if oper="update" then
			sql="select * FROM "&m_tabella&" where idord="&m_idord
			rs_ordine.Open sql, conn, 3, 3
			'memorizzo situazione iniziale
			n=rs_ordine.fields.count-1
			Set Flds = rs_ordine.Fields
			redim nome_campo(n)
			redim valore_campo(n)
			'response.write "n:"&n
			'memorizzo i nomi dei campi e i valori attuali
			for i=0 to n
				'response.write "i:"&i
				nome_campo(i)=rs_ordine.Fields(i).Name
				valore_campo(i)=converti_typevar14(rs_ordine(i))
			next
			
			if request.form("d_indirizzo")<>"" then
				idconsegna=aggiungi_consegna(iduser,request.form("d_azienda"),request.form("d_indirizzo"),request.form("d_citta"),request.form("d_cap"))
			else
				if request.form("consegna")<>"" then
					idconsegna=request.form("consegna")
				else
					idconsegna=0
				end if
			end if

			nome=AllFirstUp(trim(request.form("nome")))
			cognome=AllFirstUp(trim(request.form("cognome")))
			azienda=AllFirstUp(trim(request.form("azienda")))
			
			
			idintestazione= aggiungi_intestazione(iduser,azienda,cognome,nome,ucase(trim(request.form("cf"))),trim(request.form("piva")),trim(request.form("Indirizzo")),allfirstup(trim(request.form("citta"))),trim(request.form("cap")),trim(request.form("provincia")))
			rs_ordine("idintestazione")=idintestazione
			if request.form("data_documento")<>"" then
				rs_ordine("data")=request.form("data_documento")
			end if
			if request.form("seleziona_ordine")<>""  then
				conn.execute("insert into ordini_sostituzioni (idord, idsostituzione) values("&request.form("seleziona_ordine")&","&m_idord&")")
			end if
			if mod_larochelle then 
				rs_ordine("spettanze")=checkbox("spettanze")
			end if
			if ha_il_permesso("Z1") and request.form("idagente")<>"" then
				rs_ordine("idagente")=request.form("idagente")
			end if
			
			
			
			call add2log("spettanze:"&checkbox("spettanze"),0)

			
		elseif oper="add" then
			'Nuovo ordine
			
			'Creo intestazione
			'idintestazione=request.form("idintestazione")
		
			
			
			'Creo dati_consegna
			idconsegna=0
			if request.form("d_indirizzo")<>"" then
				idconsegna=aggiungi_consegna(iduser,request.form("d_azienda"),request.form("d_indirizzo"),request.form("d_citta"),request.form("d_cap"))
			end if
			if request.form("consegna")<>"" then
				idconsegna=request.form("consegna")
			end if
			
			
			
			
			
			
			anno=request.form("anno")
			if anno="" then
				if request.form("data_documento")<>"" then
					data=request.form("data_documento")
				else
					
					data=date()
				end if
				anno=year(data)
			else
				'Anno impostato quindi precedente
				'Cerco la data
				data=conn.execute("select max(data) from "&m_tabella&" where tipo_documento='"&m_tipo_documento&"' and anno="&anno)(0)
			end if

			m_nord=max_ord(m_tabella,anno)
		
			sql="select * FROM "&m_tabella
			rs_ordine.Open sql, conn, 3, 3
			
			Set rs_utenti = Server.CreateObject("ADODB.Recordset")
			rs_utenti.Open "select utenti.*, utenti_clienti.stato_estero FROM utenti left join utenti_clienti on utenti.iduser=utenti_clienti.iduser where utenti.iduser="&m_iduser, conn, 3, 3

			rs_ordine.addnew
			if m_tabella="ordini" then
				'response.write "m_tipo_documento:"&m_tipo_documento
				rs_ordine("tipo_documento")=m_tipo_documento
			end if
			if m_tipo_documento="ordine" then
				rs_ordine("primo")=get_primo_ordine(m_iduser)
			end if
			rs_ordine("data")=data
			rs_ordine("anno")=anno
			rs_ordine("nord")=m_nord
			if mod_larochelle then
				rs_ordine("stato")=1
				else
				rs_ordine("stato")=0
			end if
			rs_ordine("iva_ordine")=iva(date())
			rs_ordine("eliminato")=false
			rs_ordine("iduser")=m_iduser
			
			'rs_ordine("sconto_ordine")=0
			'rs_ordine("cognome")=rs_utenti("cognome")
			'rs_ordine("nome")=rs_utenti("nome")
			if che_form="admin" then
				rs_ordine("creato_da_admin")=sessionIDUser
				rs_ordine("comunicazioni_precedenti")= now()&" <b>"&session("nominativo")&":</b> Creato " & cosa
				
			else
				rs_ordine("comunicazioni_precedenti")= now()&" <b>Utente:</b> Creato " & cosa
			
			end if
			if rs_utenti("stato_estero")<>1 then
				rs_ordine("iva_0")=0
			else
				rs_ordine("iva_0")=1
			end if
			if m_tabella="ordini" or m_tabella="preventivi" then
				

				idagente=rs_utenti("idagente")
				if isnull(idagente) then idagente=0
				rs_ordine("idagente")=idagente
			end if
			rs_ordine("idintestazione")=rs_utenti("idintestazione")
			if mod_larochelle then 
				rs_ordine("spettanze")=determina_cliente_ha_spettanze()
			end if

		end if
		if request.form("piva_editabile")="NO" then
			rs_ordine("piva")=rs_utenti("piva")
			rs_ordine("cf")=rs_utenti("cf")
			rs_ordine("fattura_sp")=rs_utenti("fattura_sp")
		else
			rs_ordine("piva")=ucase(request.form("piva"))
			rs_ordine("cf")=ucase(request.form("cf"))
			
			if request.form("fattura_sp")="si" then
				rs_ordine("fattura_sp")=1
				else
				rs_ordine("fattura_sp")=0
				end if
		end if
		
		
		

		'Campi solo ordini
		if m_tabella="ordini" and che_form="admin" then

			if request.form("no_magazzino")="si" then
				rs_ordine("no_magazzino")=1
			else
				rs_ordine("no_magazzino")=0
			end if
			

		end if

		
		'campi ordini e preventivi
		if (m_tabella="ordini" or m_tabella="preventivi") and che_form="admin" then
			
			if request.form("trattamento_iva_ordine")<>"" then
				rs_ordine("trattamento_iva_ordine")=request.form("trattamento_iva_ordine")
			end if
			
			'Campi MEPA per ordini e preventivi
			rs_ordine("cig")=trim(ucase(request.form("cig")))
			rs_ordine("mepa_tipo")=cint(request.form("mepa_tipo"))
			rs_ordine("mepa_testo")=trim(request.form("mepa_testo"))
			if request.form("mepa_data")="" then
				rs_ordine("mepa_data")=NULL
			else
				rs_ordine("mepa_data")=request.form("mepa_data")
			end if
			rs_ordine("impegno_spesa")=trim(request.form("impegno_spesa"))
			if request.form("fattura_sp")="si" then
				rs_ordine("fattura_sp")=1
			else
				rs_ordine("fattura_sp")=0
			end if
		end if
		if m_tabella="preventivi" and che_form="admin" then
			if request.form("nascondi_totali")="" then
				rs_ordine("nascondi_totali")=0
			else
				rs_ordine("nascondi_totali")=1
			end if
		end if
		rs_ordine("idconsegna")=idconsegna
		rs_ordine("tipopagamento")=request.form("tipopagamento")
		rs_ordine("tipo_trasporto")=request.form("trasporto")
		rs_ordine("note_spedizione")=trim(request.form("note-trasporto"))
		rs_ordine("note")=request.form("note-su-ordine")
		if m_foglioinordine=true then
			rs_ordine("tipopagamento")=rs_utenti("pag_accordato")
			rs_ordine("trattamento_iva_ordine")=rs_utenti("trattamento_iva")
		end if

		
		
		
		
		rs_ordine.Update 

		if che_form="user" then
		'Potrebbe non servire memorizzare i dati nella tabella ordini
			rs_utenti("noteacq")=""
			rs_utenti("note")=""
			rs_utenti.update
		end if
		if oper="add" then
			rs_utenti.close
			set rs_utenti=nothing
			m_nominativo=nominativo(rs_ordine("nome"),rs_ordine("cognome"),rs_ordine("azienda"))
		end if

		if oper="update" then
			for i=0 to n
				val_rs_ordine=converti_typevar14(rs_ordine(i).value)
				if ((valore_campo(i)<>val_rs_ordine) or (isnull(valore_campo(i)) and not isnull(rs_ordine(i)))) then
						modifiche_rs_ordine=modifiche_rs_ordine&"modificato campo "&nome_campo(i)&" da "&valore_campo(i)&" a "&rs_ordine(i)&"<br>"
				end if
			next
			if (m_tabella="ordini" or m_tabella="preventivi") then
				ordine_Dictionary.item("trattamento_iva_ordine")=rs_ordine("trattamento_iva_ordine")
			end if
			rs_ordine("comunicazioni_precedenti")=rs_ordine("comunicazioni_precedenti")&"<hr>"& now()&" <b>"&session("nominativo")&":</b> Modificati dati intestazione:" & modifiche_rs_ordine
			
		end if
		
		rs_ordine.update
		if oper="add" then
			m_idord=Get_last_id(m_tabella)
			m_modifiche_rs_ordine=modifiche_rs_ordine

		end if
		rs_ordine.close
		
		if modifiche_rs_ordine="" then modifiche_rs_ordine="Nessuna modifica in "&cosa
		
		'Qui ricalcolo i totali
		if oper="update" then
			n=CalcolaTotale()
		end if
		'---------------------
		
		if request.form("salva_anagrafica")<>"" then
			sql="select * FROM utenti where iduser="&ordine_Dictionary.item("iduser")
			Set rs = Server.CreateObject("ADODB.Recordset")

			rs.Open sql, conn, 3, 3
			n=rs.fields.count-1
			redim nome_campo(n)
			redim valore_campo(n)
			for i=0 to n
				'response.write "i:"&i
				nome_campo(i)=RS.Fields(i).Name
				valore_campo(i)=rs(i)
			next
			rs("azienda")=request.form("azienda")
			rs("citta")=request.form("citta")
			rs("cAP")=request.form("cAP")
			rs("indirizzo")=request.form("indirizzo")
			rs("provincia")=ucase(request.form("provincia"))
			rs("regione")=ucase(request.form("regione"))
			rs("d_azienda")=request.form("d_azienda")
			rs("d_citta")=request.form("d_citta")
			rs("d_cAP")=request.form("d_cAP")
			rs("d_indirizzo")=request.form("d_indirizzo")
			rs("d_provincia")=ucase(request.form("d_provincia"))
			rs("tipo_trasporto_2")=request.form("tipo_trasporto_2")
			rs("note_trasporto_2")=request.form("note_trasporto_2")
			rs("tipologia")=replace(request.form("tipologia")," ","")
			if request.form("trattamento_iva_ordine")<>"" then
				rs("trattamento_iva")=request.form("trattamento_iva_ordine")
			end if
			rs.update
			for i=0 to n
				if valore_campo(i)<>rs(i) then
						'modifiche_rs2=modifiche_rs2&"modificato campo "&nome_campo(i)&"<br>"
						modifiche_rs2=modifiche_rs2&"modificato campo "&nome_campo(i)&" da "&valore_campo(i)&" a "&rs(i)&"<br>"
				end if
			next
			if modifiche_rs2="" then modifiche_rs2="<br>Nessuna modifica in anagrafica" else modifiche_rs2="<br>Modificato campi in anagrafica:<br>"&modifiche_rs2
			rs.close	
		end if
		set rs = nothing
		if oper="update" then
			add2log "Modifica dati intestazione "&m_cosa_log(m_tabella,m_idord,m_nord)&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&modifiche_rs_ordine&modifiche_rs2 &vbcrlf&"ClasseOrdine",2	
		end if	
	End sub
	
	
	
	
	
	Private function determina_cliente_ha_spettanze()
		dim n_dipendenti
		determina_cliente_ha_spettanze=0
		sql="select count(*) from utenti_dipendenti where iduser="&m_iduser
		n_dipendenti=clng(conn.execute(sql)(0))
		if n_dipendenti>0 then
			determina_cliente_ha_spettanze=1
		end if
	end function
	
	
	private sub TrasferisciCarrello(iduser_carrello)
			dim totale_merce_ordine, totcosto
			totale_merce_ordine=0
			totcosto=0
			sql="select * from "&m_tabella&"_dett"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql, conn, 3, 3
			if m_tabella="ordini" then
				Set rso = Server.CreateObject("ADODB.Recordset")
				rso.open "select * from ordini_dett_originale", conn, 3, 3
			end if
			ordine=0
			sql="select carrello.*, prodotti.codice, prodotti.articolo, prodotti.variante1, prodotti.variante2, prodotti.prezzo, prodotti.costo, prodotti.Promozione, prodotti.Sconto, prodotti.Prodata, prodotti.richiedi_seriale, varianti_a.variante_a, varianti_a.prezzo_ve_va, varianti_b.variante_b, prodotti.idfor, prodotti.fileimg, prodotti.aggiungi_a_ordine FROM varianti_b RIGHT JOIN (varianti_a RIGHT JOIN (carrello INNER JOIN prodotti ON carrello.idpro = prodotti.IDpro) ON varianti_a.IDvara = carrello.idvara) ON varianti_b.IDvarb = carrello.idvarb where iduser="&iduser_carrello
			if m_tabella="ordini_fornitori" then sql=sql& " and idfor="&m_iduser
			'response.write sql
			set rs_carrello=conn.execute(sql)
			do until rs_carrello.eof
				'tabella ordini_dett
				if converti_typevar14(rs_carrello("prezzo_ve_va"))>0 then prezzo=rs_carrello("prezzo_ve_va") else prezzo=rs_carrello("prezzo")
				prezzo=cdbl(prezzo)
				if m_tabella="ordini_fornitori" then
					prezzo=rs_carrello("costo")
				elseif rs_carrello("promozione")=1 and (rs_carrello("prodata")>date() or isnull(rs_carrello("prodata"))) then 
					prezzo=prezzo*(1-rs_carrello("sconto")/100)
				end if
				totale_riga_tmp=totale_riga(prezzo,0,rs_carrello("quantita"))
				totale_merce_ordine=totale_merce_ordine+totale_riga_tmp
				totcosto=totcosto+rs_carrello("quantita")*cdbl(rs_carrello("costo"))
				rs.addnew
				rs("idord")=m_idord
				rs("idpro")=rs_carrello("idpro")
				rs("var1")=rs_carrello("variante_a")
				rs("var2")=rs_carrello("variante_b")
				rs("um")=rs_carrello("um")
				rs("quantita")=rs_carrello("quantita")
				rs("prezzo")=prezzo
				rs("modificato")=false
				rs("sconto_prodotto")=0
				rs("codice_ordine")=rs_carrello("codice")
				rs("articolo_ordine")=rs_carrello("articolo")
				rs("variante1_ordine")=rs_carrello("variante1")
				rs("variante2_ordine")=rs_carrello("variante2")
				rs("idvara")=rs_carrello("idvara")
				rs("idvarb")=rs_carrello("idvarb")
				rs("varianti_ordine")=codifica_varianti_ordine(rs_carrello("variante_a"), rs_carrello("variante1"), rs_carrello("variante_b"), rs_carrello("variante2"))
				rs("ordine")=ordine
				if m_tabella="ordini" then
					'rs("numeri_di_serie")=conteggia_numeri_di_serie(rs_carrello("idpro"),rs_carrello("idvara"),rs_carrello("idvarb"),0)
				end if
				rs("totale_riga")=totale_riga_tmp
				if m_tabella<>"ordini_fornitori" then rs("costo_ordine")=rs_carrello("costo")
				rs.update
				if m_tabella="ordini" then
				'tabella ordini_dett_originale
					rso.addnew
					rso("idord")=m_idord
					rso("idpro")=rs_carrello("idpro")
					rso("var1")=rs_carrello("variante_a")
					rso("var2")=rs_carrello("variante_b")
					rso("um")=rs_carrello("um")
					rso("quantita")=rs_carrello("quantita")
					rso("prezzo")=prezzo
					rso("modificato")=false
					rso("codice_ordine")=rs_carrello("codice")
					rso("articolo_ordine")=rs_carrello("articolo")
					rso("variante1_ordine")=rs_carrello("variante1")
					rso("variante2_ordine")=rs_carrello("variante2")
					rso("idvara")=rs_carrello("idvara")
					rso("idvarb")=rs_carrello("idvarb")
					rso.update
				end if
				txt=""
				if rs_carrello("idvara")<>0 then
					txt =rs_carrello("variante1") & ": " & rs_carrello("variante_a")
				end if
				if rs_carrello("idvarb")<>0 then
					txt =txt & " " & rs_carrello("variante2") & ": " & rs_carrello("variante_b")
				end if
				
				'Sezione aggiungi a ordine
				aggiungi_a_ordine=rs_carrello("aggiungi_a_ordine")
				if aggiungi_a_ordine>0 and (m_tabella="ordini" or m_tabella="preventivi") then
					aggiunto=false
					sql="select aggiungi_a_ordine_dett.* from aggiungi_a_ordine_dett where idgruppo="&aggiungi_a_ordine
					set rs_aggiungi=conn.execute(sql)
					for n=1 to cint(rs_carrello("quantita"))
						do while not rs_aggiungi.eof
							rs.addnew
							rs("idord")=m_idord
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
							aggiunto=true
							ordine=ordine+incremento_ordine
						loop
						if aggiunto then 
							rs_aggiungi.movefirst
						end if
						n=n+1
					next
				end if
				'Fine sezione aggiungi a ordine
				rs_carrello.movenext
				ordine=ordine+incremento_ordine
			loop
			timerstr=timerstr&"--Trasferimento carrello: "&formatnumber(timer()-timertmp,3)&" - "&time()&"<br>"
			timertmp=timer()
	
			rs_carrello.close
			set rs_carello=nothing
			if m_tabella="ordini" then
				rso.close  'chiude tabella dett_ordini_originale
				set rso = Nothing 
			end if
			rs.close
			Set rs = Nothing
			if request.form("conserva_carrello")="" then
				conn.execute ("DELETE FROM carrello WHERE carrello.iduser=" & iduser_carrello)
				session("cache_carrellino")=""
			end if
			'totale_merce_ordine
			dati_trasporto=split(conn.execute("select impostazioni.dati_trasporto FROM impostazioni")(0),vbcrlf)

			tipotrasporto=request.form("Trasporto")
			if tipotrasporto<>"" and not isnull(tipotrasporto) then tipotrasporto=cint(tipotrasporto)
			if m_tabella="ordini_fornitori" then 
				'Trasporto zero perchè ordine a fornitore
				trasporto=0
			else
				trasporto=calcola_importo_trasporto(tipotrasporto,totale_merce_ordine)
				maggiorazione_gestione=converti_typevar14( metodo_pagamento(request.form("tipopagamento"),"maggiorazione_gestione"))
				if maggiorazione_gestione>0 then
					importo_maggiorazione_gestione=roundup(totale_merce_ordine*maggiorazione_gestione/100,2)
					trasporto=trasporto+importo_maggiorazione_gestione
				end if
				
			end if
			sql="update "&m_tabella&" set totale_merce_ordine="&aggiusta_decimale(totale_merce_ordine,"sql")&", trasporto="&aggiusta_decimale(trasporto,"sql")&" where idord="&m_idord
			'response.write sql
			conn.execute (sql)
			
			

			'Svuoto carrello

	end sub
	
	
	Private function crea_ddt()
		
		on error goto 0
		dim rs_ddt, txt_log
		m_tipo_ddt="completo"
		txt_log=""
		sub_idord=request("idord")
		'iduser=request.form("iduser")
		idconsegna=0
		idrollback=0
		if sub_idord="" then
			sub_idord=m_idord
			m_tipo_ddt="no_ordine"
		else
			'Memorizzospettanze
			idrollback= memorizza_rollback(sub_idord,"Situazione ordine "&sub_idord&" antecedente DDT "&m_nord)
		end if
		if request.form("spettanze")<>"" then
			parziale_txt="parziale "
			sub_idord=request.form("idord")
			m_tipo_ddt="parziale"
		end if	
		
		
		Set rs_ddt = Server.CreateObject("ADODB.Recordset")
		'aggiungi_ddt di ordini.asp
		'verifica che ddt non esiste già
		sql="select ordini.idord, ddt.IDddt,ddt.tipo_ddt, ddt.parziale FROM ddt INNER JOIN ordini ON ddt.idord = ordini.idord WHERE (((ordini.idord)="&idord&"));"
		set rs_ddt=conn.execute(SQL)
		if not rs_ddt.eof then
			if rs_ddt("tipo_ddt")="completo" then
				'esce se esiste già ddt completo
				crea_ddt="-1"
				exit function
			end if
		end if
		rs_ddt.close
		call add2log("creo ddt tipo_ddt:"&m_tipo_ddt&" sub_idord:"&sub_idord,0)
'		SQL="select Max(nord) AS MAX_IDDDT FROM ordini WHERE tipo_documento='ddt' and anno="&year(now())&";"
'		set rs_ddt=conn.execute(SQL)
'		MAX_IDDDT=rs_ddt("MAX_IDDDT")
'		if isnull(MAX_IDDDT) then MAX_IDDDT=0
'		MAX_IDDDT=MAX_IDDDT+1
'		rs_ddt.close
		sql="select * from ddt"
		rs_ddt.Open sql, conn, 3, 3
		data=date()
		rs_ddt.addnew
'		rs_ddt("NDDT")=MAX_IDDDT
		rs_ddt("tipo_ddt")=m_tipo_ddt
'		rs_ddt("data")=data
'		rs_ddt("anno")=year(data)
'		rs_ddt("iduser")=request.form("iduser")
		rs_ddt("idord")=m_idord
		rs_ddt("sub_idord")=sub_idord
		rs_ddt("causale")=request.form("causale")
		rs_ddt("porto")=request.form("porto")
		rs_ddt("imballo")=request.form("imballo")
		rs_ddt("colli")=request.form("colli")
		rs_ddt("peso")=request.form("peso")
		rs_ddt("dimensione")=request.form("dimensione")
		rs_ddt("annotazioni")=request.form("annotazioni")
		rs_ddt("data_inizio")=request.form("data_inizio")
		rs_ddt("incaricato")=request.form("incaricato")
		rs_ddt("vettore")=request.form("vettore")
		'rs_ddt("creata_da")=session("nominativo")
		rs_ddt("da_fatturare")=checkbox("da_fatturare")
		rs_ddt("idrollback")=idrollback
		
'		rs_ddt("eliminato")=0
		rs_ddt.update
		'rs_ddt.movefirs_ddtt
		idddt=get_last_id("ddt")
		crea_ddt=idddt&"|"&MAX_IDDDT
		rs_ddt.close
		set rs_ddt = Nothing
				
		if sub_idord>0 then
			sql="select * from ordini where idord="&sub_idord
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql, conn, 3, 3
			sub_nord=rs("nord")
			rs("peso")=request.form("peso")
			rs("dimensione")=request.form("dimensione")
			rs("colli")=request.form("colli")
			rs("porto")=request.form("porto")
			rs("pagamentobollettino")=request.form("pagamentobollettino")
			on error resume next
			rs("tipopagamento")=request.form("pagamento")
			if err.number<>0 then
				call add2log("Errore in crea_ddt(): pagamento"&request.form("pagamento"),0)
			end if
			Set ccausale_ddt = new cl_causale_ddt
			if checkbox(request.form("da_fatturare"))=1 then
				rs("verde")=0
			else
				rs("verde")=1
			end if
			rs.update
			rs.close
			Set rs = Nothing	
			ordine_txt="per [ordine="&sub_idord&"]"&sub_nord&"[/ordine]"
		end if
		if m_tipo_ddt="completo" then
			call consegna_tutto(sub_idord,true)
		end if
		set ccausale_ddt=Nothing
		txt_log=txt_log&queryeform()	
		call add2log ("Creato [ddt="&m_idord&"]"&m_nord&"/"&year(data)&"[/ddt] "&parziale_txt&ordine_txt&", eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&txt_log,2)
	
	end function
	
	public sub calcolo_totale_merce_ddt(idord)
		set rs_ddt=conn.execute("select * from ddt where idord="&idord)
		if rs_ddt("tipo_ddt")="parziale" then
		
			totale_merce=0
			set rs_dett=conn.execute("select ddt_dett_ordini.quantita, ordini_dett.prezzo, ordini_dett.sconto_prodotto from ddt_dett_ordini inner join ordini_dett on ddt_dett_ordini.iddett=ordini_dett.iddett where idddt="&idord)
			do while not rs_dett.eof 
				totale_merce= totale_merce+totale_riga(rs_dett("prezzo"),rs_dett("sconto_prodotto"),rs_dett("quantita"))
				rs_dett.MoveNext
			loop
		elseif rs_ddt("tipo_ddt")="no_ordine" then
			totale_merce=0
			set rs_dett=conn.execute("select ordini_dett.quantita, ordini_dett.prezzo, ordini_dett.sconto_prodotto from ordini_dett  where idord="&idord)
			do while not rs_dett.eof 
				totale_merce= totale_merce+totale_riga(rs_dett("prezzo"),rs_dett("sconto_prodotto"),rs_dett("quantita"))
				rs_dett.MoveNext
			loop
		
		elseif rs_ddt("tipo_ddt")="completo" then
		
			
		
		end if
		set rs_ddt = nothing
		conn.execute("update ordini set totale_merce_ordine="&aggiusta_decimale(totale_merce,"mysql")&" where idord="&idord)
		call add2log("calcolo_totale_merce_ddt idord:"&idord &" totalemerce:"&totale_merce,0)
		call ApriRecord("completo")
		call CalcolaTotale()
	end sub
	
	Public sub salva_dati_ddt()
		idddt=request.form("idddt")
		call add2log("Modifico idddt:"&idddt,0)
		Set rs_ddt = Server.CreateObject("ADODB.Recordset")
		sql="select * from ddt where idddt="&idddt
		rs_ddt.Open sql, conn, 3, 3
		rs_ddt("causale")=request.form("causale")
		rs_ddt("porto")=request.form("porto")
		rs_ddt("imballo")=request.form("imballo")
		rs_ddt("colli")=request.form("colli")
		rs_ddt("peso")=request.form("peso")
		rs_ddt("dimensione")=request.form("dimensione")
		rs_ddt("annotazioni")=request.form("annotazioni")
		rs_ddt("data_inizio")=request.form("data_inizio")
		rs_ddt("incaricato")=request.form("incaricato")
		rs_ddt("vettore")=request.form("vettore")
		rs_ddt("da_fatturare")=checkbox("da_fatturare")
		rs_ddt.update
		rs_ddt.Close
		set rs_ddt = Nothing
		call add2log("Modificati dati ddt"&vbcrlf&queryeform(),0)		
		
	end sub

	
	
	
	
	private sub nuovo(iduser_ordine,iduser_carrello)
		dim txt_record
		txt_record=""
		m_nuovo=true
	    m_log_txt=m_log_txt&",nuovo"
		m_iduser=iduser_ordine
		call crea_modifica_ordine("add")
	    m_log_txt=m_log_txt&"<br>fine crea_modifica_ordine m_idord:"&m_idord
		call TrasferisciCarrello(iduser_carrello)
	    m_log_txt=m_log_txt&"<br>fine TrasferisciCarrello m_idord:"&m_idord
		call ApriRecord("totali")
		n=CalcolaTotale()
		cosa=tabella_cosa(m_tabella)
		set rs=conn.execute ("select * from "&m_tabella&" where idord="&m_idord)
		n=rs.fields.count-1

		for i=0 to n
			txt_record=txt_record&rs.Fields(i).Name&"="&rs(i).value&"<br> "
		next
		Set rs = Nothing
		if utente_admin then
			add2log "Creato "&m_cosa_log(m_tabella,m_idord,m_nord)&" da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&queryeform()&"<b>Record ordine</b><br>"&txt_record,2
		else
			add2log "Creato "&m_cosa_log(m_tabella,m_idord,m_nord)&" da [utente="&iduser_ordine&"]"&m_nominativo&"[/utente] per importo "&ordine_Dictionary.item("totale")&vbcrlf&queryeform()&"<b>Record ordine</b><br>"&txt_record,2
		end if
		call reset_tabella_ordini(m_tabella)
		if not utente_admin then
			call email_ordine (m_idord,tabella_singolare(m_tabella), true)
		end if
	end sub
	
	
	private sub aggiungi_spettanze()
		
		'Creo_array dipendenti
		
		'loop su articoli spettanze
			'Se almeno una quantità dipendente aggiungo a array articoli_spettanze
			'Creo array idpro_idvara_idvarb_iddip con quantità tabella per dipendente
		
		
		
		'Apro recordset
		'Loop su articoli
			'confronto con  array articoli_spettanze
			
			'Se in recordset e in spettanze confronto dipendenti
			
			'Se solo in spettanze aggiungo a recordset
			
			'Se solo in recordset elimino da recordset
		
		

		
		call da_calcolare()
		
		
		
		dim txt_log
		txt_log=""
	    m_log_txt=m_log_txt&"<br>aggiungi_spettanze()"
				
	    
		sql="select dipendenti.iddip, cognome, nome from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.iduser="&m_iduser
		
	    response.write "<br>"&sql
        set rs_dipendenti=conn.execute (sql)
	    arrRS=rs_dipendenti.getrows()
	    ubound_arrRS=ubound(arrRS,2)
	    
	    m_log_txt=m_log_txt&", trovati "&ubound_arrRS&" dipendenti"
		'Loop su articoli
		sql="select prodotti_spettanze.id, prodotti.idpro, prodotti.codice, prodotti.articolo, prodotti.prezzo, prodotti.costo, prodotti.variante1,  prodotti.variante2, prodotti_spettanze.idvara, prodotti.tipo_taglia, varianti_a.variante_a, prodotti_spettanze.idvarb, varianti_b.variante_b from ((prodotti_spettanze INNER JOIN prodotti on prodotti_spettanze.idpro=prodotti.idpro) left join varianti_a on prodotti_spettanze.idvara = varianti_a.idvara ) left join varianti_b on prodotti_spettanze.idvarb = varianti_b.idvarb where iduser="&m_iduser&" order by prodotti_spettanze.ordine"
	
		set rs_prodotti=conn.execute(sql)
		do while not rs_prodotti.EOF
			m_log_txt=m_log_txt&"<br><b>Ciclo su articolo "&rs_prodotti("articolo")&" per "&ubound_arrRS&" dipendenti</b>"
			idpro=rs_prodotti("idpro")
			prezzo=request.form("prezzo_"&idpro)

response.write "prezzo:"&prezzo
			if prezzo="" then
					prezzo=rs_prodotti("prezzo")
			end if
			'Loop sui dipendenti per ogni articolo
	        for n = 0 to ubound_arrRS
				iddip=arrRS(0,n)
				stringa=idpro&"_"&rs_prodotti("idvara")&"_"&rs_prodotti("idvarb")&"_"&iddip
		        quantita=request.form("quantita_"&stringa)
			    m_log_txt=m_log_txt&", analizzo quantita_"&stringa&" trovato valore:"&quantita
				if quantita<>"" then
					m_log_txt=m_log_txt&", trovato quantita"
					'response.write "<br>"&sql
					taglia=recupera_taglia(rs_prodotti("tipo_taglia"),iddip)
					
					idvara=0
					idvarb=0

					if taglia=-2 then
						m_log_txt=m_log_txt&", non trovate misure per il dipendente "&arrRS(1,n)
						session("report")=session("report")&"<br>Taglia/Misura non trovata per il dipendente "&arrRS(1,n)								
					else
						testotaglia=""
						m_log_txt=m_log_txt&", trovato taglia:"&taglia
						variante1=""
						variante2=""
						var1=""
						var2=""
						variante1_ordine=""
						variante2_ordine=""
						'Cerco quale variante contiene le taglie
						variante1= lcase(rs_prodotti("variante1"))
						variante2= lcase(rs_prodotti("variante2"))
'						if variante1="taglia" or variante1="misura" then
'							'cerco in varianteA
'							set rs_var=conn.execute("select idvara, prezzo_ve_va from varianti_a where idpro="&idpro&" and variante_a='"&taglia&"'")
'							if rs_var.eof then
'								session("report")=session("report")&"<br>Taglia "&taglia&" non trovata per l'articolo "&rs_prodotti("codice")&" "&rs_prodotti("articolo")		
'								testotaglia=taglia
'							else
'								idvara=rs_var("idvara")
'								if cdbl(rs_var("prezzo_ve_va"))>0 then
'									prezzo=rs_var("prezzo_ve_va")
'								end if
'								var1=taglia
'								variante1_ordine= variante1
'							end if
'						elseif variante2="taglia" or variante2="misura" then
'							'Cerco in varianteb
'							
'							
'							set rs_var=conn.execute("select idvarb from varianti_b where idpro="&idpro&" and variante_b='"&taglia&"'")
'							if rs_var.eof then
'								session("report")=session("report")&"<br>Taglia "&taglia&" non trovata per l'articolo "&rs_prodotti("codice")&" "&rs_prodotti("articolo")		
'								testotaglia=taglia
'							else
'								idvarb=rs_var("idvarb")
'								var2=taglia
'								variante2_ordine= variante2
'							end if
'						else
'							session("report")=session("report")&"<br>Variante taglia o misura non trovata per l'articolo "&rs_prodotti("codice")&" "&rs_prodotti("articolo")
							testotaglia=taglia
'						end if
						if vartype(testotaglia)=2 then
							testotaglia=""
						end if
						
					end if
					
					m_log_txt=m_log_txt&"<br>Aggiungo a ordine: aggiungi_articolo(idpro:"&idpro&",idvara:"&idvar&",idvarb:"&idvarb&",0,quantita:"&quantita&",iddip:"&iddip&",0, txt_log)"
					
					totale_riga_tmp = aggiungi_articolo(idpro,idvara,idvarb,0,quantita,iddip,0,prezzo, txt_log)
				end if
	        Next	
	        'Scorta tecnica
			iddip=0
			idpro=rs_prodotti("idpro")
			stringa=idpro&"_"&rs_prodotti("idvara")&"_"&rs_prodotti("idvarb")&"_"&iddip
	        quantita=request.form("quantita_"&stringa)
		    m_log_txt=m_log_txt&", analizzo scorta tecnica quantita_"&stringa&" trovato valore:"&quantita
	        if isnumeric(quantita) then
		        if quantita>0 then
					totale_riga_tmp = aggiungi_articolo(idpro,idvara,idvarb,0,quantita,null,0, prezzo,txt_log)
				end if
		    end if
	        	            
		rs_prodotti.MoveNext
		Loop

	    
	    
	    conn.execute ("update ordini set spettanze=1 where idord="&m_idord)
	    
	    
	    
	    n=CalcolaTotaleMerce(false)
		'n=CalcolaTotale()
		cosa=tabella_cosa(m_tabella)
		add2log "Aggiunte spettanze su "&m_cosa_log(m_tabella,m_idord,m_nord)&" da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&queryeform()&"<b>Record ordine</b><br>"&txt_record,2
		
		
	    m_log_txt=m_log_txt&", Fine aggiungi_spettanze()"
	    'Response.write "Fine aggiungi_spettanze()<br>"&m_log_txt
		call add2log(m_log_txt,0)
	end sub	'aggiungi_spettanze()
	
	
	
	
	public function aggiungi(aggiungo)
		dim rs
		dim ordine
		dim totale_riga_tmp
		dim testo_log

		on error goto 0
	    m_log_txt=m_log_txt&",aggiungi"
			
			call da_calcolare() 'Imposto calcolato=0
			
			'Calcolo ordine
			ordine=cint(request("ordine"))
			if ordine=-1 then
				SQL="select Max("&m_tabella&"_dett.ordine) AS ordine FROM "&m_tabella&"_dett WHERE idord="&idord&";"
				set rs=conn.execute(SQL)
				ordine=RS("ordine")
				rs.close
				Set rs = Nothing
				if isnull(ordine) then ordine=0
				ordine=ordine+2
			else
				ordine=ordine+1
			end if
		    m_log_txt=m_log_txt&",totale_merce_ordine prima:"&ordine_Dictionary.item("totale_merce_ordine")
			
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
				
				if request.form("nota")<>"" then
					Set rsnota = Server.CreateObject("ADODB.Recordset")
					rsnota.open "select * from ordini_dett_note", conn,3,3
					rsnota.addnew
					rsnota("iddett")=iddett
					rsnota("nota")=request.form("nota")
					rsnota.update
					rsnota.Close
					set rsnota=Nothing
					
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
			
			ordine_Dictionary.item("totale_merce_ordine")=ordine_Dictionary.item("totale_merce_ordine")+totale_riga_tmp
		    m_log_txt=m_log_txt&",totale_merce_ordine dopo:"&ordine_Dictionary.item("totale_merce_ordine")
			n=CalcolaTotale()
			sql="UPDATE "&m_tabella&" SET comunicazioni_precedenti= CONCAT(comunicazioni_precedenti,'<hr>"& now()&" <b>"&session("nominativo")&":</b><br>"&replace(testo_log,"'","''")&"') WHERE idord = "&m_idord&";"
			conn.execute (sql)
			
			call add2log ("Modificato "&m_cosa_log(m_tabella,m_idord,m_nord)&vbcrlf&testo_log&vbcrlf&"totale_merce_ordine: "&ordine_Dictionary.item("totale_merce_ordine")&vbcrlf&"Versione finale:"&ordine_Dictionary.item("versione"),2)
	end function	'aggiungi(aggiungo)
		
		
		
	private function aggiungi_articolo(idpro,idvara,idvarb,sconto,quantita,iddip,ordine,prezzo, testo_log) 'Restituisce l'importo
				dim raggruppa
				raggruppa=false	'Flag raggruppamento articoli nell'ordine
				dim tipo_taglia
				tipo_taglia=null
				dim sql_prodotti
				sql_prodotti=""
				
				
				sqlString = "select iddett,prezzo FROM "&m_tabella&"_dett WHERE idord=" & idord& " and idpro=" & idpro  
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
					rs("idord")=idord
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
				call add2log("nota:"&queryeform(),0)
				if request.form("nota")<>"" then
					call aggiungi_nota(iddett,request.form("nota"))
				end if
				if request.form("nota_"&idpro)<>"" then
					call aggiungi_nota(iddett,request.form("nota_"&idpro))
				end if
				
				
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
		
		
		
	private sub aggiungi_nota(iddett,testo)
		Set rsnota = Server.CreateObject("ADODB.Recordset")
		rsnota.open "select * from ordini_dett_note where iddett="&iddett, conn,3,3
		if rsnota.eof then
			rsnota.addnew
			rsnota("iddett")=iddett
			rsnota("nota")=testo
		else
			rsnota("nota")=testo
		end if
		rsnota.update
		rsnota.Close
		set rsnota=Nothing
	end sub	
		
	private sub aggiungi_spettanza(iddett,iddip,idpro,tipo_taglia, quantita, byref testo_log)
		
				'Aggiungo riferimento dipendente (spettanze)
				if (m_tabella="ordini" or m_tabella="preventivi") and mod_dip then
					'if isnumeric(iddip) then
						'ordine_Dictionary.item("spettanze")=1
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
						
						if ordine_Dictionary.item("spettanze")=0 then
							ordine_Dictionary.item("spettanze")=1
						    conn.execute ("update ordini set spettanze=1 where idord="&m_idord)
					    end if
					    testo_log=testo_log&" per dipendente "&dipendente_log(iddip)
					'end if
				end if
	end sub
		
	private sub foglioinordine(idfoglio)
		dim txt_log, rs, rst
		txt_log=""
	    m_log_txt=m_log_txt&"<br>foglioinordine()"
				
		sql="select * from utenti_spettanze where idfoglio="&idfoglio
				
		set rs=	conn.execute(sql)
		
		do while not rs.EOF

		sql="select * from dipendenti where iddip="&rs("iddip")
		set rst=conn.execute(sql)
		if not rst.eof then
			totale_riga_tmp = aggiungi_articolo(rs("idpro"),0,0,0,rs("quantita"),rs("iddip"),0,null, txt_log)
		end if

			
		rs.movenext
		loop
		
		
		
		
		Set rs_foglio = Server.CreateObject("ADODB.Recordset")
		rs_foglio.open "select * from utenti_foglio_spettanze where id="&idfoglio, conn, 3,3
		rs_foglio("idord")=m_idord

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open "select * from ordini where idord="&m_idord,conn,3,3
		rs("spettanze")=1
		rs("comunicazioni_precedenti")=comunicazioni_precedenti(rs("comunicazioni_precedenti"),"Articoli importati da foglio spettanze data "&rs_foglio("creato")&" id:"&idfoglio)
		rs("note_cliente")=rs_foglio("note")
		rs("trasporto")=12
		rs("cig")=rs_foglio("cig")
		rs("mepa_testo")=rs_foglio("determinazione")
		if rs_foglio("tipo_ordine")=0 then
			rs("mepa_tipo")=3
		else
			rs("mepa_tipo")=2
		end if
		rs("mepa_data")=rs_foglio("completo")
		rs.update
		rs.Close
		set rs = nothing

		rs_foglio.update
		rs_foglio.Close
		
		
		set rs_foglio = Nothing
		
		ordine_Dictionary.item("trasporto")=12
	    n=CalcolaTotaleMerce(false)
		cosa=tabella_cosa(m_tabella)
		add2log "Aggiunte articoli su "&m_cosa_log(m_tabella,m_idord,m_nord)&" da foglio spettanze del "&dataFoglio&" [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&queryeform()&"<b>Record ordine</b><br>"&txt_record,2
	    m_log_txt=m_log_txt&", Fine foglioinordine()"
	    'Response.write "Fine aggiungi_spettanze()<br>"&m_log_txt
		call add2log(m_log_txt,0)

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
				on error resume next
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
				rs_a("iva_ordine")=iva(date())
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
				m_idord=Get_last_id(m_tabella)
				call add2log("step2 m_idord:"&m_idord,0)
				rs_a.close
				set rs_a=nothing
				
				
				call copia_articoli_spettanze(idord, tabella)
				
				'ApriRecord("totali")
		        'm_cosa_log=replace(tabella_singolare(m_tabella)," ","_")
		        'm_cosa_log="["&m_cosa_log&"="&m_idord&"]"&m_nord&"[/"&m_cosa_log&"]"

				add2log "Creato "&m_cosa_log(m_tabella,m_idord,m_nord)&" da "&m_cosa_log_da&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]",2
				call ApriRecord("totali")
				call CalcolaTotaleMerce(false)
				'CalcolaTotale()
	end sub
	
	public sub copia_articoli_spettanze(da_idord, da_tabella)
				Set rs_dett_pre = Server.CreateObject("ADODB.Recordset")
				Set rs_dett_ord = Server.CreateObject("ADODB.Recordset")
				Set rs_ord_nota = Server.CreateObject("ADODB.Recordset")
				sql1="select "&da_tabella&"_dett.* from ("&da_tabella&"_dett left join prodotti on "&da_tabella&"_dett.idpro= prodotti.idpro ) where idord="&da_idord&";"
				rs_dett_pre.Open sql1, conn, 3, 3
				
				
				
				sql2="select * from "&m_tabella&"_dett;"
				rs_dett_ord.Open sql2, conn, 3, 3
				'call add2log("sql rs_dett_pre:"&sql1&" sql  rs_dett_ord:"&sql2,0)
				n=rs_dett_pre.fields.count-1
				do while not rs_dett_pre.eof
				id_rs_dett_pre=rs_dett_pre("iddett")
				
				
				rs_dett_ord.addnew
				for i=1 to n
					on error resume next
					rs_dett_ord(rs_dett_pre.fields(i).name)=converti_typevar14(rs_dett_pre(rs_dett_pre.fields(i).name))
					if err.number<>0 then
						'response.write "errore nella copia del campo:"&	rs_dett_pre.fields(i).name
						call add2log("errore nella copia del campo:"&	rs_dett_pre.fields(i).name &" con valore "&rs_dett_pre(rs_dett_pre.fields(i).name)&"sql rs_dett_pre:"&sql1&" sql  rs_dett_ord:"&sql2,0)
						Err.Clear
					end if
				next
				on error goto 0
				rs_dett_ord("idord")=m_idord
				rs_dett_ord.update
				
				newid=get_last_id(m_tabella&"_dett")
				
				
				'Cerco le note
				set rs_nota=conn.execute("select * from preventivi_dett_note where iddett="&id_rs_dett_pre)
				if not rs_nota.eof then
					rs_ord_nota.Open "select * from ordini_dett_note", conn, 3, 3
					rs_ord_nota.addnew
					rs_ord_nota("iddett")=newid
					rs_ord_nota("nota")=rs_nota("nota")
					rs_ord_nota.update
					rs_ord_nota.close
					
				end if
				
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
	
	
	
	
	
	
	private sub sposta_copia_articoli(operazione,nuovo_ordine)
		dim v_elenco_articoli
		m_log_txt=m_log_txt&"sposta_copia_articoli(operazione:"&operazione&")"&" stringa_iddett:"&stringa_iddett&"<br>"

		'stringa_iddett=request.form("stringa_iddett")
		
		sql = "select * FROM "&m_tabella&"_dett WHERE iddett in(" & stringa_iddett &") and idord="&m_idord
		if operazione="D" then
			sql = "select * FROM "&m_tabella&"_dett WHERE idord="&m_idord
		end if
		
		sql_new = "select * FROM "&m_tabella&"_dett"
		Set rs = Server.CreateObject("ADODB.Recordset")
		
		set rs_new= Server.CreateObject("ADODB.Recordset")
		on error resume next
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
			sql="select * from "&m_tabella&" where idord="&m_idord
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
			m_log_txt=m_log_txt&">>Ricalcolo ordine sorgente idord:"&m_idord&"<br>"
			call ApriRecord("totali")
			call CalcolaTotaleMerce(false)
		end if
		'Calcolo ordine destinazione
		m_idord=nuovo_ordine
		m_log_txt=m_log_txt&">>Ricalcolo ordine duplicato idord:"&m_idord&"<br>"
		call ApriRecord("totali")
		call CalcolaTotaleMerce(false)
		
	end sub
	
	private function duplica()
		m_log_txt=m_log_txt&"duplica()<br>"
		Set rs_from = Server.CreateObject("ADODB.Recordset")
		Set rs_to = Server.CreateObject("ADODB.Recordset")
		sql="select * from "&m_tabella&" where idord="&m_idord
		on error resume next
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
		m_log_txt=m_log_txt&">>duplicato "&m_tabella&" "&m_idord&" su "&duplica&"<br>"
	end function

	private sub conteggia_tutti_numeri_di_serie ()
		dim rs_dett
		sql="select "&m_tabella&"_dett.* from "&m_tabella&"_dett where idpro<>'' and idpro<>0 and idord="&m_idord
		Set rs_dett = Server.CreateObject("ADODB.Recordset")
		rs_dett.Open sql, conn, 3, 3
		do while not rs_dett.EOF
			rs_dett("numeri_di_serie")=conteggia_numeri_di_serie(rs_dett("idpro"),rs_dett("idvara"),rs_dett("idvarb"),0)
			rs_dett.update
			rs_dett.MoveNext
		loop
		set rs_dett = Nothing
		conn.Execute("update ordini set versione=2 where idord="&m_idord)
		ordine_Dictionary.item("versione")=2
	end sub

	Public sub elenco_dipendenti
		dim rs_dipendenti, dipendenti, n_dipendenti
		dipendenti=""
		n_dipendenti=0
		if ordine.campo("spettanze")=1 then
			sql="select dipendenti.*, utenti_gradi.nome_grado, utenti_gradi.colore, utenti_gradi.colore_nominativo from (dipendenti inner join ( select distinct ordini_dett_spettanze.iddip from ordini_dett_spettanze inner join ordini_dett on ordini_dett_spettanze.iddett = ordini_dett.iddett where idord="&m_idord&") as t on dipendenti.iddip = t.iddip) left join utenti_gradi on dipendenti.grado = utenti_gradi.id order by sesso desc, cognome,nome"
			sql=replace(sql,"ordini", m_tabella)
			set rs_dipendenti=conn.execute (sql)
			do while not rs_dipendenti.EOF
				str_dipendente="<a href=""#"" id=""dipendente"&rs_dipendenti("iddip")&""" class=""dipendente""><b>"&dipendente_colorato(  rs_dipendenti("cognome")&" "&rs_dipendenti("nome"),rs_dipendenti("sesso"),rs_dipendenti("colore_nominativo") )&"</b></a> "&  grado_colorato(rs_dipendenti("nome_grado"),rs_dipendenti("colore"))
				call concatena_stringa(dipendenti,", ",str_dipendente )
				n_dipendenti=n_dipendenti+1
				rs_dipendenti.MoveNext
			loop
			n=clng(conn.execute("select count(*) from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.iduser="&m_iduser)(0))
			if n>0 then
				response.write "<table width=""100%"" border=""1"" cellpadding=""0"" cellspacing=""0"" bordercolor=""#E5E5E5"" class=""tabella1"" ><tbody><tr><td class=""ui-widget-header"">Dipendenti "&n_dipendenti&"/"&n&"<span class=""pull-right""><a href=""pag_adm_rep23.asp?iduser="&m_iduser&""">Elenco taglie</a></span></td></tr></tbody><tr><td>"&dipendenti&"</td></tr>"
				response.write "<tr><td></td></tr>"
				response.write "</table>"
			end if
			%>
			

			
			<%
		end if
		response.write "<span style=""float: right;""><input type=""button"" value=""Spettanze"" onClick=""location.href='pag_adm_spettanze.asp?idord="&m_idord&"'""><input type=""button"" value=""Spettanze per agente"" onClick=""location.href='pag_adm_spettanze_vert.asp?idord="&m_idord&"'""><input type=""button""  value=""Situazione spettanze"" onClick=""location.href='pag_adm_spettanze_situaz.asp?idord="&m_idord&"'""></span>"
	end sub
	Public sub dati_fatturazione_consegna()
		%>
          <tr>
            <td bgcolor="#E5E5E5">Dati fatturazione</td>
            <td width="50%" bgcolor="#E5E5E5">Dati consegna</td>
          </tr>
          <tr >
            <td width="50%" valign="top" style="border-bottom:1px solid;">
	            <%=dati_fatturazione(ordine_Dictionary("idintestazione"))%><br>
              Pagamento: <%=metodo_pagamento(ordine_Dictionary("tipopagamento"),"descrizione")%>
  				<%if ordine_Dictionary("idagente")>0 then %>
				<br >
				Agente: <%=conn.execute("select nominativo from agenti where idagente="&ordine_Dictionary("idagente"))(0)%>
				
				<%end if %>

              </td>
            <td width="50%" valign="top" style="border-bottom:1px solid;">
				<%=dati_consegna(ordine_Dictionary("idconsegna"))%>
                <br />Metodo spedizione: <%=tipo_trasporto(ordine_Dictionary("tipo_trasporto"))%>
              </td>
          </tr>
		<%
		if ordine_Dictionary("note_spedizione")<>"" then%>
		<tr class="ui-widget-header">	
			<td colspan="2">Note spedizione:</td>
		</tr>
		<tr>
			<td colspan="2"> <%=ordine_Dictionary("note_spedizione")%></td>
		</tr>
		<%
		end if
	end sub
	
	Public SUb dati_mepa()
		
		if ordine_Dictionary.item("cig")<>"" or ordine_Dictionary.item("mepa_tipo")>0 or ordine_Dictionary.item("mepa_data")<>"" then
			
        if ordine_Dictionary.item("cig")<>"" then
	        cig="Codice CIG: "&ordine_Dictionary.item("cig")
	    else
		    cig=""
		end if
        if ordine_Dictionary.item("mepa_tipo")>0 then
	        mepa_tipo_v=mepa_tipo( ordine_Dictionary.item("mepa_tipo"))&": "&ordine_Dictionary.item("mepa_testo")
	    else
		    mepa_tipo_v=""
		end if
        if ordine_Dictionary.item("mepa_data")<>"" then
	        mepa_data_v="Data ordine: "&ordine_Dictionary.item("mepa_data")
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

		if ordine_Dictionary.item("impegno_spesa")<>"" then%>
          <tr>
	          <td>Impegno di spesa: <%=ordine_Dictionary.item("impegno_spesa")%></td>
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
		if ordine_Dictionary.item("calcolato")=1 then
			conn.execute ("UPDATE "&m_tabella&" SET calcolato = 0 WHERE idord = "&m_idord&";")
			ordine_Dictionary.item("calcolato")=0
		end if
	end sub
	
	Private sub imposta_stato_articoli()	
		'Non più usata
		dim sql
		Set rs_dettaglio = Server.CreateObject("ADODB.Recordset")
		'Cuclo su spettanze ordine
		sql="select ordini_dett.iddett,id from ordini_dett inner join ordini_dett_spettanze on ordini_dett.iddett = ordini_dett_spettanze.iddett where idord="&m_idord
		set rs= conn.execute(sql)
		do while not rs.EOF
		quantita=request.form("quantitas_"&rs("id"))
			if quantita<>"" then
				
				rs_dettaglio.open "select * from ordini_dett_spettanze_dettaglio where iddett="&rs("iddett")&" and idspettanza="&rs("id")&" and stato="&request.form("stato_articolo"), conn, 3,3
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
				'sql="INSERT INTO ordini_dett_spettanze_dettaglio( iddett, idspettanza,stato, quantita, data) VALUES ("&rs("iddett")&","&rs("id")&","&request.form("stato_articolo")&","&quantita&",now())"
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
		
		call memorizza_rollback(m_idord,testo_rollback)
		
		
		call add2log("Inizio aggiorna quantità spettanze per idord:"&m_idord&" idddt:"&idddt,0)	
		iddett=0
		tot_in_consegna=0
		Set rs_dettaglio = Server.CreateObject("ADODB.Recordset")
		'Cuclo su spettanze ordine
		sql="select id,ordini_dett.iddett from ordini_dett inner join ordini_dett_spettanze on ordini_dett.iddett = ordini_dett_spettanze.iddett where idord="&m_idord
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
				quantita=cint(conn.execute("select quantita from ordini_dett_spettanze where id="&id)(0))
				
				sql="update ordini_dett_spettanze set in_ordine="&in_ordine&", pronto="&pronto
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
		call add2log("aggiorna_quantita_spettanze &idord"&m_idord,0)
	end sub
	Private sub aggiorna_totali_spettanze()
		sql="select * from ordini_dett where idord="&m_idord
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn,3,3
		do while not rs.EOF
			if cdbl(rs("quantita"))>0 then
				set conteggio=conn.execute ("select sum(in_ordine) as in_ordine, sum(pronto) as pronto, sum(consegnato) as consegnato from ordini_dett_spettanze where iddett="&rs("iddett"))
					
				rs("in_ordine")= zerosenull(conteggio("in_ordine"))

				rs("pronto")=zerosenull(conteggio("pronto"))
				rs("consegnato")=zerosenull(conteggio("consegnato"))
				rs.update
			end if
			rs.MoveNext
		loop
	end sub
	Private sub aggiungi_ddt_dett_ordini(idddt,iddett,quantita)
		dim sql
		if quantita>0 then
			
'			set rs_from=conn.execute("select ordini_dett.* from ordini_dett where iddett="&iddett )
'			Set rs_to = Server.CreateObject("ADODB.Recordset")
'			rs_to.open "ordini_dett",conn,3,3
'			rs_to.addnew
'			
'		    for i=1 to rs_to.Fields.Count-1
'		        'Response.Write(Rs.Fields(i).Name + "<br>")
'				rs_to(rs_to.Fields(i).Name)=rs_from(rs_to.Fields(i).Name)
'			next
'			rs_to("idord")=idddt
'			rs_to("quantita")=quantita
'			rs_to("totale_riga")=cdbl(rs_from("prezzo"))*quantita
'			rs_to.update			
			
			sql="insert into ddt_dett_ordini (idddt,iddett,quantita) values ("&idddt& ","&iddett&","&quantita&")"
			conn.execute (sql)
			
		end if
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
			var aggiungi_a_id=<%=m_idord%>;
		</script>
		<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1" id="tab_articoli">
		<thead id="tab_articoli_testa">
			<%if utente_andrea then %>
			<tr>
				<td colspan="6" >m_tabella:<%=m_tabella%> m_idord:<%=m_idord%> m_tipo_ddt:<%=m_tipo_ddt%> </td>
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
	Public sub riga_pulsanti(modifica_articoli,sposta,ordina_fornitore)
		%>
		
		<tr valign="middle" >
            <td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
			<%if ordina_fornitore then %>
				<input type="submit" value="Ordina a fornitore" id="ordina_fornitore" class="oper_articoli">
			<%end if %>
			<%if modifica_articoli then %><input type="submit" value="Sposta o copia" id="sposta_articoli" class="oper_articoli"><%end if %>
			<%if false then %>
			<input type="submit" value="Sostituzione" id="sostituzione" class="oper_articoli">
			<%end if %>
			<%if modifica_articoli then%><span style="float: right;"><input type="button" value="Aggiungi articolo" id="puls_aggiungi_articolo">
			<input type="submit" name="Aggiorna_ordine" id="Aggiorna_ordine" value="Aggiorna dati"></span>
			<%end if%>
            </td>
        </tr>

		
		
		<%
	
		
	end sub
		
	function comunicazioni_precedenti(comunicazione_precedente,testo)
		if isnull(comunicazione_precedente) or len(trim(comunicazione_precedente))=0  then
			comunicazioni_precedenti=comunicazione_precedente
		else
			comunicazioni_precedenti=comunicazione_precedente&"<hr />"& testo
		end if
	end function
		
		
		
	Private function get_primo_ordine(m_iduser)
		get_primo_ordine=clng(conn.execute("select count(*) from ordini where tipo_documento='ordine' and eliminato=0 and iduser="&m_iduser)(0))
		call add2log("Conteggio ordini:"&get_primo_ordine,0)
		if get_primo_ordine>0 then
			get_primo_ordine=0
		else
			get_primo_ordine=1
		end if
	end function

		
'******************************************************************************************************************************************************************
'******************************************************************************************************************************************************************
			
	Public Property Get totale_ok()
		totale_ok=m_totale_ok
	end property
	Public Property Get idord()
		idord=m_idord
	end property
	public property let idord(value)
         m_idord = value
    end property
   	Public Property Get nord()
   		
		nord=m_nord
	end property
	Public Property Get manca_numeri_di_serie()
		manca_numeri_di_serie=m_manca_numeri_di_serie
	end property
	Public Property Get numeri_di_serie()
		numeri_di_serie=m_numeri_di_serie
	end property
	

	Public Property Get modifica_articoli()
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		modifica_articoli=m_modifica_articoli
	end property
	Public Property Let modifica_articoli(value)
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		m_modifica_articoli=value
	end property
	Public Property Let mostra_spettanze(value)
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		m_spettanze=value
	end property
	Public Property Let mostra_giacenze(value)
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		
		m_mostra_giacenze=converti_bool( value)
	end property
	Public Property Let tipo_ddt(value)
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		m_tipo_ddt= value
	end property
	Public Property Let sposta_articoli(value)
		'Ritorna se si può ancora modificare l'elenco articoli (per visualizzare pulsanti aggiornamento)
		m_sposta_articoli=value
	end property
	
	
		
	'-----------------------------------------------------------------------------------------------------------------------------------------------
	'  												FINE CLASSE
	'-----------------------------------------------------------------------------------------------------------------------------------------------
End Class

function get_tipo_documento(m_tabella)
	select case m_tabella
		case "ordini"
			get_tipo_documento="ordine"
		case "preventivi"
			get_tipo_documento="preventivo"
		case "ordini_fornitori"
			get_tipo_documento="fornitore"
		case "fatture"
			get_tipo_documento="fattura"
			
		case "notacredito"
			get_tipo_documento="notacredito"
			m_tabella="ordini"
			
		case "ddt"
			get_tipo_documento="ddt"
			m_tabella="ordini"
		case "sostituzione"
			get_tipo_documento="sostituzione"
			m_tabella="ordini"
			
			
	end select
end function

function Calcolo_iva(totale_imponibile,imposta,trattamento_iva_ordine)
		dim return_array(3)
		
			'Calcolo IVA
		testo_esenzione=" "&replace(trattamento_iva(ordine_Dictionary.item("trattamento_iva_ordine")),"Normale","")
		if esenzione_iva(ordine_Dictionary.item("trattamento_iva_ordine")) then
			caso_iva=1
			valore_iva=0
			if ordine_Dictionary.item("trattamento_iva_ordine")=10 then
				caso_iva=2
				imposta=ordine_Dictionary.item("iva_ordine")
				valore_iva=roundup(totale_imponibile*ordine_Dictionary.item("iva_ordine")/100,2)
				testo_esenzione=""
			end if
		else
			caso_iva=3
			imposta=ordine_Dictionary.item("iva_ordine")
			valore_iva=roundup(totale_imponibile*ordine_Dictionary.item("iva_ordine")/100,2)
			totale_imponibile=totale_imponibile+valore_iva
		end if
		return_array(0)=imposta
		return_array(1)=valore_iva
		return_array(2)=totale_imponibile
		return_array(3)=testo_esenzione
		Calcolo_iva()=return_array
end function


function conteggia_numeri_di_serie(idpro,idvara,idvarb,seriali)
	dim  rs_prodotto
	dim sql
		'Conteggio il seriale del prodotto
		sql="select magazzino.db_ven, magazzino.idmag, prodotti.richiedi_seriale FROM magazzino INNER JOIN prodotti on magazzino.idpro = prodotti.idpro where magazzino.idpro="&idpro&" and magazzino.idvara="&idvara&" and magazzino.idvarb="&idvarb
		set rs_prodotto=conn.execute (sql)
		if not rs_prodotto.eof then
			seriali=seriali + cint(rs_prodotto("richiedi_seriale"))
			if rs_prodotto("db_ven")=1 then
				sql="select distinta_base.*, magazzino.idpro, magazzino.idvara,magazzino.idvarb, magazzino.db_ven, magazzino.idmag FROM magazzino inner JOIN distinta_base ON distinta_base.idmag_acq = magazzino.idmag where distinta_base.idmag_ven="&rs_prodotto("idmag")
				'response.write sql
				set rs_db = conn.execute ( sql)
				do while not rs_db.EOF
					'response.write "trovato idpro:"&rs_db("idpro_acq")&"<br>"
					seriali=  conteggia_numeri_di_serie(rs_db("idpro_acq"),rs_db("idvara_acq"),rs_db("idvarb_acq"),seriali)
					rs_db.MoveNext
				loop
			end if
		else
			seriali=seriali+0
		end if
	conteggia_numeri_di_serie=seriali
end function

function dettaglio_spettanze(byref m_tabella,byval iddett, quantita_tot, edit, byref stile_cella)
	dim txt, txt2, tot_pronto,tot_in_ordine, quantita_tot2, tot_consegnato, quantita
	'Cerco nelle spettanze
	txt=""
	tot_pronto=0
	tot_in_ordine=0
	tot_consegnato=0
	quantita_tot=cdbl(quantita_tot)
	quantita_tot2=quantita_tot
	sql="select "&m_tabella&"_dett_spettanze.* , dipendenti.cognome, dipendenti.nome, dipendenti.sesso, dipendenti.matricola, utenti_gradi.nome_grado, utenti_gradi.colore_nominativo from (("&m_tabella&"_dett_spettanze left join dipendenti on "&m_tabella&"_dett_spettanze.iddip=dipendenti.iddip) left join utenti_gradi on dipendenti.grado = utenti_gradi.id) where iddett="&iddett &" order by sesso desc, cognome, nome"
    'sql=replace(sql,"ordini",m_tabella)
	set rs_spettanze=conn.execute (sql)
	do while not rs_spettanze.eof 
		chiudi_span=false
		txt2=""
        quantita_tot=quantita_tot-cdbl(rs_spettanze("quantita"))
		span=""
		max=rs_spettanze("quantita")-rs_spettanze("consegnato")
		tot_in_ordine=tot_in_ordine+rs_spettanze("in_ordine")
		tot_pronto=tot_pronto+rs_spettanze("pronto")
		tot_consegnato=tot_consegnato+rs_spettanze("consegnato")
		quantita=rs_spettanze("quantita")
		if txt<>"" then txt=txt&"<br>"
		id=rs_spettanze("id")
		
		
		
		if edit then 'and rs_spettanze("consegnato")<quantita then
			
			
			in_ordine=rs_spettanze("in_ordine")
			pronto= rs_spettanze("pronto")
			consegna=""
			consegnato=rs_spettanze("consegnato")
			if rettifica then
				max=rs_spettanze("quantita")
			else
				if quantita-consegnato<pronto then
					pronto=0
				end if	
			end if
			
			'Togliere data-rule-max se non uso validator
			txt=txt&"<input type=""hidden"" id=""quantitam"&id&""" value="""&max&""">"
			if (rs_spettanze("pronto")+rs_spettanze("in_ordine"))<max or rettifica then
				txt=txt&"<input type=""text"" name=""in_ordine"&id&""" id=""in_ordine"&id&""" class=""in_ordine do"" autocomplete=""off"" data-rule-max="""&max&""" value="""&vuotosezero( in_ordine)&""">"
			end if
			if rs_spettanze("pronto")<max or rettifica then
				txt=txt&"<input type=""text"" name=""pronto"&id&""" id=""pronto"&id&""" class=""pronto do"" autocomplete=""off"" data-rule-max="""&max&""" value="""&vuotosezero( pronto)&""">"
			end if
			if rs_spettanze("consegnato")<quantita or rettifica then
				if rettifica then
					valore=vuotosezero(consegnato)
				else
					
					valore=""
				end if
				txt=txt&"<input type=""text"" name=""consegna"&id&""" id=""consegna"&id&""" class=""consegna do""  autocomplete=""off"" data-rule-max="""&max&""" value="""&valore&""">&nbsp;/"
			end if
		end if
		
		if rs_spettanze("in_ordine")>0 then
			if rs_spettanze("in_ordine")>=quantita then
				span="<span class=""in_ordines"">"
			else
				call concatena_stringa(txt2,",","<span class=""in_ordines"">"&rs_spettanze("in_ordine")&" in ordine</span>")
			end if
		end if
		if rs_spettanze("pronto")>0 then
			if rs_spettanze("pronto")>=quantita then
				span="<span class=""prontos"">"
			else
				call concatena_stringa(txt2,",","<span class=""prontos"">"&rs_spettanze("pronto")&" pronto</span>")
			end if
		end if
		if rs_spettanze("consegnato")>0 then
			if rs_spettanze("consegnato")>=quantita then
				span="<span class=""consegnatos"">"
			else
				call concatena_stringa(txt2,",","<span class=""consegnatos"">"&rs_spettanze("consegnato")&" consegnato</span>")
			end if
		end if
		if span<>"" then
			txt=txt&span
		end if
		txt=txt&"&nbsp;<span id=""quantitas"&id&""">"&rs_spettanze("quantita")&"</span>"
		if txt2<>"" then txt=txt&"("&txt2&")"
		if isnull(rs_spettanze("cognome")) then
			txt=txt&"x Scorta tecnica"
		else
			txt=txt&"x "
			if not edit then txt=txt&"<a href=""#"" id=""dipendente"&rs_spettanze("iddip")&""" class=""dipendente"">"
			txt=txt&dipendente_colorato(  rs_spettanze("cognome")&" "&rs_spettanze("nome"),rs_spettanze("sesso"),rs_spettanze("colore_nominativo") )
			
			if not edit then txt=txt&"</a>"'&" "&rs_spettanze("nome_grado")
			'if utente_andrea then txt=txt&" [soloio]max:"&max&" consegnato:"&rs_spettanze("consegnato")
		end if
		if rs_spettanze("taglia_misura")<>"" then
			txt=txt&" ("&rs_spettanze("taglia_misura")&")"
		end if
		
		if span<>"" then
			txt=txt&"</span>"
		end if
				if rs_spettanze("consegnato")>rs_spettanze("quantita") then
			
			txt=txt&"<span style=""background: red;"">Quantit&agrave; consegnata maggiore della quantit&agrave; ordinata "
			
			end if

		rs_spettanze.movenext
	loop
	set rs_spettanze = Nothing
	if quantita_tot>0 then
		call concatena_stringa(txt,"<br>","&nbsp;"&quantita_tot&"x scorta tecnica...")
	end if
	
	if txt<>"" then
		'Determino stile_cella
		if tot_consegnato>=quantita_tot2 then
			stile_cella="consegnatos"
		elseif tot_pronto>=quantita_tot2 then
			stile_cella="prontos"
		elseif tot_in_ordine>=quantita_tot2 then
			stile_cella="in_ordines"
		end if
	
		txt= "<br>"&txt
	end if	
	dettaglio_spettanze=txt
end function
sub consegna_tutto(sub_idord,true_false)
	if true_false then
		consegnato="quantita"
	else
		consegnato="0"
	end if
	if mod_dip then
		set rs=conn.execute ("select iddett from ordini_dett where idord="&sub_idord)
		do while not rs.EOF
			conn.execute("update ordini_dett_spettanze set consegnato="&consegnato&" where iddett="&rs("iddett"))
			rs.MoveNext
		loop
		
	end if
	conn.execute ("update ordini_dett set consegnato="&consegnato&" where idord="&sub_idord)
end sub
function memorizza_rollback(idord,testo)
	
	conn.execute("insert into rollback (idord,data,descrizione) values ("&idord&", now(),'"&testo&"')")
	idrollback=get_last_id("rollback")
	set rs=conn.execute("select *  from ordini_dett where idord="&idord)
	do while not rs.EOF
		sql="INSERT INTO rollback_ordini_dett_spettanze (idrollback, idspett, in_ordine, pronto, consegnato)  SELECT "&idrollback&", id, in_ordine, pronto, consegnato FROM `ordini_dett_spettanze` WHERE iddett = "&rs("iddett")
		conn.execute(sql)
		rs.MoveNext
	loop
	sql="INSERT INTO rollback_ordini_dett (idrollback, iddett, in_ordine, pronto, consegnato)  SELECT "&idrollback&", iddett, in_ordine, pronto, consegnato FROM `ordini_dett` WHERE idord = "&idord
	conn.execute (sql)
	call add2log("Memorizzo rollback per "&testo,2)
	memorizza_rollback=idrollback
end function
sub ripristina_rollback(idrollback)
	
	set rs_rollback=conn.execute("select * from rollback where id="&idrollback)
	idord=rs_rollback("idord")
	sql="UPDATE ordini_dett_spettanze a, rollback_ordini_dett_spettanze da SET a.in_ordine = da.in_ordine, a.pronto = da.pronto, a.consegnato = da.consegnato WHERE a.id = da.idspett AND da.idrollback = "&idrollback
	conn.execute sql,num1
	sql="UPDATE ordini_dett a, rollback_ordini_dett da SET a.in_ordine = da.in_ordine, a.pronto = da.pronto, a.consegnato = da.consegnato WHERE a.iddett = da.iddett AND da.idrollback = "&idrollback
	conn.execute sql,num2
	call add2log("Ripristino rollback  "&rs_rollback("descrizione")&" del "&rs_rollback("data")&", "&num2&" situazione articoli,"&num1&" situazione spettanze",2)
	conn.execute("update rollback set eseguito=1 where id="&idrollback)
	set rs_rollback = Nothing
	
end sub
	
	
function stato_articolo(stato)
	dim result
	'stato=cint(stato)
	select case stato
		case -1
		result=3
		case 1
		result="In ordine"
		case 2
		result="Pronto"
		case 3
		result="Consegnato"
		end select

	stato_articolo=result
	
end function
function vuotosezero(val)
	if 	val=0 then vuotosezero="" else vuotosezero=val
end function
sub imposta_data_stato(byref field)
	if field.value="" then field=now()
	
end sub


%>