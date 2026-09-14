<%
const F_IDMAG=0
const F_QUANTITA_MAGAZZINO=1
const F_DB_VEN=2
const F_DB_ACQ=3
const F_QUANTITA_MOVIMENTO=4
const F_IDMAG_VEN=5
const F_IDMAG_ACQ=6
const F_C_UITILIZZO=7

Class ClasseMagazzino
	Private c_idord
	Private c_idfor
	Private c_operazione
	Private c_txt_log
	Private idmag_no_indietro
	Private c_livello
	Private c_scan_ven
	Private c_scan_acq

    Private sub Class_Initialize
	    c_txt_log="Init<br>"
	    c_idord="NULL"
	    c_idfor="NULL"
	    c_livello=NULL
		c_scan_ven="|"
		c_scan_acq="|"
	End Sub



    Private Sub Class_Terminate
	    
	End Sub
	
	
	Public sub scarica_ordine(idord)
		
	    c_txt_log=c_txt_log&"scarica_ordine("&idord&")<br>"
		c_idord=idord
		c_operazione=2
		
		sql="select magazzino.idmag, magazzino.quantita_magazzino, magazzino.db_ven,magazzino.db_acq, ordini_dett.quantita as quantita_movimento FROM magazzino INNER JOIN ordini_dett ON (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro) where ordini_dett.idpro>0 and ordini_dett.idord="&c_idord
		'0=MAGAZZINO.IDMAG, 1=MAGAZZINO.QUANTITA_MAGAZZINO, 2=MAGAZZINO.DB_VEN, 3=MAGAZZINO.DB_ACQ, 4=QUANTITA_MOVIMENTO, 5=IDMAG_VEN, 6=IDMAG_ACQ, 7=CIEFF_UTILIZZO
		''on error resume next
		set rs_magazzino=conn.execute (sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		''on error goto 0
		'response.end
		rs_magazzino_arr=rs_magazzino.getrows()
		call scorri_magazzino(rs_magazzino_arr,2,0)
		response.write c_txt_log
		'call add2log(c_txt_log,0)
	end sub


    Private sub scorri_magazzino(rs_magazzino_in,operazione,quantita_massima)
	    dim rs_magazzino, n_cicli,n
	    rs_magazzino=rs_magazzino_in
		if isnull(c_livello) then
			c_livello=0
		else
			c_livello=c_livello+1
		end if
	    
	    
	    if operazione<>5 then
		    idmag_no_indietro=0
	    end if
	    c_txt_log=c_txt_log&"<b>scorri_magazzino(rs_magazzino, operazione:"&operazione&")</b> quantita_massima:"&isnull(quantita_massima)&"<br>"
	    
		'0=MAGAZZINO.IDMAG, 1=MAGAZZINO.QUANTITA_MAGAZZINO, 2=MAGAZZINO.DB_VEN, 3=MAGAZZINO.DB_ACQ, 4=QUANTITA_MOVIMENTO, 5=IDMAG_VEN, 6=IDMAG_ACQ, 7=CIEFF_UTILIZZO

	    
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
	    n_cicli=ubound(rs_magazzino,2)
	    
	    if n_cicli>=0 then
		    for n= LBound(rs_magazzino,2) to n_cicli
		    'do while not rs_magazzino.EOF
		    	quantita_magazzino=cdbl(rs_magazzino(F_QUANTITA_MAGAZZINO,n))
		    	quantita_movimento=cdbl(rs_magazzino(F_QUANTITA_MOVIMENTO,n))
			    txt_prima="----idmag:"&rs_magazzino(F_IDMAG,n)&"(quantita_magazzino:"&quantita_magazzino&", quantita_movimento:"&quantita_movimento&")"
	
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
						quantita_scarico=quantita_movimento*rs_magazzino(F_C_UITILIZZO,n)
						'quantita_movimento
					    txt_prima=txt_prima&", coefficiente:"&rs_magazzino(F_C_UITILIZZO,n)&")"
						
						if quantita_scarico>quantita_magazzino then
							quantita_scarico=quantita_magazzino
						end if
						
						'Inverto il segno per lo scarico
						'quantita_scarico=-quantita_scarico
						quantita_magazzino=quantita_magazzino-quantita_scarico
						
					
						quantita_tmp=quantita_magazzino/rs_magazzino(F_C_UITILIZZO,n)
						if quantita_massima>quantita_tmp or isnull(quantita_massima) then quantita_massima=quantita_tmp
					    if quantita_massima=0 then
						    quantita_movimento=0
						    quantita_magazzino=0
						end if
					    txt_prima=txt_prima&" quantita massima:"&quantita_massima&")<br>"
			   
		    		case else	
		    	End Select
			    c_txt_log=c_txt_log&txt_prima&" dopo (quantita_magazzino:"&quantita_magazzino&", quantita_movimento:"&quantita_movimento&")<br>"
				c_txt_log=c_txt_log&" INSTR "&c_scan_ven&","&"|"&rs_magazzino(F_IDMAG,n)&"|"&"<br>"
		    	if rs_magazzino(F_DB_VEN,n)=1 and instr(c_scan_ven,"|"&rs_magazzino(F_IDMAG,n)&"|")=0 then	'MAGAZZINO_DBVEN
			    	c_scan_ven=c_scan_ven&rs_magazzino(F_IDMAG,n)&"|"
				    c_txt_log=c_txt_log&">>>> c_scan_ven:"& c_scan_ven&")<br>"
					call scorri_db_prodotti_acquisto(rs_magazzino(F_IDMAG,n),quantita_movimento)
				else
					'Se non ha componenti scarico
					call aggiorna_magazzino(rs_magazzino(F_IDMAG,n),quantita_magazzino )
					'if operazione<>5 then
						call inserisci_movimento(rs_magazzino(F_IDMAG,n),quantita_movimento,operazione,idord,idfor)
					'end if
				end if

				''on error resume next
								    
				if rs_magazzino(F_db_ACQ,n)=1 and instr(c_scan_acq,"|"&rs_magazzino(F_IDMAG,n)&"|")=0 then
					c_scan_ven=c_scan_ven&rs_magazzino(F_IDMAG,n)&"|"
					c_scan_acq=c_scan_acq&rs_magazzino(F_IDMAG,n)&"|"
				    c_txt_log=c_txt_log&">>>> c_scan_acq:"& c_scan_acq&")<br>"
					quantita_massima=null
					call scorri_db_prodotti_vendita(rs_magazzino(F_IDMAG,n),quantita_magazzino)
				end if
		    Next
			    'rs_magazzino.movenext
		    'loop
		else
			'In rs_magazzino non ci sono record
		    c_txt_log=c_txt_log&"--In rs_magazzino non ci sono record<br>"
		end if
		set rs_magazzino = Nothing
    end sub
    
    
    Private sub scorri_db_prodotti_acquisto(idmag_ven,quantita)
	    c_txt_log=c_txt_log&"scorri_db_prodotti_acquisto(idmag_ven:"&idmag_ven&", quantita:"&quantita&")<br>"
	    c_txt_log=c_txt_log&">>>> imposto no_indietro:"& idmag_no_indietro&")<br>"
	    dim quantita_magazzino, quantita_massima, quantita_scarico, quantita_tmp
		quantita_massima=null
	    
	    'Scorro i componenti di acquisto
    	sql="select magazzino.idmag, magazzino.quantita_magazzino, magazzino.db_ven, magazzino.db_acq,"&quantita&", distinta_base.idmag_ven,distinta_base.idmag_acq, distinta_base.quantita FROM magazzino inner JOIN distinta_base ON distinta_base.idmag_acq = magazzino.idmag where distinta_base.idmag_ven="&idmag_ven
    	
		'0=MAGAZZINO.IDMAG, 1=MAGAZZINO.QUANTITA_MAGAZZINO, 2=MAGAZZINO.DB_VEN, 3=MAGAZZINO.DB_ACQ, 4=QUANTITA_MOVIMENTO, 5=IDMAG_VEN, 6=IDMAG_ACQ, 7=CIEFF_UTILIZZO
    	
		on error resume next
    	set rs_db = conn.execute ( sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
			response.end
		end if
		''on error goto 0
		operazione=5	'Scarico magazzino per ordine (DB)
    	call scorri_magazzino(rs_db.getrows(),operazione,quantita_massima)
    	set rs_db = nothing
	    c_txt_log=c_txt_log&"Aggiorno magazzino del prodotto venduto(idmag_ven:"&idmag_ven&", quantita_massima:"&quantita_massima&")<br>"
	
		'Aggiorno magazzino del prodotto venduto
		call aggiorna_magazzino(idmag_ven,quantita_massima)    
    End Sub
    
    
    Private sub scorri_db_prodotti_vendita(idmag_acq, quantita_magazzino)
	    dim rs_magazzino_arr
	    ''on error goto 0
	    c_txt_log=c_txt_log&"scorri_db_prodotti_vendita(idmag_acq:"&idmag_acq&", quantita_magazzino:"&quantita_magazzino&")<br>"
		dim quantita_tmp, quantita_massima
	    quantita_massima=NULL
	    'sql="select magazzino.idmag, distinta_base.*,  magazzino.quantita_magazzino,"&quantita_magazzino&" as quantita_movimento, magazzino.db_acq, magazzino.db_ven FROM magazzino INNER JOIN distinta_base ON distinta_base.idmag_ven =magazzino.idmag  where idmag_acq="&idmag_acq	    
    	sql="select magazzino.idmag, magazzino.quantita_magazzino, magazzino.db_ven, magazzino.db_acq, "&quantita_magazzino&",distinta_base.idmag_ven,distinta_base.idmag_acq, distinta_base.quantita FROM magazzino inner JOIN distinta_base ON distinta_base.idmag_ven = magazzino.idmag where idmag_acq="&idmag_acq	
    	
		'0=MAGAZZINO.IDMAG, 1=MAGAZZINO.QUANTITA_MAGAZZINO, 2=MAGAZZINO.DB_VEN, 3=MAGAZZINO.DB_ACQ, 4=QUANTITA_MOVIMENTO, 5=IDMAG_VEN, 6=IDMAG_ACQ, 7=CIEFF_UTILIZZO
	    'Scorro i componenti di vendita
		'on error resume next
    	set rs_db = conn.execute ( sql)
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
			response.end
		end if
		'on error goto 0
		'c_txt_log=c_txt_log&sql
		operazione=5	'Scarico magazzino per ordine (DB)
		rs_magazzino_arr =rs_db.getrows()
    	call scorri_magazzino(rs_magazzino_arr,operazione,quantita_massima)
	    c_txt_log=c_txt_log&"TROVATO QUANTITA MASSIMA:"&quantita_massima&")<br>"
    	
    	set rs_db = nothing
    
    End Sub
    
    
	    
    Private sub aggiorna_magazzino(idmag,quantita)
	    dim sql, num
	    c_txt_log=c_txt_log&"<b>aggiorna_magazzino</b>(idmag:"&idmag&", quantita:"&quantita&")<br>"
		sql="UPDATE magazzino SET magazzino.quantita_magazzino = "&quantita&",data_aggiornamento=now() WHERE (((magazzino.idmag)="&idmag&"));"
		'on error resume next
		conn.execute sql,num	
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
			response.end
		end if
		'on error goto 0
    end sub
    Private sub inserisci_movimento(idmag,quantita,operazione,idord,idfor)
	    c_txt_log=c_txt_log&"<b>inserisci_movimento</b>(idmag:"&idmag&", quantita:"&quantita&", operazione:"&operazione&", idord:"&idord&", idfor:"&idfor&")<br>"
	    dim sql, num
	    sql="INSERT INTO magazzino_movimenti (data,idmag,quantita,causale,iduser,idord,idfor) values (now(),"&idmag&","&metti_punto(quantita)&","&operazione&","&sessionIDUser&","&c_idord&","&c_idfor&");"
		''on error resume next
		conn.execute sql,num	
		if err.number<>0 then
			response.write c_txt_log&sql&"<br>"
		end if
		''on error goto 0
    end sub

End Class
	
	

	
	
	
	
	
	
	
	
	

	
	
	
	
%>