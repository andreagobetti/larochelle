<%
sub modifica_ordini_fatture()
	'Modifico idord puntando al ddt
	
	set rs= conn.Execute("select ordini_fatture.id, ordini_fatture.idfat, ordini_fatture.idord, ddt.idord as idordddt from ordini_fatture inner join ddt on ordini_fatture.idord = ddt.sub_idord")
	do while not rs.EOF
		conn.execute("update ordini_fatture set idord="&rs("idordddt")&" where id="&rs("id"))
		conn.execute("update fatture set tipo_fattura='ddt' where idfat="&rs("idfat"))
		rs.MoveNext
	loop
	
end sub	
function duplica_apici(testo)
	if isnull(testo) then
		duplica_apici=""
	else
		duplica_apici=replace(testo,"'","''")
	end if

end function

sub imposta_tipo_fattura()
	response.write "imposta_tipo_fattura"
	sql="select ordini_fatture.*, ordini.tipo_documento FROM ordini_fatture inner join ordini on ordini_fatture.idord = ordini.idord where tipo_documento='ddt'"
	set rs= conn.execute (sql)
	do while not rs.eof 
		conn.execute "update fatture set tipo_fattura='ddt' where idfat="&rs("idfat"),num
		
		response.write "Imposto idfat:"&rs("idfat")&" a ddt "&num&"<br>"
		rs.MoveNext
	loop
	
end sub	
	
sub migra_ddt()
	set rs=conn.execute ("ddt")
	Set rs_ordini = Server.CreateObject("ADODB.Recordset")
	Set rs_to = Server.CreateObject("ADODB.Recordset")
	rs_to.open "select * from ordini_dett",conn,3,3
	rs_ordini.open "ordini",conn,3,3
	do while not rs.EOF
	
		sub_idord=rs("idord").value
		set rs2=conn.execute("select ordini.* from ordini where idord="&sub_idord)
		idintestazione=rs2("idintestazione")
		idconsegna=rs2("idconsegna")
		rs_ordini.addnew
		rs_ordini("tipo_documento")="ddt"
		rs_ordini("data")=rs("data")
		rs_ordini("anno")=year(rs("data"))
		rs_ordini("nord")=rs("nddt")
		rs_ordini("iduser")=rs("iduser")
		rs_ordini("idintestazione")=idintestazione
		rs_ordini("idconsegna")=idconsegna
		
		rs_ordini.update
		idord=Get_last_id("ordini")
		tipo="0"
		if rs("parziale")=1 then
			'Migro voci parziali
			tipo="1"
			set rs_from=conn.execute ("select ordini_dett.*,ddt_dett_ordini.quantita from ddt_dett_ordini inner join ordini_dett on ddt_dett_ordini.iddett = ordini_dett.iddett where idddt="&rs("idddt"))
			
			do while not rs_from.EOF
			
				rs_to.addnew
				
			    for i=1 to rs_to.Fields.Count-1
			        'Response.Write(Rs.Fields(i).Name + "<br>")
					rs_to(rs_to.Fields(i).Name)=rs_from(rs_to.Fields(i).Name)
				next
				set rstmp=conn.execute("select * from ddt_dett_ordini where idddt="&rs("idddt"))
				rs_to("idord")=idord
				quantita=clng(rs_from("quantita"))
				rs_to("quantita")=quantita
				rs_to("totale_riga")=cdbl(rs_from("prezzo"))*quantita
				rs_to.update			
				rs_from.MoveNext
			loop
		end if
		conn.execute("update ddt set idord="&idord &", sub_idord="&sub_idord&", tipo='"&tipo&"' where idddt="&rs("idddt"))
		rs.MoveNext
	loop
end sub
sub popola_utenti_consegna()
	response.write "<h1><b>popola_utenti_consegna</b></h1><br>"
	conn.execute("truncate table utenti_consegna")
	on error resume next
	'DDT
	if session("consegna_ddt")="" then
		response.write "<b>DDT</b><br>"
		sql="select ddt.idddt, ddt.nddt,ordini.* from ddt inner join ordini on ddt.idord = ordini.idord "
		set rs= conn.execute (sql)
		do while not rs.eof
			if rs("d_indirizzo")<>"" then
				indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
				indirizzo=replace(indirizzo,"'","''")
				'response.write "Trovato indirizzo:"&indirizzo&" per ddt"&rs("nddt")
				'cerco negli indirizzi
				sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&indirizzo&"' and iduser="&rs("iduser")
				''response.write "<br>"&sql&"<br>"
				set rs2=conn.execute(sql)
				if err.number<>0 then
					response.write "Errore:"&sql
					response.end
				end if
				if rs2.eof then
					'Indirizzo non trovato
						set rs_consegna=Server.CreateObject("ADODB.Recordset")
						rs_consegna.open "utenti_consegna",conn,3,3
						rs_consegna.addnew
						rs_consegna("iduser")=rs("iduser")
						rs_consegna("d_azienda")=rs("d_azienda")
						rs_consegna("d_indirizzo")=rs("d_indirizzo")
						rs_consegna("d_citta")=rs("d_citta")
						'rs_consegna("d_regione")=rs("d_regione")
						rs_consegna("d_provincia")=rs("d_provincia")
						rs_consegna("d_cap")=rs("d_cap")
						rs_consegna.update
						idindirizzo=Get_last_id("utenti_consegna")
						rs_consegna.close
						set rs_consegna = Nothing
						'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
				else
					idindirizzo=rs2("id")
					'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
				end if
				conn.execute("update ddt set idconsegna="&idindirizzo&" where idddt="&rs("idddt"))
				'response.write "<br>"
				'response.end
			end if
			rs.movenext
		loop
		rs.close
		set rs = nothing
		session("consegna_ddt")="ok"
	end if
	'ORDINI
	if session("consegna_ordini")="" then
		response.write "<b>ORDINI</b><br>"
		sql="select ordini.* from ordini "
		set rs= conn.execute (sql)
		do while not rs.eof
			if rs("d_indirizzo")<>"" then
				indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
				'response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
				'cerco negli indirizzi
				indirizzo=replace(indirizzo,"'","''")
				sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&replace(indirizzo,"'","''")&"' and iduser="&rs("iduser")
				''response.write "<br>"&sql&"<br>"
				set rs2=conn.execute(sql)
				if err.number<>0 then
					response.write "Errore:"&sql
					response.end
				end if
				if rs2.eof then
					'Indirizzo non trovato
						set rs_consegna=Server.CreateObject("ADODB.Recordset")
						rs_consegna.open "utenti_consegna",conn,3,3
						rs_consegna.addnew
						rs_consegna("iduser")=rs("iduser")
						rs_consegna("d_azienda")=rs("d_azienda")
						rs_consegna("d_indirizzo")=rs("d_indirizzo")
						rs_consegna("d_citta")=rs("d_citta")
						'rs_consegna("d_regione")=rs("d_regione")
						rs_consegna("d_provincia")=rs("d_provincia")
						rs_consegna("d_cap")=rs("d_cap")
						rs_consegna.update
						idindirizzo=Get_last_id("utenti_consegna")
						rs_consegna.close
						set rs_consegna = Nothing
						'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
				else
					idindirizzo=rs2("id")
					'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
				end if
				conn.execute("update ordini set idconsegna="&idindirizzo&" where idord="&rs("idord"))
				'response.write "<br>"
				'response.end
			end if
			rs.movenext
		loop
		rs.close
		set rs = nothing
		 session("consegna_ordini")="ok"
	end if
	'PREVENTIVI
	if session("consegna_preventivi")="" then
		response.write "<b>PREVENTIVI</b><br>"
		sql="select * from preventivi "
		set rs= conn.execute (sql)
		do while not rs.eof
			if rs("d_indirizzo")<>"" then
				indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
				'response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
				'cerco negli indirizzi
				indirizzo=replace(indirizzo,"'","''")
				sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&replace(indirizzo,"'","''")&"' and iduser="&rs("iduser")
				''response.write "<br>"&sql&"<br>"
				set rs2=conn.execute(sql)
				if err.number<>0 then
					response.write "Errore:"&sql
					response.end
				end if
				if rs2.eof then
					'Indirizzo non trovato
						set rs_consegna=Server.CreateObject("ADODB.Recordset")
						rs_consegna.open "utenti_consegna",conn,3,3
						rs_consegna.addnew
						rs_consegna("iduser")=rs("iduser")
						rs_consegna("d_azienda")=rs("d_azienda")
						rs_consegna("d_indirizzo")=rs("d_indirizzo")
						rs_consegna("d_citta")=rs("d_citta")
						'rs_consegna("d_regione")=rs("d_regione")
						rs_consegna("d_provincia")=rs("d_provincia")
						rs_consegna("d_cap")=rs("d_cap")
						rs_consegna.update
						idindirizzo=Get_last_id("utenti_consegna")
						rs_consegna.close
						set rs_consegna = Nothing
						'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
				else
					idindirizzo=rs2("id")
					'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
				end if
				conn.execute("update preventivi set idconsegna="&idindirizzo&" where idord="&rs("idord"))
				'response.write "<br>"
				'response.end
			end if
			rs.movenext
		loop
		rs.close
		set rs = nothing
		session("consegna_preventivi")="ok"
	end if
	'ORDINI_FORNITORI
	if session("consegna_fornitori")="" then
		response.write "<b>ORDINI_FORNITORI</b><br>"
		sql="select * from ordini_fornitori "
		set rs= conn.execute (sql)
		do while not rs.eof
			if rs("d_indirizzo")<>"" then
				indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
				'response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
				'cerco negli indirizzi
				indirizzo=replace(indirizzo,"'","''")
				sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&replace(indirizzo,"'","''")&"' and iduser="&rs("iduser")
				''response.write "<br>"&sql&"<br>"
				set rs2=conn.execute(sql)
				if err.number<>0 then
					response.write "Errore:"&sql
					response.end
				end if
				if rs2.eof then
					'Indirizzo non trovato
						set rs_consegna=Server.CreateObject("ADODB.Recordset")
						rs_consegna.open "utenti_consegna",conn,3,3
						rs_consegna.addnew
						rs_consegna("iduser")=rs("iduser")
						rs_consegna("d_azienda")=rs("d_azienda")
						rs_consegna("d_indirizzo")=rs("d_indirizzo")
						rs_consegna("d_citta")=rs("d_citta")
						'rs_consegna("d_regione")=rs("d_regione")
						rs_consegna("d_provincia")=rs("d_provincia")
						rs_consegna("d_cap")=rs("d_cap")
						rs_consegna.update
						idindirizzo=Get_last_id("utenti_consegna")
						rs_consegna.close
						set rs_consegna = Nothing
						'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
				else
					idindirizzo=rs2("id")
					'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
				end if
				conn.execute("update ordini_fornitori set idconsegna="&idindirizzo&" where idord="&rs("idord"))
				'response.write "<br>"
				'response.end
			end if
			rs.movenext
		loop
		rs.close
		set rs = nothing
		session("consegna_fornitori")="ok"
	end if
	'FATTURE
	if session("consegna_fatture")="" then
		response.write "<b>FATTURE</b><br>"
		sql="select fatture.idfat,ordini.* from fatture inner join ordini on fatture.idord = ordini.idord "
		set rs= conn.execute (sql)
		do while not rs.eof
			if rs("d_indirizzo")<>"" then
				indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
				indirizzo=replace(indirizzo,"'","''")
				'response.write "Trovato indirizzo:"&indirizzo&" per ddt"&rs("idfat")
				'cerco negli indirizzi
				sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&indirizzo&"' and iduser="&rs("iduser")
				''response.write "<br>"&sql&"<br>"
				set rs2=conn.execute(sql)
				if err.number<>0 then
					response.write "Errore:"&sql
					response.end
				end if
				if rs2.eof then
					'Indirizzo non trovato
						set rs_consegna=Server.CreateObject("ADODB.Recordset")
						rs_consegna.open "utenti_consegna",conn,3,3
						rs_consegna.addnew
						rs_consegna("iduser")=rs("iduser")
						rs_consegna("d_azienda")=rs("d_azienda")
						rs_consegna("d_indirizzo")=rs("d_indirizzo")
						rs_consegna("d_citta")=rs("d_citta")
						'rs_consegna("d_regione")=rs("d_regione")
						rs_consegna("d_provincia")=rs("d_provincia")
						rs_consegna("d_cap")=rs("d_cap")
						rs_consegna.update
						idindirizzo=Get_last_id("utenti_consegna")
						rs_consegna.close
						set rs_consegna = Nothing
						'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
				else
					idindirizzo=rs2("id")
					'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
				end if
				conn.execute("update fatture set idconsegna="&idindirizzo&" where idfat="&rs("idfat"))
				'response.write "<br>"
				'response.end
			end if
			rs.movenext
		loop
		rs.close
		set rs = nothing
		session("consegna_fatture")="ok"
	end if
	
	
	'response.write "ESEGUITO popola_utenti_consegna<br>"

	
	
end sub
sub popola_utenti_intestazioni()
	conn.execute("truncate table utenti_intestazioni")
	'UTENTI
	on error resume next
	response.write "<b>UTENTI</b><br>"
	sql="select utenti.* from utenti "
	set rs= conn.execute (sql)
	do while not rs.eof
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			'response.write "Trovato indirizzo:"&indirizzo&" per iduser"&rs("iduser")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			''response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
			if err.number<>0 then
				response.write "Errore:"&sql
				response.end
			end if
			if rs2.eof then
				'Indirizzo non trovato
					set rs_consegna=Server.CreateObject("ADODB.Recordset")
					rs_consegna.open "utenti_intestazioni",conn,3,3
					rs_consegna.addnew
					rs_consegna("iduser")=rs("iduser")
					rs_consegna("azienda")=rs("azienda")
					rs_consegna("cognome")=rs("cognome")
					rs_consegna("nome")=rs("nome")
					rs_consegna("cf")=rs("cf")
					rs_consegna("piva")=rs("piva")
					rs_consegna("indirizzo")=rs("indirizzo")
					rs_consegna("citta")=rs("citta")
					'rs_consegna("d_regione")=rs("d_regione")
					rs_consegna("provincia")=rs("provincia")
					rs_consegna("cap")=rs("cap")
					rs_consegna.update
					idindirizzo=Get_last_id("utenti_intestazioni")
					rs_consegna.close
					set rs_consegna = Nothing
					'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update utenti set idintestazione="&idindirizzo&" where iduser="&rs("iduser"))
			'response.write "<br>"
			'response.end
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	
	
	
	'DDT
	response.write "<b>DDT</b><br>"
	sql="select ddt.idddt, ddt.nddt,ordini.* from ddt inner join ordini on ddt.idord = ordini.idord "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			'response.write "Trovato indirizzo:"&indirizzo&" per ddt"&rs("nddt")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			''response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
			if err.number<>0 then
				response.write "Errore:"&sql
				response.end
			end if
			if rs2.eof then
				'Indirizzo non trovato
					set rs_consegna=Server.CreateObject("ADODB.Recordset")
					rs_consegna.open "utenti_intestazioni",conn,3,3
					rs_consegna.addnew
					rs_consegna("iduser")=rs("iduser")
					rs_consegna("azienda")=rs("azienda")
					rs_consegna("cognome")=rs("cognome")
					rs_consegna("nome")=rs("nome")
					rs_consegna("cf")=rs("cf")
					rs_consegna("piva")=rs("piva")
					rs_consegna("indirizzo")=rs("indirizzo")
					rs_consegna("citta")=rs("citta")
					'rs_consegna("d_regione")=rs("d_regione")
					rs_consegna("provincia")=rs("provincia")
					rs_consegna("cap")=rs("cap")
					rs_consegna.update
					idindirizzo=Get_last_id("utenti_intestazioni")
					rs_consegna.close
					set rs_consegna = Nothing
					'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ddt set idintestazione="&idindirizzo&" where idddt="&rs("idddt"))
			'response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	
	'ORDINI
	response.write "<b>ORDINI</b><br>"
	sql="select ordini.* from ordini "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			'response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			indirizzo=replace(indirizzo,"\","")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			''response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
			if err.number<>0 then
				response.write "Errore:"&sql
				response.end
			end if
			if rs2.eof then
				'Indirizzo non trovato
					set rs_consegna=Server.CreateObject("ADODB.Recordset")
					rs_consegna.open "utenti_intestazioni",conn,3,3
					rs_consegna.addnew
					rs_consegna("iduser")=rs("iduser")
					rs_consegna("azienda")=rs("azienda")
					rs_consegna("cognome")=rs("cognome")
					rs_consegna("nome")=rs("nome")
					rs_consegna("cf")=rs("cf")
					rs_consegna("piva")=rs("piva")
					rs_consegna("indirizzo")=rs("indirizzo")
					rs_consegna("citta")=rs("citta")
					'rs_consegna("d_regione")=rs("d_regione")
					rs_consegna("provincia")=rs("provincia")
					rs_consegna("cap")=rs("cap")
					rs_consegna.update
					idindirizzo=Get_last_id("utenti_intestazioni")
					rs_consegna.close
					set rs_consegna = Nothing
					'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ordini set idintestazione="&idindirizzo&" where idord="&rs("idord"))
			'response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	'PREVENTIVI
	response.write "<b>PREVENTIVI</b><br>"
	sql="select * from preventivi "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			'response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			''response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
			if err.number<>0 then
				response.write "Errore:"&sql
				response.end
			end if
			if rs2.eof then
				'Indirizzo non trovato
					set rs_consegna=Server.CreateObject("ADODB.Recordset")
					rs_consegna.open "utenti_intestazioni",conn,3,3
					rs_consegna.addnew
					rs_consegna("iduser")=rs("iduser")
					rs_consegna("azienda")=rs("azienda")
					rs_consegna("cognome")=rs("cognome")
					rs_consegna("nome")=rs("nome")
					rs_consegna("cf")=rs("cf")
					rs_consegna("piva")=rs("piva")
					rs_consegna("indirizzo")=rs("indirizzo")
					rs_consegna("citta")=rs("citta")
					'rs_consegna("d_regione")=rs("d_regione")
					rs_consegna("provincia")=rs("provincia")
					rs_consegna("cap")=rs("cap")
					rs_consegna.update
					idindirizzo=Get_last_id("utenti_intestazioni")
					rs_consegna.close
					set rs_consegna = Nothing
					'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update preventivi set idintestazione="&idindirizzo&" where idord="&rs("idord"))
			'response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	'ORDINI_FORNITORI
	response.write "<b>ORDINI_FORNITORI</b><br>"
	sql="select * from ordini_fornitori "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			'response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			''response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
			if err.number<>0 then
				response.write "Errore:"&sql
				response.end
			end if
			if rs2.eof then
				'Indirizzo non trovato
					set rs_consegna=Server.CreateObject("ADODB.Recordset")
					rs_consegna.open "utenti_intestazioni",conn,3,3
					rs_consegna.addnew
					rs_consegna("iduser")=rs("iduser")
					rs_consegna("azienda")=rs("azienda")
					rs_consegna("cognome")=rs("cognome")
					rs_consegna("nome")=rs("nome")
					rs_consegna("cf")=rs("cf")
					rs_consegna("piva")=rs("piva")
					rs_consegna("indirizzo")=rs("indirizzo")
					rs_consegna("citta")=rs("citta")
					'rs_consegna("d_regione")=rs("d_regione")
					rs_consegna("provincia")=rs("provincia")
					rs_consegna("cap")=rs("cap")
					rs_consegna.update
					idindirizzo=Get_last_id("utenti_intestazioni")
					rs_consegna.close
					set rs_consegna = Nothing
					'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ordini_fornitori set idintestazione="&idindirizzo&" where idord="&rs("idord"))
			'response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	'FATTURE
	response.write "<b>FATTURE</b><br>"
	sql="select fatture.idfat,ordini.* from fatture inner join ordini on fatture.idord = ordini.idord "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			'response.write "Trovato indirizzo:"&indirizzo&" per ddt"&rs("idfat")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			indirizzo=replace(indirizzo,"\","")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			''response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
			if err.number<>0 then
				response.write "Errore:"&sql
				response.end
			end if
			if rs2.eof then
				'Indirizzo non trovato
					set rs_consegna=Server.CreateObject("ADODB.Recordset")
					rs_consegna.open "utenti_intestazioni",conn,3,3
					rs_consegna.addnew
					rs_consegna("iduser")=rs("iduser")
					rs_consegna("azienda")=rs("azienda")
					rs_consegna("cognome")=rs("cognome")
					rs_consegna("nome")=rs("nome")
					rs_consegna("cf")=rs("cf")
					rs_consegna("piva")=rs("piva")
					rs_consegna("indirizzo")=rs("indirizzo")
					rs_consegna("citta")=rs("citta")
					'rs_consegna("d_regione")=rs("d_regione")
					rs_consegna("provincia")=rs("provincia")
					rs_consegna("cap")=rs("cap")
					rs_consegna.update
					idindirizzo=Get_last_id("utenti_intestazioni")
					rs_consegna.close
					set rs_consegna = Nothing
					'response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				'response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update fatture set idintestazione="&idindirizzo&" where idfat="&rs("idfat"))
			'response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	
	
	'response.write "ESEGUITO popola_utenti_intestazioni<br>"

	
	
end sub
sub migra_ddt2()
	conn.execute("truncate table ddt_dett_ordini")
	set rs_ddt=conn.execute ("select  ordini.idord, ordini.nord, ddt.sub_idord,ddt.tipo from ddt inner join ordini on ddt.idord = ordini.idord where tipo='1'")
	
	do while not rs_ddt.EOF
		response.write "DDT: "&rs_ddt("nord")&"<br>"
		sql="select b.iddett, a.quantita,a.idord from ordini_dett a inner join (select * from ordini_dett where idord="&rs_ddt("sub_idord")&") as b on a.idpro=b.idpro and a.idvara=b.idvara and a.idvarb=b.idvarb where a.idord="&rs_ddt("idord")
		response.write sql&"<br>"
		set rs=conn.execute(sql)
		do while not rs.EOF
			sql="INSERT INTO ddt_dett_ordini (idddt, iddett, quantita) VALUES("&rs_ddt("idord")&", "&rs("iddett")&", "&rs("quantita")&")"
			conn.execute (sql)
			rs.MoveNext
		loop

		rs_ddt.MoveNext
	loop
end sub





%>