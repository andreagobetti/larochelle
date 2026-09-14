<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Crea intestazioni<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	conn.execute("TRUNCATE TABLE `utenti_intestazioni` ")
	
	'UTENTI
	Response.write "<b>UTENTI</b><br>"
	sql="select utenti.* from utenti "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			response.write "Trovato indirizzo:"&indirizzo&" per iduser"&rs("iduser")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update utenti set idintestazione="&idindirizzo&" where iduser="&rs("iduser"))
			response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	
	
	
	'DDT
	Response.write "<b>DDT</b><br>"
	sql="select ddt.idddt, ddt.nddt,ordini.* from ddt inner join ordini on ddt.idord = ordini.idord "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			response.write "Trovato indirizzo:"&indirizzo&" per ddt"&rs("nddt")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ddt set idintestazione="&idindirizzo&" where idddt="&rs("idddt"))
			response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	
	'ORDINI
	Response.write "<b>ORDINI</b><br>"
	sql="select ordini.* from ordini "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ordini set idintestazione="&idindirizzo&" where idord="&rs("idord"))
			response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	'PREVENTIVI
	Response.write "<b>PREVENTIVI</b><br>"
	sql="select * from preventivi "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update preventivi set idintestazione="&idindirizzo&" where idord="&rs("idord"))
			response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	'ORDINI_FORNITORI
	Response.write "<b>ORDINI_FORNITORI</b><br>"
	sql="select * from ordini_fornitori "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ordini_fornitori set idintestazione="&idindirizzo&" where idord="&rs("idord"))
			response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	'FATTURE
	Response.write "<b>FATTURE</b><br>"
	sql="select fatture.idfat,ordini.* from fatture inner join ordini on fatture.idord = ordini.idord "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("indirizzo")<>"" then
			indirizzo=rs("azienda")&rs("indirizzo")&rs("citta")&rs("cognome")&rs("nome")&rs("cf")&rs("piva")
			response.write "Trovato indirizzo:"&indirizzo&" per ddt"&rs("idfat")
			'cerco negli indirizzi
			indirizzo=replace(indirizzo,"'","''")
			sql="select id from utenti_intestazioni where concat(azienda,indirizzo,citta,cognome,nome,cf,piva) = '"&indirizzo&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update fatture set idintestazione="&idindirizzo&" where idfat="&rs("idfat"))
			response.write "<br>"
			'response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	
	
	
	response.write "ESEGUITO"
end if
call connclose()

%>