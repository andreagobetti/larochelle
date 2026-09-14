<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Crea indirizzi consegna su ddt<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	
	'DDT
	Response.write "<b>DDT</b><br>"
	sql="select ddt.idddt, ddt.nddt,ordini.* from ddt inner join ordini on ddt.idord = ordini.idord "
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("d_indirizzo")<>"" then
			indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
			response.write "Trovato indirizzo:"&indirizzo&" per ddt"&rs("nddt")
			'cerco negli indirizzi
			sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&indirizzo&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ddt set idconsegna="&idindirizzo&" where idddt="&rs("idddt"))
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
		if rs("d_indirizzo")<>"" then
			indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
			response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			
			sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&replace(indirizzo,"'","''")&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ordini set idconsegna="&idindirizzo&" where idord="&rs("idord"))
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
		if rs("d_indirizzo")<>"" then
			indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
			response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			
			sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&replace(indirizzo,"'","''")&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update preventivi set idconsegna="&idindirizzo&" where idord="&rs("idord"))
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
		if rs("d_indirizzo")<>"" then
			indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
			response.write "Trovato indirizzo:"&indirizzo&" per ordine"&rs("nord")
			'cerco negli indirizzi
			
			sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&replace(indirizzo,"'","''")&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update ordini_fornitori set idconsegna="&idindirizzo&" where idord="&rs("idord"))
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
		if rs("d_indirizzo")<>"" then
			indirizzo=rs("d_azienda")&rs("d_indirizzo")&rs("d_citta")
			response.write "Trovato indirizzo:"&indirizzo&" per ddt"&rs("idfat")
			'cerco negli indirizzi
			sql="select id from utenti_consegna where concat(d_azienda,d_indirizzo,d_citta) = '"&indirizzo&"' and iduser="&rs("iduser")
			'response.write "<br>"&sql&"<br>"
			set rs2=conn.execute(sql)
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
					response.write "<b> indirizzo aggiunto id:"&idindirizzo&"</b>"
			else
				idindirizzo=rs2("id")
				response.write "<b> indirizzo esistente id:"&idindirizzo&"</b>"
			end if
			conn.execute("update fatture set idconsegna="&idindirizzo&" where idfat="&rs("idfat"))
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