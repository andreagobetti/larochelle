<!--#include virtual="/md5-2.asp"-->
<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/classeModificheRS.asp" -->
<%
if session("idadmin") = "" then call login()

if session("idadmin")="" then
	add2log "Richiesta senza privilegio admin"& vbcrlf&queryeform()&"<br>"&Request.ServerVariables("HTTP_REFERER"),3
	call connclose()
	response.end
end if
'sessiontest(str_add2log)
Dim Js
oper=request("oper")
if oper<>"" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	select case oper

	case "aggiorna_incassi"
		idfat=request("idfat")
		if idfat="" then idfat=0
		idord=request("idord")
		if idord="" then idord=0
		call tabella_incassi(idord,idfat,request.form("totale"),false,false)
		call connclose()
		response.end

	case "aggiorna_scadenze"
		idfat=request("idfat")
		if idfat="" then idfat=0
		idord=request("idord")
		if idord="" then idord=0
		call tabella_scadenze(idfat,idord)
		call connclose()
		response.end

	case "aggiorna_scadenze_for"
		idfat=request("idfat")
		call tabella_scadenze_for(idfat)
		call connclose()
		response.end


	case "salva_nota_campo"
		nota=request.form("nota")
		campo=request.form("campo")


		id=request.form("id")

			rs.open "select * from dipendenti_misure where id="&id, conn,3,3
			rs("note_"&campo)=nota
			iddip=rs("iddip")
			rs.update
			rs.close

		txt_log=aggiorna_misure_in_ordini(iddip)


		Set Js = jsObject()

		js("success")=true
		js.Flush
		set js=Nothing
		call connclose()
		response.end

	case "salva_nota"
		nota=request.form("nota")
		call add2log(queryeform(),0)
		iddett=request.form("iddett")
		tabella=request.form("tabella")
		Set Js = jsObject()
		if nota="" then
			'elimino nota
			conn.execute("delete from ordini_dett_note where iddett="&iddett)
			js("message")="Nota "&iddett&" eliminata"
		else

			rs.open "select * from "&tabella&"_dett_note where iddett="&iddett, conn,3,3
			if rs.eof then
				rs.addnew
				rs("iddett")=iddett
			end if
			rs("nota")=nota
			rs.update
			rs.Close
			js("message")="Nota "&iddett&" salvata"
		end if
		js("success")=true
		js.Flush
		set js=Nothing
		call connclose()
		response.end
	case "aggiungi_dipendente_esistente"

		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		iddip=request.form("iddip")
		iduser=request.form("iduser")

		esiste=clng(conn.execute("select count(*) from utenti_dipendenti where iduser="&iduser&" and iddip="&iddip)(0))
		if esiste=0 then
				set rs_dip=conn.execute ("select * from dipendenti where iddip="&iddip)
				nominativo_dip=rs_dip("cognome")&" "&rs_dip("nome")
				set rs_dip = Nothing
				sql="insert into utenti_dipendenti (iddip, iduser) values ("&iddip&","&iduser&")"
				conn.execute (sql)
				call add2log("Aggiunto dipendente esistente "&nominativo_dip&" da "&utente_log(iduser),2)
			js("success")=true
			js("message")="Dipendente aggiunto"
		else
			js("success")=false
			js("message")="Dipendente esistente"

		end if
		js.Flush
		set js=Nothing
		call connclose()
		response.end
		case "salva_dipendente"


			action=request.form("action")
			iddip=request.form("iddip")
			iduser=request.form("iduser")


			iduser=request("iduser")
			sql="select dipendenti.* from dipendenti"
			if action="update" then
				sql=sql&" where iddip="&iddip
				txt_oper="Modificato "
			end if
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql, conn, 3, 3
			if action="add" then
				txt_oper="Aggiunto "
				rs.addnew
				'rs("iduser")=iduser
				'setta spettanze in ordini
				setta_spettanze_in_ordini(iduser)
			end if
			if action="update" then
				set modificheRS= (new ClasseModificheRS)(oper)
				modifichers.leggi(rs)
			end if
			rs("nuovo")=0
			rs("cognome")=ucase(request.form("cognome"))
			rs("nome")=ucase(request.form("nome"))
			rs("matricola")=request.form("matricola")
			rs("Contatto")=request.form("Contatto")
			rs("sesso")=request.form("sesso")
			rs("grado")=request.form("grado")
			rs("arma")=request.form("arma")
			rs("lato_arma")=request.form("lato_arma")&""
			rs("nota_dipendente")=request.form("nota_dipendente")
			lingue=request.form("lingue")
			if request.form("lingue-altro")<>"" then lingue=lingue&","&request.form("lingue-altro")
			rs("lingue")=lingue
			if action="update" then
				modifiche_rs=modificheRS.confronta(rs)
				set modificheRS=Nothing
			end if

			nominativo_v=rs("cognome")&" "&rs("nome")
			rs.update
			if action="add" then
				iddip=Get_last_id("dipendenti")
				conn.execute ("insert into utenti_dipendenti (iddip,iduser) values ("&iddip&","&iduser&")")
			end if
			if action="update" then
				modifiche_rs=modifiche_rs&aggiorna_misure_in_ordini(iddip)
			end if
			rs.close
			set rs=nothing
			call add2log(txt_oper&" dipendente "&nominativo_v&" di "&utente_log(iduser)&vbcrlf&modifiche_rs,2)

		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
			js("success")=true
			js("message")="Dipendente aggiunto"
			js("iddip")=iddip
		js.Flush
		set js=Nothing
		call connclose()
		response.end


	case "dimetti_dipendente"
		iddip=request.form("iddip")
		iduser=request.form("iduser")



				set rs_dip=conn.execute ("select * from dipendenti where iddip="&iddip)
				nominativo_dip=rs_dip("cognome")&" "&rs_dip("nome")
				set rs_dip = Nothing
				sql="delete from utenti_dipendenti where iddip="&iddip&" and iduser="&iduser
				conn.execute (sql)
				call add2log("Dimesso dipendente "&nominativo_dip&" da "&utente_log(iduser),2)

			Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
			js("success")=true
			js("message")="Dipendente dimesso"
		js.Flush
		set js=Nothing
		call connclose()
		response.end
	case "elimina_dipendente"
		iddip=request.form("iddip")
		iduser=request.form("iduser")



				set rs_dip=conn.execute ("select * from dipendenti where iddip="&iddip)
				nominativo_dip=rs_dip("cognome")&" "&rs_dip("nome")
				set rs_dip = Nothing
				sql="delete from utenti_dipendenti where iddip="&iddip
				conn.execute (sql)
				sql="delete from dipendenti_misure where iddip="&iddip
				conn.execute (sql)
				sql="delete from dipendenti where iddip="&iddip
				conn.execute (sql)

				call add2log("Eliminato dipendente "&nominativo_dip&" da "&utente_log(iduser),2)

			Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
			js("success")=true
			js("message")="Dipendente dimesso"
		js.Flush
		set js=Nothing
		call connclose()
		response.end





	case "tabella_tipo"

	%>

	<table border="1" cellpadding="2" cellspacing="0" >
		<tr>
			<td></td>
			<td colspan="5" style="text-align: center;">Tipo cliente</td>
		</tr>

		<tr>
			<td></td>
			<td>Società</td>
			<td>Ditta individuale</td>
			<td>Privato</td>
			<td>Associazione</td>
			<td>Sconosciuto</td>
		</tr>
		<tr>
			<td >Nome e cognome</td>
			<td>Facoltativo</td>
			<td>Si</td>
			<td>Si</td>
			<td>Facoltativo</td>
			<td>Almeno uno tra nome, cognome, azienda</td>
		</tr>
		<tr>
			<td >Azienda</td>
			<td>Si</td>
			<td>Facoltativo</td>
			<td>No</td>
			<td>Si</td>
			<td>Almeno uno tra nome, cognome, azienda</td>
		</tr>
		<tr>
			<td >Codice fiscale</td>
			<td>Si: numerico</td>
			<td>Si</td>
			<td>Si</td>
			<td>Si: numerico</td>
			<td>Facoltativo</td>
		</tr>
		<tr>
			<td >Partita iva</td>
			<td>Si</td>
			<td>Si</td>
			<td>No</td>
			<td>No</td>
			<td>Facoltativo</td>
		</tr>

	</table>



	<%
		call connclose()
		response.end

	case "altri_tab"
		iduser=request("iduser")
		Set Js = jsArray()
		'ordini
		sql="select count(*) FROM ordini where eliminato=0 and tipo_documento='ordine' and iduser="& iduser
		if not ha_il_permesso("C5") then
			sql=sql&" and creato_da_admin="&sessioniduser
		end if
		n_ordini=clng(conn.execute( sql)(0))
		if n_ordini>0 then
			Set Js(Null) = jsObject()
			Js(null)("text")="Ordini ("&n_ordini&")"
			Js(null)("url")="pag_adm_user.asp?tab=2&iduser="&iduser
		end if

		if ha_il_permesso("B2") then
			'fatture
			sql="select Count(fatture.IDfat) AS ConteggioDiIDfat, Count(scadenze.idScadenza) AS ConteggioDiidScadenza FROM fatture LEFT JOIN scadenze ON fatture.IDfat = scadenze.idfat WHERE iduser="& iduser
			set rs1= conn.execute (sql)
			fatture=clng(rs1("ConteggioDiIDfat"))
			scadenze=clng(rs1("ConteggioDiidScadenza"))
			rs1.close
			if fatture>0 then
				Set Js(Null) = jsObject()
				Js(null)("text")="Fatture ("&fatture&")"
				Js(null)("url")="pag_adm_user.asp?tab=3&iduser="&iduser
			end if
			if scadenze>0 then
				Set Js(Null) = jsObject()
				Js(null)("text")="Scadenze ("&fatture&")"
				Js(null)("url")="pag_adm_user.asp?tab=9&iduser="&iduser
			end if
		end if	' ha_il_permesso("B2")

		sql="select count(*) as conteggio FROM ordini where eliminato=0 and tipo_documento='ddt' and iduser="& iduser
		set rs1= conn.execute (sql)
		ddt=clng(rs1("conteggio"))
		if ddt>0 then cancellabile=false
		rs1.close
		if ddt>0 then
			Set Js(Null) = jsObject()
			Js(null)("text")="DDT ("&ddt&")"
			Js(null)("url")="pag_adm_user.asp?tab=4&iduser="&iduser
		end if
		if n_ordini>0 then
			Set Js(Null) = jsObject()
			Js(null)("text")="Articoli ordinati"
			Js(null)("url")="pag_adm_user.asp?tab=7&iduser="&iduser
		end if

		if ha_il_permesso("D2") then
			n_ordini_for=clng(conn.execute("select count(*) FROM ordini_fornitori where eliminato=0 and iduser="& iduser )(0))
			if n_ordini_for>0 then
				Set Js(Null) = jsObject()
				Js(null)("text")="Ordini a fornitore ("&n_ordini_for&")"
				Js(null)("url")="pag_adm_user.asp?tab=19&iduser="&iduser
			end if
		end if	'ha_il_permesso("D2")



		sql="select count(iduser) as conteggio FROM preventivi where eliminato=0 and iduser="& iduser
		if not ha_il_permesso("C5") then
			sql=sql&" and creato_da_admin="&sessioniduser
		end if
		set rs1= conn.execute (sql)
		n_preventivi=clng(rs1("conteggio"))
		if n_preventivi>0 then cancellabile=false
		rs1.close
		sql="select count(iduser)  FROM carrello where iduser="& iduser
		set rs1= conn.execute (sql)
		n_carrello=clng(rs1(0))


		set rs1=Nothing
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		js.Flush
		set js=Nothing
	case "cancella_report"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		Session.Contents.remove("report")
		js("isError")="true"
		js("errorMessage")="Questo codice articolo esiste già"
		js.Flush
		set js=Nothing

	case "verifica_new_nfat"
		Set Js = jsObject()
		set rs=conn.execute("select anno, pa from fatture where idfat="&request.form("idfat"))
		anno=rs("anno")
		pa=rs("pa")
		set rs=conn.execute("select * , utenti_intestazioni.nome, utenti_intestazioni.cognome, utenti_intestazioni.azienda from fatture left join utenti_intestazioni on fatture.idintestazione = utenti_intestazioni.id where pa="&pa&" and anno="&anno &" and nfat="&request.form("new_nfat"))
		if not rs.eof then
			js("error")=true
			js("message")="<b>Attenzione</b><br>Questo numero di fattura &egrave; gi&agrave; utilizzato dalla fattura del "&rs("data")&" intestata a<br><b>"&denominazione(rs("nome"),rs("cognome"),rs("azienda"))&"</b>"
		else
			js("error")=false
		end if
		set rs = nothing

		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		js.Flush
		set js=Nothing

	case "verifica_elimina_fattura"
		Set Js = jsObject()
		set rs=conn.execute("select anno, pa, nfat,fattura_sp from fatture where idfat="&request.form("idfat"))
		anno=rs("anno")
		pa=rs("pa")
		fattura_sp=rs("fattura_sp")
		nfat=cint(rs("nfat"))
		max=max_fatt(pa,fattura_sp,anno)
		if max-1=nfat then
			js("error")=false
			js("message")="Elimina l'ultima fattura."
		else
			js("error")=true
			js("message")="<b>Attenzione</b><br>Questa non &egrave; l'ultima fattura.<br><b>Eliminando la fattura avrai un numero mancante nella numerazione</b>"
		end if
		set rs = nothing

		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		js.Flush
		set js=Nothing

	case "chk_settore"

		idsettore=request.form("idsettore")
		idadded=request.form("idadded")
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		cerca_loop=cercaloop(false, idadded,"|"&idsettore&"|")
		if cerca_loop then
			js("success")=false
		else
			js("success")=true
		end if
		js("message")="cercaloop (0,"&idadded&", "&idsettore&")="&cerca_loop
		js.Flush
		set js=Nothing


	case "elimina_incasso"
		add2log "elimina_incasso chiamato:"&queryeform(),0
		idfat=request("idfat")
		if idfat="" then idfat=0
		idord=request("idord")
		if idord="" then idord=0
		idincasso=request("idincasso")
		call elimina_incasso(idincasso)
		call tabella_incassi(idord,idfat,"null",false,false)
		call connclose()
		response.end
	case "elimina_scadenza"
		add2log "elimina_scadenza chiamato:"&queryeform(),0
		idfat=request("idfat")
		if idfat="" then idfat=0
		idscadenza=request("idscadenza")
		call elimina_scadenza(idscadenza)
		call tabella_scadenze(idfat,0)
		call connclose()
		response.end
	case "elimina_scadenza_for"
		add2log "elimina_scadenza chiamato:"&queryeform(),0
		idscadenza=request("idscadenza")
		'Recupero idfat
		idfat=conn.execute("select idfat from scadenze_for where id="&idscadenza)(0)
		conn.execute("delete from scadenze_for where id="&idscadenza)
		call tabella_scadenze_for(idfat)
		call connclose()
		response.end


	case "paga_scadenza"
		add2log "paga_scadenza chiamato:"&queryeform(),0
		idscadenza=request("idscadenza")
		conn.execute("update scadenze set pagato="&request.form("pagato")&" where idscadenza="&idscadenza)
		call tabella_scadenze(0,request.form("idord"))
		call connclose()
		response.end

	case "codiceunico"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		sql="select * from prodotti where codice = '"&request("codice")&"' and idpro<>"&request("idpro")
		set rs=conn.execute(sql)
		if rs.eof then
			js("isError")="false"
		else
			js("isError")="true"
			js("errorMessage")="Questo codice articolo esiste già"
		end if
		set rs=nothing
		js.Flush
		set js=Nothing
	case "elimina_scadenza"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		sql="delete from scadenze where idscadenza = "&request("idscadenza")
		conn.execute sql,num
		if num=0 then
			js("success")="false"
		else
			js("success")="true"
			js("message")="Scadenza eliminata"
		end if
		set rs=nothing
		js.Flush
		set js=Nothing
	case "copia_preferiti"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		iduser=request.form("iduser")
		sql="select * from preferiti inner join prodotti on preferiti.idpro = prodotti.idpro where iduser = "&iduser&" order by codice"
		ordine=0
		set rs = conn.execute (sql)
		Set rs_spettanze = Server.CreateObject("ADODB.Recordset")
		do while not rs.EOF
			rs_spettanze.open "select * from prodotti_spettanze where iduser="&iduser&" and idpro="&rs("idpro")&" and idvara=0 and idvarb=0",conn,3,3
			if rs_spettanze.eof then
				rs_spettanze.addnew
				rs_spettanze("iduser")=iduser
				rs_spettanze("idpro")=rs("idpro")
				rs_spettanze("ordine")=ordine
				ordine=ordine+2
				rs_spettanze.update
			end if
			rs_spettanze.close
			rs.MoveNext
		loop


		'sql="INSERT INTO prodotti_spettanze (iduser, idpro, idvara, idvarb ,ordine) SELECT "&iduser&", idpro, 0,0 ,0 FROM preferiti WHERE iduser="&iduser
		'conn.execute (sql)
			js("success")="true"
			js("message")=ordine/2 & "aggiunti"
			js("numero")=ordine/2
		js.Flush
		set js=Nothing




	'**************************************
	case "salva_misura_dipendente"
		id=request("id")
		iddip=request.form("iddip")
		sql="select dipendenti_misure.* from dipendenti_misure"
		if id<>"" then
			sql=sql&" where id="&id
			txt_log="Modificate taglie "
		end if
		rs.Open sql, conn, 3, 3
		if id="" then
			txt_log="Aggiunto taglie "
			rs.addnew
			rs("iddip")=request.form("iddip")
		else
			iddip=rs("iddip")
		end if

		'on error resume next
		for n = 2 to (rs.Fields.count-1)
			if n=46 then
				rs(n)=request.form("campo"&n)
			else
				rs(n)=request.form("campo"&n)
			end if
			if err.number<>0 then
				call add2log("Errore in salva_misura_dipendente campo "&rs(n).name&":"&request.form("campo"&n)&vbcrlf&queryeform(),0)

						Response.ContentType = "application/json; charset=utf-8"
						Response.CodePage = 65001
						Set Js = jsObject()
							js("success")=false
							js("message")="Errore nel campo "&rs(n).name
						js.Flush
						set js=Nothing
				response.end
			end if
			'session("errorelog")=session("errorelog")&rs.Fields(n).Name&" campo"&n&"<br>"

		next
		rs.update
		id=get_last_id("dipendenti_misure")
		rs.close
		set rs=nothing

		txt_log=txt_log&determina_dipendente(iddip)&aggiorna_misure_in_ordini(iddip)
		call add2log(txt_log,2)
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
			js("success")=true
			js("message")="Misura aggiunta"
		js.Flush
		set js=Nothing






	case "elimina_misura_dipendente"
		id=request("id")
		sql="delete  from dipendenti_misure where id="&id
		conn.execute sql,n
		call add2log("Eliminata "&n&" misura dipendente "&id,2)
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
			js("success")="true"
			js("message")="Misura eliminata"
		js.Flush
		set js=Nothing
	case "salva_grado"
		id=request("id")
		sql="select utenti_gradi.* from utenti_gradi"
		if id<>"" then
			sql=sql&" where id="&id
			txt_log="Modificato "
		end if
		rs.Open sql, conn, 3, 3
		if id="" then
			txt_log="Aggiunto "
			rs.addnew
			rs("iduser")=request.form("iduser")
			rs("nome_grado")=request.form("campo1")
		end if
		'on error resume next
		session("errorelog")=""
		session("errore")=0
		for n = 2 to 14
			rs(n)=request.form("campo"&n)
			session("errorelog")=session("errorelog")&rs.Fields(n).Name&" campo"&n&"<br>"
			if err.number<>0 then
				session("errore")=n
				exit for
			end if

		next
		rs.update
		id=get_last_id("utenti_gradi")
		rs.close
		set rs=nothing
		call add2log(txt_log&" grado "&id&" "&request.form("campo1")&" "&session("errorelog")&" errore a:"&session("errore"),2)
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
			js("success")="true"
			js("message")="Misura aggiunta"
		js.Flush
		set js=Nothing


	case "aggiungi_articolo_spettanze"

		id=request("id")
		sql="select prodotti_spettanze.* from prodotti_spettanze"
		if id<>"" then
			sql=sql&" where id="&id
			txt_log="Modificato "
		else
			sql2="select count(*) from prodotti_spettanze where iduser="&request.form("iduser")&" and idpro="&Request.form("idpro")
			if clng(conn.execute (sql2)(0))>0 then
				Set Js = jsObject()
					js("success")=false
					js("message")="Articolo gi&agrave; esistente"
				js.Flush
				set js=Nothing
				call connclose()
				response.end


			end if

		end if
		rs.Open sql, conn, 3, 3
		if id="" then
			txt_log="Aggiunto "
			rs.addnew
			rs("iduser")=request.form("iduser")
			rs("idpro")=Request.form("idpro")
			idvara=Request.form("idvara")
			if idvara="" then idvara=0
			rs("idvara")=idvara
			idvarb=Request.form("idvarb")
			if idvarb="" then idvarb=0
			rs("idvarb")=idvarb
		end if
		ordine=request.form("ordine")
		if ordine="" then
			ordine=conn.execute("select max(ordine) from  prodotti_spettanze where iduser="&request.form("iduser"))(0)
			if isnull(ordine) then ordine=2
		end if
		rs("ordine")=ordine
		iduser=rs("iduser")
		rs.update
		id=get_last_id("dipendenti_misure")
		rs.close
		set rs=nothing
		call riordina_sql("prodotti_spettanze"," where iduser="&iduser)
		call add2log(txt_log&" articolo spettanze "&id&" "&request.form("id"),2)
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
			js("success")="true"
			js("message")="Articolo aggiunto"
		js.Flush
		set js=Nothing
	case "elimina_articolo_spettanze"
		'on error resume next
		id=request("id")
		sql="delete from prodotti_spettanze where id="&id
		conn.execute sql,n
		if err.number<>0 then
			response.write sql

		end if
		on error goto 0
		call add2log("Eliminato "&n&" articolo spettanze "&id,2)
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
			js("success")="true"
			js("message")="Articolo rimosso"
		js.Flush
		set js=Nothing
		'call add2log ("Rimosso articolo spettanze id:"&id,2)
		'call riordina_sql("select ordine from prodotti_spettanze where iduser="&request.form("iduser")&" order by ordine")
	'**************************************
	case "elimina_seriale"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()

		id=request.form("id")
		set rs=conn.execute("select numeri_seriali.seriale, prodotti.codice, prodotti.articolo from numeri_seriali inner JOIN prodotti on numeri_seriali.idpro = prodotti.idpro where id= "&id)
		if not rs.eof then
			arrRS = rs.GetRows()
			sql="delete from numeri_seriali where id="&id
			conn.execute sql,n
			call add2log("Eliminato numero di serie "&arrRS(0,0)&" "&arrRS(1,0)&" "&arrRS(2,0),2)
			js("success")=true
			js("message")="Numero seriale eliminato"
		else
			js("success")=false
			js("message")="Numero seriale non trovato"
		end if
		set rs = Nothing
		js.Flush
		set js=Nothing
	case "elimina_seriale_da_ordine"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()

		id=request.form("id")
		set rs=conn.execute("select numeri_seriali.seriale, prodotti.codice, prodotti.articolo from numeri_seriali inner JOIN prodotti on numeri_seriali.idpro = prodotti.idpro where id= "&id)
		if not rs.eof then
			arrRS = rs.GetRows()
			sql="update numeri_seriali set idord=NULL, iddett=NULL where id="&id
			conn.execute sql,n
			call add2log("Rimosso numero di serie "&arrRS(0,0)&" "&arrRS(1,0)&" "&arrRS(2,0)&" dall'ordine",2)
			js("success")=true
			js("message")="Numero seriale dinuovo disponibile"
		else
			js("success")=false
			js("message")="Numero seriale non trovato"
		end if
		set rs = Nothing
		js.Flush
		set js=Nothing





	case "ordina_immagini"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()

		idpro=request("idpro")
		getItems = Request.QueryString("id_file[]")
		'split them junks into array
		getItems = Split(getItems, ",")
		'For each one, define and then execute the sql statement
		i_max=UBound(getItems)
		i_max=i_max*2
		For i = LBound(getiTems) TO i_max step 2
		   sql_update = "UPDATE files SET ordine = " & i & " WHERE idfiles = " & getItems(i/2)


		   'execute this bitch
		   conn.Execute sql_update
		   if i=0 then
			   set rs_files=conn.execute ("select files.* from files where idfiles="& getItems(i/2))
			   rs.Open "select prodotti.* from prodotti where idpro="&idpro, conn, 1, 3
			   rs("fileimg")=rs_files("filename")
			   rs.update
			   rs.close


		   end if
		   'below used for debugging - comment out
		   'Response.Write("ID " & getItems(i) & " was set to position " & i)
			'response.write "."
		Next
		js("success")=true
		js("Message")=(i/2)&" Immagini ordinate"
		js.Flush
		set js=Nothing
	case "elimina_immagine"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		idfiles=request("id_file")
		idpro=request("idpro")
		set rs_all=conn.execute( "select * FROM files where idfiles="& idfiles )
		if not rs_all.eof then
			'Memorizzo recordset in variabili
			rs_path=rs_all("path")
			rs_filename=ucase(rs_all("filename"))
			rs_cosa=rs_all("cosa")
			rs_idcosa=rs_all("idcosa")
			set rs_all=nothing
			if rs_cosa=4 then
				set rs_pro=conn.execute ("select prodotti.* from prodotti where idpro="&rs_idcosa)
				testo=" immagine [articolo="& rs_pro("idpro")&"]"&rs_pro("codice")&"[/articolo] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"
				set rs_pro=nothing
			else
				testo="Eliminato [allegato="&idfiles&"]"&rs_filename&"[/allegato] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin] "
			end if
			Set fso = CreateObject("Scripting.filesystemObject")
			'Elimino immagine principale
			pathFile=server.MapPath(rs_path&rs_filename)
			if fso.FileExists(pathFile) then
				testo=testo&VbCrLf&"File "&rs_filename&" eliminato"
				fso.DeleteFile(pathFile)
			else
				testo=testo&VbCrLf&"File "&rs_filename&" non trovato"
			end if
			'Elimino immagine in /s
			pathFile=server.MapPath(rs_path&"s/"&rs_filename)
			if fso.FileExists(pathFile) then
				testo=testo&VbCrLf&"File s/"&rs_filename&" eliminato"
				fso.DeleteFile(pathFile)
			else
				testo=testo&VbCrLf&"File s/"&rs_filename&" non trovato"
			end if
			'Elimino immagine in /bk
			pathFile=server.MapPath(rs_path&"bk/"&rs_filename)
			if fso.FileExists(pathFile) then
				testo=testo&VbCrLf&"File bk/"&rs_filename&" eliminato"
				fso.DeleteFile(pathFile)
			else
				testo=testo&VbCrLf&"File bk/"&rs_filename&" non trovato"
			end if


			conn.execute "delete FROM files where idfiles="& idfiles,num
			testo=testo&VbCrLf&num&" records eliminati"

			'riordino
			ordine=0
			Set rs_all = Server.CreateObject("ADODB.Recordset")
			sql="select files.* from files where cosa=4 and idcosa="&idpro&" order by ordine;"
			rs_all.open sql, conn,1,3
			if rs_all.eof then
			   rs.Open "select prodotti.* from prodotti where idpro="&idpro, conn, 1, 3
			   rs("fileimg")="no_img.jpg"
			   rs.update
			   rs.close
				js("success")=true
				testo="Eliminata ultima"&testo
				js("Message")="Ultima immagine eliminata"
			else
				js("success")=true
				testo="Eliminata"&testo
				js("Message")="Immagine eliminata"
			end if
			do while not rs_all.eof
				if ordine=0 then
				   rs.Open "select prodotti.* from prodotti where idpro="&idpro, conn, 1, 3
				   rs("fileimg")=rs_all("filename")
				   rs.update
				   rs.close
				end if
				rs_all("ordine")=ordine
				rs_all.update
				ordine=ordine+2
				rs_all.movenext
			loop
		else
			set rs_all=nothing
			testo="Allegato non trovato"
			js("success")=false
			js("Message")=testo
		end if
		add2log testo,2
		js.Flush
		set js=Nothing
	case "ridimensiona_immagine"
		maxWidth=800*2
		maxHeight=1120*2

	'on error resume next
		dim testo_log
		Set Js = jsObject()
		idfiles=request("id_file")
		modificato=false
		sql="select * from files where idfiles="&idfiles
		Set rs_files = Server.CreateObject("ADODB.Recordset")
		rs_files.Open sql , conn, 3,2

		if not rs_files.eof then

			'ricostruisco il nome per il nuovo file
			savefolder=upl_img_cat	'"/public/img_prod/"
			inizio_nome="IMG_"
			id_tipo_allegato=rs_files("idcosa")
			if rs_files("ordine")>2 then
				n_file= rs_files("ordine")/2
			else
				n_file=1
				rs_files("ordine")=0
			end if
			fine_nome=replace(time(),":","")

			fileName=rs_files("filename")
			if InStr(fileName, ".") > 0 then
				estensione = "." & Right(fileName, Len(fileName) - InStrRev(fileName, "."))
			else
				estensione = ".tmp"
			end if

			newname = inizio_nome&id_tipo_allegato&"_"&n_file&"_"&fine_nome&estensione

			Set Jpeg = Server.CreateObject("Persits.Jpeg")
			' Open source image
			imgurl=rs_files("path")&rs_files("filename")
			Jpeg.Open Server.MapPath(imgurl)
			' 800px x 1120px, rapporto larghezza/altezza: 0.70 - 0.73

			AspectRatio=jpeg.OriginalWidth/jpeg.OriginalHeight
			testo_log="Dimensioni immagine:"&jpeg.OriginalWidth&"x"&jpeg.OriginalHeight&":"&roundup(AspectRatio,2)
			if AspectRatio>0.73 then	'Altezza troppo bassa, immagine larga
				'Calcolo altezza ottimale
				optimalHeight=jpeg.OriginalWidth/0.7142

				testo_log=testo_log&"<br>Aumento altezza a:"&int(optimalHeight)

				'Calcolo delta
				deltaHeight=int((optimalheight-jpeg.OriginalHeight)/2)

				'Croppo in negativo solo in Y
				jpeg.Canvas.Brush.Color = &HFFFFFF
				jpeg.Crop 0, deltaHeight*-1, jpeg.OriginalWidth , jpeg.originalHeight + deltaHeight
				modificato=true
			elseif AspectRatio<0.70 then	'Larghezza troppo bassa, immagine alta
				'Calcolo larghezza ottimale
				optimalWidth=jpeg.OriginalHeight*0.7142
				testo_log=testo_log&"<br>Aumento larghezza a:"&int(optimalWidth)

				'Calcolo delta
				deltaWidth=int((optimalWidth-jpeg.OriginalWidth)/2)

				'Croppo in negativo solo in X
				jpeg.Canvas.Brush.Color = &HFFFFFF
				jpeg.Crop deltaWidth*-1, 0, jpeg.OriginalWidth+deltaWidth , jpeg.OriginalHeight
				modificato=true
			else
				testo_log=testo_log&"<br>Immagine conforme"
			end if

			jpeg.PreserveAspectRatio  = true

			fattore_riduzione_w=jpeg.width/maxWidth
			fattore_riduzione_h=jpeg.height/maxHeight
			testo_log=testo_log&" Fattori di riduzione: w"&fattore_riduzione_w&" h"&fattore_riduzione_h
			if fattore_riduzione_w<=1 and fattore_riduzione_h<=1 then
				'nessuna riduzione
				testo_log=testo_log&" Nessuna riduzione"
			elseif fattore_riduzione_w>=fattore_riduzione_h then
				jpeg.Width= maxWidth
				'Image.Height= image.OriginalHeight/fattore_riduzione_w
				testo_log=testo_log&" Regolato larghezza w="&jpeg.Width&" h="&jpeg.Height
				modificato=true
			else
				jpeg.Height= maxHeight
				'Image.Width=image.OriginalWidth/fattore_riduzione_h
				testo_log=testo_log&" Regolato altezza w="&jpeg.Width&" h="&jpeg.Height
				modificato=true
			end if








			if modificato then
				testo_log=testo_log&"<br>Dimensione finale:"&jpeg.width&"x"&jpeg.height&":"&roundup(jpeg.width/jpeg.height,2)
				rs_files("chk")=chk_immagine(jpeg.width,jpeg.height)
				Jpeg.Save Server.MapPath(imgurl)

				Dim fs
				' Creo una istanza dell'offetto filesystemObject
				Set fs = Server.CreateObject("Scripting.filesystemObject")

				' rinominiamo il file da , a
				fs.MoveFile Server.MapPath(imgurl),Server.MapPath(savefolder&newname)
				set f=fs.GetFile(Server.MapPath(savefolder&newname))
				fsize=f.Size
				' Faccio pulizia
				set f=Nothing

				Set fs = Nothing
				rs_files("path")=savefolder
				rs_files("filename")=newname
				rs_files("filesize")=fsize
				idpro=rs_files("idcosa")
				idfile=rs_files("idfiles")
				rs_files.update
				rs_files.close
				set rs_files=nothing
				'Apro prodotti per modificare il nome del file
				Set rs_prodotti = Server.CreateObject("ADODB.Recordset")
				sql="select prodotti.* from prodotti where idpro="&idpro
				rs_prodotti.Open sql , conn, 0,3
				if not rs_prodotti.eof then
					if rs_prodotti("fileimg")=fileName then rs_prodotti("fileimg")=newname
					rs_prodotti.update
					txt_log="Ridimensionata immagine [articolo="& rs_prodotti("idpro")&"]"&rs_prodotti("codice")&"[/articolo] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"
				end if
				rs_prodotti.Close
				set rs_prodotti=Nothing
			end if
			js("success")=true
			js("Message")=testo_log
			add2log txt_log&vbcrlf&testo_log&vbcrlf&imgurl&vbcrlf&savefolder&newname&vbcrlf&"idfiles:"&idfile,2
		end if
		If Err.Number <> 0 Then
		   js("success")=false
		   js("Message")= Err.Description& "<br><br>"


		End If
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001

		js.Flush
		set js=Nothing



	case "aggiorna_elenco_immagini"
		id_tipo_allegato=request("id_tipo_allegato")
		tipo_allegato=request("tipo_allegato")
		idpro=request("idpro")
	%>
	          <!--#include virtual="/sub_allegati_img.asp" -->
	<%
	case "aggiorna_elenco_allegati"
		id_tipo_allegato=request("id_tipo_allegato")
		tipo_allegato=request("tipo_allegato")
		idpro=request("idpro")
	%>
	          <!--#include virtual="/sub_allegati.asp" -->
	<%

	case "load_descrizione_img"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		idfile=request("id_file")
		Set Js = jsObject()
		rs.Open "select * FROM files where idfiles="& idfile, conn, 1, 3
		if rs.eof then
			js("success")=false
			js("Message")="File non trovato"
		else
			js("success")=true
			js("descrizione_img")=rs("descrizione")
		end if
		rs.close
		js.Flush
		set js=Nothing
	case "save_descrizione_img"
		id_file=request("id_file")
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		sql="select files.* FROM files where idfiles="& id_file
		rs.Open sql, conn, 1, 3
		if rs.eof then
			js("success")=false
			js("Message")="File non trovato"
		else
			rs("descrizione")=request("descrizione_img")
			rs.update
			js("success")=true
			js("Message")="Descrizione immagine salvata"
		end if
		rs.close
		js.Flush
		set js=Nothing
	case "riordina_articoli"
		idsettore=request("idsettore")
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		ordine=split(request("id[]"),",")
		dim ordine_Dictionary
		Set ordine_Dictionary=Server.CreateObject("Scripting.Dictionary")
		for n=0 to ubound(ordine)
		ordine_Dictionary.add cint(ordine(n)),n
		text=text&ordine(n)&"="&n&","
		next
		text=text&" Aggiunti "&n&" record"
		Set Js = jsObject()
		sql="select settori_prodotti.* FROM settori_prodotti where idsettore="& idsettore
		rs.Open sql, conn, 1, 3
		do while not rs.eof
		ordinetxt=ordine_Dictionary.item(cint(rs("idpro")))
		text=text&"loop"&rs("id")&"="&ordinetxt&","
		rs("ordine")=ordinetxt
		rs.update
		rs.movenext
		loop
		rs.close
		js("success")=true
		js("Message")="Ordine articoli salvato"&text
		js.Flush
		set js=Nothing

		'recupero informazioni per log
		set rs_settore=conn.execute ("select settori.* from settori where idsettore="&idsettore)
		add2log "Riordinati articoli [settore="&idsettore&"]"&rs_settore("nome_settore")&"[/settore]eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]",2
		set rs_settore=Nothing

	case "assegna_upc_ean"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		idpro=request("idpro")
		set rs=conn.execute("SELECT *  FROM  `upc_ean` WHERE idpro IS NULL LIMIT 0 , 1")
		if not rs.eof then
		   sql = "UPDATE upc_ean SET idpro = " & idpro & " WHERE upc_ean = '"& rs("upc_ean")&"'"
			conn.execute sql,result
			if result=1 then
				js("success")=true
				js("message")="Codice assegnato"
				js("upc_ean")=rs("upc_ean")
			else
				js("success")=false
				js("message")="Errore nell'assegnazione"
			end if
		else
			js("success")=false
			js("message")="Non ci sono codici disponibili"

		end if


		js.Flush
		set js=Nothing
	case "rimuovi_upc_ean"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		idpro=request("idpro")
	   sql = "UPDATE upc_ean SET idpro = NULL WHERE idpro = "& idpro
		conn.execute sql,result
		if result=1 then
			js("success")=true
			js("message")="Codice assegnato"
		else
			js("success")=false
			js("message")="Errore nella rimozione"
		end if
		js.Flush
		set js=Nothing







	case "elimina_file"
		id_file=request("id_file")
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		Set rs_all = Server.CreateObject("ADODB.Recordset")
		rs_all.Open "select * FROM files where idfiles="& id_file , conn, 3,2
		if not rs_all.eof then
			testo="Eliminato [allegato="&idfiles&"]"&rs_all("filename")&"[/allegato] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin] "
			Set fso = CreateObject("Scripting.filesystemObject")
			if fso.FileExists(server.MapPath(rs_all("path")&rs_all("filename"))) then
				testo=testo&VbCrLf&"File eliminato"
				fso.DeleteFile(server.MapPath(rs_all("path")&rs_all("filename")))
			else
				testo=testo&VbCrLf&"File non trovato"
			end if
			rs_all.delete
			js("success")=true
			js("Message")="File eliminato"
		else
			testo="Allegato non trovato"
			js("success")=false
			js("Message")="File non trovato"
		end if
		rs_all.close
		set rs_all=nothing
		add2log testo,2
		js.Flush
		set js=Nothing
	case "reset_messaggio"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		js("success")=true
		js("Message")="Messaggio eliminato"
		session("visualizza_messaggio")=""
		js.Flush
		set js=Nothing
	case "invia_pw"
	%>

	<!--#include virtual="/mail_registrazione_inc.asp" -->
	<%
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001

		Set Js = jsObject()

		res= invia_mail_registrazione(request("iduser"),false)

		if res  then
			js("success")=true
			js("Message")="Email inviata"
			js("data_invio")=now()
		else
			js("success")=false
			js("Message")="Email non inviata"
			js("data_invio")=now()
		end if
		js.Flush
		set js=nothing
		set js=Nothing

	case "errore"
		add2log request("txt_errore"),3
	case "magazzino"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		idpro=request("idpro")
		idvara=request("idvara")
		idvarb=request("idvarb")
		promozione=false
		sql="select magazzino.idpro, magazzino.idvara, magazzino.idvarb, magazzino.quantita_magazzino FROM magazzino WHERE (((magazzino.idpro)="&idpro&") AND ((magazzino.idvara)="&idvara&") AND ((magazzino.idvarb)="&idvarb&"));"
		set rs_magazzino=conn.execute(sql)
		Set Js = jsObject()
		if rs_magazzino.eof then
			js("success")=false
		else
			js("success")=true
			js("quantita_magazzino")=rs_magazzino("quantita_magazzino")
		end if
		trovato=false

		sql="select prezzo, promozione, prodata, sconto, prezzo_ve_min from prodotti where idpro="&idpro
		set rs=conn.execute(sql)
		if cdbl(rs("prezzo_ve_min"))=0 then
			prezzo=rs("prezzo")
			trovato=true
		end if
		if rs("promozione")=1 then
			sconto=cdbl(rs("sconto"))
			promozione=true
		end if
		if not trovato then
			sql="select prezzo_ve_va from varianti_a where idpro="&idpro&" and idvara="&idvara
			set rs=conn.execute(sql)
			if not rs.eof then
				if cdbl(rs("prezzo_ve_va"))>0 then
					prezzo=rs("prezzo_ve_va")
					trovato=true
				end if
			end if
		end if
		if promozione then
			prezzo=cdbl(prezzo)
			prezzo=roundup(prezzo*(1-sconto/100),2)
		end if
		js("prezzo")=prezzo
		js.Flush
		set js=Nothing
	case "ordina_for"
		idpro=request("idpro")
		idvara=request("idvara")
		idvarb=request("idvarb")
		quantita=request("quantita")
		idfor=request("idfor")

		idord_fornitore=crea_ordine_fornitore(idfor)
		'Cerco valore per ordinamento articoli
		set max_ordine=conn.execute ("select Max(ordini_fornitori_dett.ordine) AS MaxDiordine FROM ordini_fornitori_dett WHERE (((ordini_fornitori_dett.idord)="&idord_fornitore&"));")
		MaxDiordine=max_ordine("MaxDiordine")
		if isnull(MaxDiordine) then MaxDiordine=0
		MaxDiordine=MaxDiordine+2
		sql="select prodotti.*, varianti_a.variante_a, varianti_b.variante_b FROM (prodotti LEFT JOIN varianti_a ON prodotti.IDpro = varianti_a.idpro) LEFT JOIN varianti_b ON prodotti.IDpro = varianti_b.idpro where prodotti.idpro="&idpro
		if idvara<>"undefined" then
			sql=sql&" and idvara="&idvara
		else
			idvara=0

		end if
		if idvarb<>"undefined" then
			sql=sql&" and idvarb="&idvarb
		else
			idvarb=0
		end if

		set rs=conn.execute (sql)


		Set dett_fornitori = Server.CreateObject("ADODB.Recordset")
		dett_fornitori.Open "select ordini_fornitori_dett.* FROM ordini_fornitori_dett", conn, 3, 3
		dett_fornitori.addnew
		dett_fornitori("idord")=idord_fornitore
		dett_fornitori("idpro")=idpro
		dett_fornitori("var1")=rs("variante_a")
		dett_fornitori("var2")=rs("variante_b")
		dett_fornitori("um")=rs("um")
		dett_fornitori("quantita")=quantita
		dett_fornitori("prezzo")=rs("costo")
		dett_fornitori("codice_ordine")=rs("codice")
		dett_fornitori("articolo_ordine")=rs("articolo")
		dett_fornitori("variante1_ordine")=rs("variante1")
		dett_fornitori("variante2_ordine")=rs("variante2")
		dett_fornitori("idvara")=idvara
		dett_fornitori("idvarb")=idvarb
		dett_fornitori("ordine")=MaxDiordine
		dett_fornitori.update
		session("proc_ordina")=session("proc_ordina")&" aggiunto articolo "&rs("codice")&", "
		session("report")=session("report")&"Aggiunto articolo "&rs("codice")&" a ordine fornitore "&session("nord_fornitore")&"<br>"
		Set Js = jsObject()
		js("success")=true
		js("nord")=session("nord_fornitore")
		js("idord")=idord_fornitore
		js.Flush
		add2log session("report"),2
		set js=Nothing
	case "cerca_db"
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		idpro=request("idpro")
		idvara=request("idvara")
		idvarb=request("idvarb")
		if idvara="undefined" then idvara=0
		if idvarb="undefined" then idvarb=0
		on error resume next
		db_ven=clng(conn.execute("select count(*) FROM magazzino where idpro="&idpro&" and idvara="&idvara&" and idvarb="&idvarb&" and db_ven=1" )(0))
		if err.number<>0 then
			response.write "select count(*) FROM magazzino where idpro="&idpro&" and idvara="&idvara&" and idvarb="&idvarb&" and db_ven=1"
			response.end
		end if
		db_acq=clng(conn.execute("select count(*) FROM magazzino where idpro="&idpro&" and idvara="&idvara&" and idvarb="&idvarb&" and db_acq=1" )(0))

		js("db_ven")=db_ven
		js("db_acq")=db_acq
		js.Flush
		set js=Nothing

	case else
			add2log "Richiesta non riconosciuta"& queryeform()&"<br>"&Request.ServerVariables("HTTP_REFERER"),3
	end select
	call rsclose()
	response.end





end if 'Oper

'Funzioni per select2
dim select2,autocomplete
term=replace(request("term"),"'","''")
select2=request("select2")
if select2<>"" then
i=0
max=50
	Response.ContentType = "application/json; charset=utf-8"
	Response.CodePage = 65001
	Set Js = jsArray()
	select case select2
	case "articoli_sel2"
		sql="select idpro,articolo,codice from prodotti where  visibilita<=2 and "

		opt=request.querystring("opt")
		select case opt
			case "seriali"
				sql=sql&" richiedi_seriale=1 and "
		end select



		if request("opt")="attivi" then sql=sql&" visibilita<=2 and vendita=1 and "
		sql=sql&" CONCAT_WS(  ' ', idpro, articolo, codice ) like '%"&term&"%' order by articolo;"
		set rs=conn.execute(sql)
		do while (not rs.eof) and i<max
			Set Js(Null) = jsObject()
			Js(null)("text")=rs("codice")&" "&rs("articolo")
			Js(null)("articolo")=rs("articolo")
			Js(null)("codice")=rs("codice")
			Js(null)("id")=rs("idpro")
			rs.movenext
			i=i+1
		loop
	case "tags"
		sql="select * from tags where  nometag"
		if term<>"" then sql=sql&" like '%"&term&"%'"
		sql=sql& " order by nometag;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("idtag")
			Js(null)("text")=rs("nometag")
			rs.movenext
			i=i+1
		loop
	case "ordine_sostituzione"
		iduser=request.querystring("iduser")
		sql="select ordini.nord, ordini.data, ordini.idord from ordini left join ordini_sostituzioni on ordini.idord = ordini_sostituzioni.idord where tipo_documento='ordine' and ordini_sostituzioni.idsostituzione is null and ordini.iduser="&iduser
		'if term<>"" then sql=sql&" like '%"&term&"%'"
		sql=sql& " order by nord;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("idord")
			Js(null)("text")=rs("nord")&" del "&formatdatetime(rs("data"),2)
			rs.movenext
			i=i+1
		loop




	case "consigliati"

	sql="select articoli_consigliati.id_pro_consigliati, articoli_consigliati.nome_ragruppamento, Count(prodotti_consigliati.Id) AS ConteggioDiId FROM articoli_consigliati LEFT JOIN prodotti_consigliati ON articoli_consigliati.id_pro_consigliati = prodotti_consigliati.idconsigliato GROUP BY articoli_consigliati.id_pro_consigliati, articoli_consigliati.nome_ragruppamento having  nome_ragruppamento"
		'sql="select * from articoli_consigliati where  nome_ragruppamento"
		if term<>"" then sql=sql&" like '%"&term&"%'"
		sql=sql& " ORDER BY articoli_consigliati.nome_ragruppamento;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("id_pro_consigliati")
			Js(null)("text")=rs("nome_ragruppamento")& " ("&rs("ConteggioDiId")&" articoli)"
			rs.movenext
			i=i+1
		loop

	case "modelli"
		sql="select * from marca_modello where modello=true and nome "
		if term<>"" then sql=sql&" like '%"&term&"%'"
		sql=sql& " order by nome;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("id")
			Js(null)("text")=rs("nomeconmarca")
			rs.movenext
			i=i+1
		loop

	case "mepa_categorie"
		sql="select * from mepa_categorie where descrizione "
		if term<>"" then sql=sql&" like '%"&term&"%'"
		sql=sql& " order by descrizione;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("id")
			Js(null)("text")=rs("descrizione")
			rs.movenext
			i=i+1
		loop


	case "seriali"

		sql="SELECT numeri_seriali.seriale, prodotti.codice, prodotti.articolo, ordini.idord, ordini.nord, ordini.data FROM ( numeri_seriali INNER JOIN prodotti ON numeri_seriali.idpro = prodotti.idpro ) LEFT JOIN ordini ON numeri_seriali.idord = ordini.idord where seriale"
		sql=sql&" like '%"&term&"%'"

		sql=sql& " order by seriale;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("text")=rs("seriale")
			js(null)("articolo")=rs("codice")&" "&rs("articolo")
			js(null)("idord")=rs("idord")
			if not isnull(rs("idord")) then
				js(null)("text2")="<b>ordine "&rs("nord")&"</b>/"&year(rs("data"))
			else
				js(null)("text2")=""
			end if
			rs.movenext
			i=i+1
		loop

	case "settori"
		cerca_loop=request("cercaloop")

		if cerca_loop="" then
			cerca_loop=false
		else
			cerca_loop=true
		end if
		escludi=request("escludi")
		sql="select * from settori where nome_settore"
		if term<>"" then sql=sql&" like '%"&term&"%'"
		if escludi=0 then
			cerca_loop=false
			escludi=""
		elseif escludi<>""  then
			sql=sql& " and idsettore<>"&escludi
		end if
		sql=sql& " order by nome_settore;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("idsettore")
			Js(null)("text")=gerarchia_settori_bkw(rs("idsettore"),"")
			if cerca_loop then
				Js(null)("disabled")=cercaloop(false,rs("idsettore"),"|"&escludi&"|")
			end if

			rs.movenext
			i=i+1
		loop

	case "fornitori_sel2"

		sql="select i.iduser, i.Azienda, i.Cognome, i.Nome FROM utenti_intestazioni i inner join utenti u on i.id=u.idintestazione  WHERE u.fornitore=true and "


		if isnumeric(term) then
			sql=sql&" i.iduser="&term
		else

			parole=split(term," ")
			where="INSTR( CONCAT_WS(  ' ', i.azienda, i.nome, i.cognome ) ,  '" &parole(0)& "' ) >0"
			for i = 1 to ubound(parole)
				where="INSTR( CONCAT_WS(  ' ', i.azienda, i.nome, i.cognome ) ,  '" &parole(i)& "' ) >0"
				'where = where & " and instr( ricerca,'" &parole(i)& "')>0 "
			i=i+1
			next
			sql=sql&where
		end if
		sql=sql&" order by i.azienda"
		set rs=conn.execute(sql)
		If Not rs.EOF Then
			arrRS = rs.GetRows()
		end if
		Set rs = Nothing


		If IsArray(arrRS) Then
			maxtemp=UBound(arrRS, 2)
			if maxtemp>max then maxtemp=max
			For i = LBound(arrRS, 2) To maxtemp
				Set Js(Null) = jsObject()
				Js(null)("text")=arrRS(1, i)&" "&arrRS(3, i)&" "&arrRS(2, i)&" Id:"&arrRS(0, i)
				Js(null)("id")=arrRS(0, i)
			Next
			Erase arrRS
		End If

	case "utenti_sel2"

		'sql="select utenti.iduser, utenti.Azienda, utenti.Cognome, utenti.Nome FROM utenti WHERE  "
		sql="select u.iduser, i.Azienda, i.Cognome, i.Nome, u.idagente FROM utenti_intestazioni i inner join utenti u on i.id=u.idintestazione  WHERE secondario=0 and "
		if session("idagente")<>"" and not ha_il_permesso("C5") then
			sql=sql&" idagente="&session("idagente") &" and "
		end if
		if isnumeric(term) then
			sql=sql&" i.iduser="&term
		else

			parole=split(trim(term)," ")
			where="INSTR( CONCAT_WS(  ' ', i.azienda, i.nome, i.cognome ) ,  '" &parole(0)& "' ) >0"
			for i = 1 to ubound(parole)
				where=where&" and INSTR( CONCAT_WS(  ' ', i.azienda, i.nome, i.cognome ) ,  '" &parole(i)& "' ) >0"
				'where = where & " and instr( ricerca,'" &parole(i)& "')>0 "
			i=i+1
			next
			sql=sql&where
		end if
		sql=sql&" order by i.azienda"
		set rs=conn.execute(sql)
		If Not rs.EOF Then
			arrRS = rs.GetRows()
		end if
		Set rs = Nothing


		If IsArray(arrRS) Then
			maxtemp=UBound(arrRS, 2)
			if maxtemp>max then maxtemp=max
			For i = LBound(arrRS, 2) To maxtemp
				Set Js(Null) = jsObject()
				Js(null)("text")=arrRS(1, i)&" "&arrRS(3, i)&" "&arrRS(2, i)&" Id:"&arrRS(0, i)
				Js(null)("id")=arrRS(0, i)
			Next
			Erase arrRS
		End If



	case "varianti"
		Set Js = jsObject()

		sql="select * from varianti_a where idpro="&term
		set rs=conn.execute(sql)
		if not rs.eof then
			Js("var_a")="true"
		else
			Js("var_a")="false"
		end if
		sql="select * from varianti_b where idpro="&term
		set rs=conn.execute(sql)
		if not rs.eof then
			Js("var_b")="true"
		else
			Js("var_b")="false"
		end if

	case "vara"
		if request.querystring("prezzo")="" then
			sql="select * from varianti_a where idpro="&term
			set rs=conn.execute(sql)
			do while (not rs.eof)
				Set Js(Null) = jsObject()
				Js(null)("text")=trim(rs("variante_a"))
				Js(null)("id")=rs("idvara")
				rs.movenext
				i=i+1
			loop
		else
			sql="select * from varianti_a where idpro="&term
			set rs=conn.execute(sql)
			do while (not rs.eof)
				Set Js(Null) = jsObject()
				if cdbl(rs("prezzo_ve_va"))>0 then
					Js(null)("text")=trim(rs("variante_a") &" - euro " & formatnumber(rs("prezzo_ve_va"),2))
				else
					Js(null)("text")=trim(rs("variante_a"))
				end if
				Js(null)("id")=rs("idvara")
				Js(null)("prezzo")=rs("prezzo_ve_va")
				rs.movenext
				i=i+1
			loop
		end if
	case "varb"
			sql="select * from varianti_b where idpro="&term
			set rs=conn.execute(sql)
			do while (not rs.eof)
				Set Js(Null) = jsObject()
				Js(null)("text")=trim(rs("variante_b"))
				Js(null)("id")=rs("idvarb")
				rs.movenext
				i=i+1
			loop

	case "dipendenti"
		sql="select dipendenti.iddip, cognome ,nome, matricola from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where  utenti_dipendenti.iduser="&request("iduser")

		if term<>"" then sql=sql&" and concat_ws(' ',cognome,nome) like '%"&term&"%'"
		sql=sql& " order by cognome,nome;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("iddip")
			txt=rs("cognome")&" "&rs("nome")
			if rs("matricola")<>"" then
				txt=txt&" ("&rs("matricola")&")"
			end if
			Js(null)("text")=txt
			rs.movenext
			i=i+1
		loop
	case "tuttidipendenti"
		sql="select dipendenti.iddip, cognome ,nome, matricola from dipendenti left join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip "

		if term<>"" then sql=sql&" where concat_ws(' ',cognome,nome) like '%"&term&"%'"
		sql=sql& " order by cognome,nome;"
		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("iddip")
			txt=rs("cognome")&" "&rs("nome")
			if rs("matricola")<>"" then
				txt=txt&" ("&rs("matricola")&")"
			end if
			Js(null)("text")=txt
			rs.movenext
			i=i+1
		loop
	case "destinatari_appunti"
		sql="select *, NULL as destinatario from admin where iduser>1 and fine is null and iduser<>"&sessionIDUser&" order by nominativo "

		set rs=conn.execute(sql)
		do while (not rs.eof)
			Set Js(Null) = jsObject()
			Js(null)("id")=rs("iduser")
			Js(null)("text")=nomebreve(rs("nominativo"),"")
			rs.movenext
			i=i+1
		loop





	case "tutto"
		i=0
		max=25
		strArr=""
		salta=false

		if isnumeric(term) then
			if ha_il_permesso("C5") then
				'Cerca ordini
				sql="select  * from ordini where nord="&term &" order by anno desc, nord desc"
				set rs=conn.execute(sql)
				do while (not rs.eof) and i<max
					Set Js(Null) = jsObject()
					Js(null)("id")=rs("idord")
					Js(null)("text")="Ordine "&term&" del "&formatdatetime(rs("data"),2)
					Js(null)("t")="O"
					rs.movenext
					i=i+1
				loop
			end if
			if ha_il_permesso("B2") then
				'Cerca fatture
				sql="select  * from fatture where nfat="&term&" order by idfat desc"
				set rs=conn.execute(sql)
				do while (not rs.eof) and i<max
					Set Js(Null) = jsObject()
					Js(null)("id")=rs("idfat")
					Js(null)("text")="Fattura "&term&" del "&formatdatetime(rs("data"),2)
					Js(null)("t")="F"
					rs.movenext
					i=i+1
				loop
				'Cerca DDT
				sql="select  idord, nord, data, tipo_documento from ordini where tipo_documento='ddt' and nord="&term&" order by anno desc, nord desc"
				set rs=conn.execute(sql)
				do while (not rs.eof) and i<max
					Set Js(Null) = jsObject()
					Js(null)("id")=rs("idord")
					Js(null)("text")="DDT "&rs("nord")&" del "&formatdatetime(rs("data"),2)
					Js(null)("t")="D"
					rs.movenext
					i=i+1
				loop
			end if

			'Cerca utente per ID
			'sql="select  * from utenti where utenti.iduser="&term&""
			sql="select i.iduser, i.iduser, i.Azienda, i.Cognome, i.Nome, u.idagente FROM utenti_intestazioni i inner join utenti u on i.id=u.idintestazione  WHERE secondario=0 and u.iduser="&term&""
			if session("idagente")<>"" and not ha_il_permesso("C5") then
				sql=sql&" and idagente="&session("idagente")
			end if

			set rs=conn.execute(sql)
			do while (not rs.eof) and i<max
				Set Js(Null) = jsObject()
				Js(null)("id")=rs("iduser")
				Js(null)("text")="Utente per ID: "&denominazione(rs("nome"),rs("cognome"),rs("azienda"))
				Js(null)("t")="U"
				rs.movenext
				i=i+1
			loop

			'Cerca articolo per ID
			sql="select * from prodotti where idpro="&term&""
			set rs=conn.execute(sql)
			do while (not rs.eof) and i<max
				Set Js(Null) = jsObject()
				Js(null)("id")=rs("idpro")
				Js(null)("text")="Articolo per ID: "&rs("codice")&" "&rs("articolo")
				Js(null)("t")="P"
				rs.movenext
				i=i+1
			loop
			'Fine isnumeric
		elseif mid(term,1,2)="d=" and len(term)>2 then
			salta=true
			'Cerco nei dipendenti
			term=mid(term,3)

			sql="select dipendenti.iddip, dipendenti.nome, dipendenti.cognome, utenti_intestazioni.Azienda, utenti_intestazioni.Cognome as utcognome, utenti_intestazioni.Nome as utnome, utenti.iduser FROM dipendenti inner join utenti_dipendenti on dipendenti.iddip=utenti_dipendenti.iddip inner join utenti on utenti_dipendenti.iduser=utenti.iduser inner join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id   WHERE "

				parole=split(term,",")
				where="INSTR( CONCAT_WS(  ' ', dipendenti.cognome, dipendenti.nome ) ,  '" &parole(0)& "' ) >0"
				for i = 1 to ubound(parole)
					if where<>"" then where=where&" AND "
					where="INSTR( CONCAT_WS(  ' ', dipendenti.cognome, dipendenti.nome ) ,  '" &parole(0)& "' ) >0"
					'where = where & " and instr( ricerca,'" &parole(i)& "')>0 "
				i=i+1
				next
				sql=sql&where

				session("sql")=sql
				set rs=conn.execute(sql)
				if not rs.eof then
					Set Js(Null) = jsObject()
					Js(null)("id")=-1
					Js(null)("text")="Corrisponenza in dipendenti:"
					Js(null)("disabled")=true
				end if
				i=0
				do while (not rs.eof) and i<max
					Set Js(Null) = jsObject()
					Js(null)("id")=rs("iduser")
					Js(null)("text")=rs("cognome")&" "&rs("nome")&" di "&denominazione(rs("utnome"),rs("utcognome"),rs("azienda"))
					Js(null)("t")="D"
					Js(null)("iddip")=rs("iddip")
					i=i+1
					rs.movenext
				loop




		else
			'Cerca utenti


			sql="select u.iduser,u.secondario, i.Azienda, u.Cognome, u.Nome, u.idagente FROM utenti_intestazioni i inner join utenti u on i.id=u.idintestazione  WHERE  "
			if session("idagente")<>"" and not ha_il_permesso("C5") then
				sql=sql&" idagente="&session("idagente") &" and "
			end if


			if isnumeric(term) then
				sql=sql&" i.iduser="&term
			else

				parole=split(term,",")
				where=" INSTR( CONCAT_WS(  ' ', i.azienda, u.nome, u.cognome ) ,  '" &trim(parole(0))& "' ) >0"
				for i = 1 to ubound(parole)
					if where<>"" then where=where&" AND "
					where=where&"INSTR( CONCAT_WS(  ' ', i.azienda, u.nome, u.cognome ) ,  '" &trim(parole(i))& "' ) >0"
					'where = where & " and instr( ricerca,'" &parole(i)& "')>0 "
				i=i+1
				next
				sql=sql&where
			end if
			sql=sql&" order by i.azienda"
			set rs=conn.execute(sql)
			if not rs.eof then
				Set Js(Null) = jsObject()
				Js(null)("id")=-1
				Js(null)("text")="Corrisponenza in utenti:"
				Js(null)("disabled")=true
			end if
			i=0
			do while (not rs.eof) and i<max
				Set Js(Null) = jsObject()
				Js(null)("id")=rs("iduser")
				if rs("secondario") then
									Js(null)("text")=rs("azienda")&" "&rs("nome")&" "&rs("cognome")&" secobdario Id:"&rs("iduser")

					else
										Js(null)("text")=rs("azienda")&" "&rs("nome")&" "&rs("cognome")&" Id:"&rs("iduser")

						end if
				Js(null)("t")="U"
				i=i+1
				rs.movenext
			loop
		end if
		if not salta then
			sql="select idpro, codice, articolo from prodotti "
			sql=sql&"where INSTR( CONCAT_WS(  ' ', codice, articolo ) ,  '" &term& "' ) >0"
			set rs=conn.execute(sql)
			if not rs.eof then
				Set Js(Null) = jsObject()
				Js(null)("id")=-1
				Js(null)("text")="Corrisponenza in articoli:"
				Js(null)("disabled")=true
			end if
			i=0
			do while (not rs.eof) and i<max
				Set Js(Null) = jsObject()
				Js(null)("id")=rs("idpro")
				Js(null)("text")=rs("codice")&" "&rs("articolo")
				Js(null)("t")="P"
				rs.movenext
				i=i+1
			loop
		end if

	case else
		add2log "Richiesta non riconosciuta"& queryeform()&"<br>"&Request.ServerVariables("HTTP_REFERER"),3
	end select
	js.Flush
	set js=Nothing
	set rs=nothing
	conn.close
	set conn=nothing
	response.end

end if

'funzioni autocomplete
term=replace(request("term"),"'","''")
autocomplete=request("autocomplete")
if autocomplete<>"" and term<>"" and session("idadmin")<>"" then
	Response.ContentType = "application/json; charset=utf-8"
	Response.CodePage = 65001
	i=0
	max=25
	strArr=""

	Set Js = jsArray()
	Set rs = Server.CreateObject("ADODB.Recordset")
	select case autocomplete
	case "fornitori"
		sql="select i.iduser, i.Azienda, i.Cognome, i.Nome FROM utenti_intestazioni i inner join utenti u on i.id=u.idintestazione  WHERE u.fornitore=1 and "


		if isnumeric(term) then
			sql=sql&" i.iduser="&term
		else

			parole=split(term," ")
			where="INSTR( CONCAT_WS(  ' ', i.azienda, i.nome, i.cognome ) ,  '" &parole(0)& "' ) >0"
			for i = 1 to ubound(parole)
				where="INSTR( CONCAT_WS(  ' ', i.azienda, i.nome, i.cognome ) ,  '" &parole(i)& "' ) >0"
				'where = where & " and instr( ricerca,'" &parole(i)& "')>0 "
			i=i+1
			next
			sql=sql&where
		end if
		sql=sql&" order by i.azienda"
		set rs=conn.execute(sql)
		If Not rs.EOF Then
			arrRS = rs.GetRows()
		end if
		Set rs = Nothing


		If IsArray(arrRS) Then
			maxtemp=UBound(arrRS, 2)
			if maxtemp>max then maxtemp=max
			For i = LBound(arrRS, 2) To maxtemp
				Set Js(Null) = jsObject()
				Js(null)("value")=arrRS(1, i)&" "&arrRS(3, i)&" "&arrRS(2, i)&" Id:"&arrRS(0, i)
				Js(null)("id")=arrRS(0, i)
			Next
			Erase arrRS
		End If
	case "utenti"
		sql="select u.iduser, i.Azienda, u.Cognome, u.Nome, u.idagente,u.secondario FROM utenti_intestazioni i inner join utenti u on i.id=u.idintestazione  WHERE  "
		if session("idagente")<>"" and not ha_il_permesso("C5") then
			sql=sql&" idagente="&session("idagente") &" and "
		end if


		if isnumeric(term) then
			sql=sql&" i.iduser="&term
		else

			parole=split(term,",")
			where="INSTR( CONCAT_WS(  ' ', i.azienda, u.nome, u.cognome ) ,  '" &trim(parole(0))& "' ) >0"
			for i = 1 to ubound(parole)
				if where<>"" then where=where&" AND "
				where=where&"INSTR( CONCAT_WS(  ' ', i.azienda, u.nome, u.cognome ) ,  '" &trim(parole(i))& "' ) >0"
				'where = where & " and instr( ricerca,'" &parole(i)& "')>0 "
			i=i+1
			next
			sql=sql&where
		end if
		sql=sql&" order by i.azienda"
		session("sql")=sql
		set rs=conn.execute(sql)
		If Not rs.EOF Then
			arrRS = rs.GetRows()
		end if
		Set rs = Nothing


		If IsArray(arrRS) Then
			maxtemp=UBound(arrRS, 2)
			if maxtemp>max then maxtemp=max
			For i = LBound(arrRS, 2) To maxtemp
				Set Js(Null) = jsObject()
				if arrRS(5, i) then
									Js(null)("value")=arrRS(1, i)&" "&arrRS(3, i)&" "&arrRS(2, i)&" secondario Id:"&arrRS(0, i)

					else
										Js(null)("value")=arrRS(1, i)&" "&arrRS(3, i)&" "&arrRS(2, i)&" Id:"&arrRS(0, i)

						end if
				Js(null)("id")=arrRS(0, i)
			Next
			Erase arrRS
		End If

	case else
		add2log "Richiesta non riconosciuta"& queryeform()&"<br>"&Request.ServerVariables("HTTP_REFERER"),3
	end select
	js.Flush
	set js=Nothing
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if

add2log "Richiesta non riconosciuta"& queryeform()&"<br>"&Request.ServerVariables("HTTP_REFERER"),3

call rsclose()
function cercaloop(tmp,idcercare,idsettore)
	'Eseguo scansione all'indietro
	if idcercare=0 then
		cercaloop=false
		exit function
	end if
	idsettore=idsettore&"|"&idcercare&"|"
	'declaring
	dim sql_order,str,MessageSpacing
	'create recordset
	'find child with thread_parent=his parent id
	sql_order = "select * FROM settori_settori WHERE idfiglio = " & idcercare
	set rs_order=conn.execute( sql_order)
	do while not rs_order.eof
		'response.write "cerco in :"&idsettore&" id:"&"|"&rs_order("idpadre")&"|<br>"
		if instr(idsettore,"|"&rs_order("idpadre")&"|")>0 then
			tmp=true
			'response.write "trovato<br>"
			exit do

		else
			tmp=cercaloop(layer,rs_order("idpadre"),idsettore)
		end if
		rs_order.MoveNext

	loop
	'closing object
	'rs_order.close
	set rs_order=nothing
	cercaloop=tmp
End function


%>
