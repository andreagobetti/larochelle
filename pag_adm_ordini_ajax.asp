<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="ClasseOrdine.asp"-->
<!--#include file="ClasseFattura.asp"-->
<!--#include file="JSON_latest.asp"-->
<!--#include virtual="/classeElencoArticoli.asp" -->
<%
idord=Request("idord")
nord=Request("nord")
dove=Request("dove")
oper=lcase(request("oper"))
documento=request("documento")
select case documento
	case "ordine"
		tabella="ordini"
		testo="ordine"
	case "preventivo"
		tabella="preventivi"
		testo="preventivo"
	case "ddt"
		tabella="ordini"
		testo="ddt"
	case "ordine_fornitore"
		tabella="ordini_fornitori"
		testo="ordine fornitore"
	case "notacredito"
		tabella="ordini"
		testo="nota di credito"
	case "sostituzione"
		tabella="ordini"
		testo="sostituzione"
	case "fattura"
		tabella="fatture"
		testo="fattura"
	case ""
		tabella=request("tabella")
end select
		

str_add2log="<b>Querystring</b>:"
aggiornato=true
if tabella="" then 
	add2log "Errore, manca dove o errato:"&str_add2log,3
	response.write "Errore, contesto errato o non specificato" 
	response.end
end if
if tabella<>"" then
	if oper="riordina" then
		getItems = Request.QueryString("id[]")
		'split them junks into array
		getItems = Split(getItems, ",")
		'For each one, define and then execute the sql statement
		i_max=UBound(getItems)
		i_max=i_max*2
		For i = LBound(getiTems) TO i_max step 2
		   sql_update = "UPDATE "&tabella&"_dett SET ordine = " & i & " WHERE iddett = " & getItems(i/2)
		   'execute this bitch
		   conn.Execute sql_update
		   'below used for debugging - comment out
		   'Response.Write("ID " & getItems(i) & " was set to position " & i)
			'response.write "."
		Next
		response.write  (i/2)&" articoli ordinati correttamente"

	elseif oper="salva_nota" or oper="modifica_nota" then
		'iddett=cint(request.querystring("iddett"))
		sql="select * from "&tabella&" where idord="&idord
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		nota=trim(request.form("nota"))
		session("test_nota")=nota
		nota=replace(nota,"</p><p>","<br>")
		nota=replace(nota,"<p>","")
		nota=replace(nota,"</p>","")
		if oper="modifica_nota" then modifica_nota=" (nota modificata)"
		nota= now()&" <b>"&session("nominativo")&"</b>"&modifica_nota&":<br>"&nota
		if isnull(rs("note_gestore")) or len(trim(rs("note_gestore")))=0 or oper="modifica_nota"  then
			rs("note_gestore")=nota
		else
			rs("note_gestore")=rs("note_gestore")&"<hr />"& nota
		end if
		nota=rs("note_gestore")
		rs.update
		rs.close
		set rs=nothing
		response.write nota
	elseif oper="delete_file" then
		call elimina_allegato(request("id_file"))
		response.write "Allegato eliminato"
	elseif oper="registro" then
		cosa=request.querystring("cosa")
		id=request.querystring("id")
		if cosa<>"" and id<>"" then
			sql="select * from log " 
			sql=sql & " where instr([evento],'["&cosa&"="& id&";')"
			sql=sql&" order by idlog desc"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql, conn, 3, 3
	%>
            <div align="left" style="border: 1px solid gray; background: #E5E5E5; padding-top: 5px; padding-bottom:5px;">
			<%
			do while not rs.eof
			response.write rs("data")&" "
			response.write rs("causale")&" "
			response.write rs("evento")&"<br>"
			rs.movenext
			loop
			%>
            </div>
		  <%
			rs.close
			set rs=nothing
		  end if
	elseif oper="cronologia" then
		cronologia=conn.execute ("select comunicazioni_precedenti from "&tabella&" where idord="&idord)(0)
		%>
		        <div class="ui-widget-header">Cronologia</div>
		
		<%
		response.write cronologia
		
		call connclose()
		response.end
		
	elseif oper="stato_articoli" then
		
	
		set ordine= (new ClasseOrdine)(array("imposta_stato_articoli",tabella,idord))
	
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		
		js("success")=true
		js("message")="Numero seriale dinuovo disponibile"
		set rs = Nothing
		js.Flush
		set js=Nothing
	
	
	
	
	end if
		
	
	if instr(oper,"tbody")>0  then
			if documento="fattura" then	
				set ordine = (new ClasseFattura)(array("apri","",idord,"totali"))
		
			else
				set ordine= (new ClasseOrdine)(array("apri",tabella,idord,"totali"))
				
					
			end if
		
		
		
			'Aggiungo articolo
			if request.form("aggiungiarticolo")="si" then
				n=ordine.aggiungi(request("aggiungo"))
				
			end if
			
			ordine.modifica_articoli=true
			call ordine.elenco_articoli()
			if tabella<>"ddt" then
				ordine.elenco_totali()
			end if
			set ordine = Nothing
		
		
		
	end if
	
end if	
call connclose()
call CheckConnChiusa()
%>
