<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="JSON_latest.asp"-->
<!--#include file="ClasseOrdine.asp"-->
<!--#include file="ClasseFattura.asp"-->
<!--#include file="ClasseModificheRS.asp"-->
<%
	documento=request("documento")
	if documento="" then
		response.write "manca documento"
		call add2log("Manca documento"&queryeform(),0)
		response.end
	end if
	idord=request("idord")
	iduser=request("iduser")
	duplicaidord=request("duplicaidord")
	esiste_fattura=false
	dim nome_campo()
	dim valore_campo()
	tab_datiFatturazione=false
	tab_datiGenerali=true

	tab_datiPa=true
	tab_consegna=true
	tab_ddt=false
	tab_fattura=false
	tabella=""
	testo=""
	gestisci_spettanze=true
	call imposta_documento(documento)
	
	select case documento
		case "ddt"
			tab_ddt=true
		case "fattura"
			tab_datiGenerali=false
			tab_fattura=true
			tab_datiFatturazione=true
			tab_datiPa=false
	end Select
	
	
	
	
	cancellabile=lcase(request("cancellabile"))
if request("salva_dati")="si" then
	if documento<>"fattura" then
		set ordine= (new ClasseOrdine)(array("apri",tabella,idord))
		ordine.crea_modifica_ordine("update")
		if request("dati-ddt")<>"" then
			ordine.salva_dati_ddt()
			
		end if
		set ordine= Nothing
	elseif documento="fattura" then
		set objDocumento = (new ClasseFattura)(array("apri","",idord,"totali"))
		objDocumento.modifica_fattura()
	
	else
		response.end
	end if
	
	
	
	Response.ContentType = "application/json; charset=utf-8"
	Set Js = jsObject()
	Js("status")="success"
	Js("message")="Dati salvati"
	session("test")= session("test") &"<br>"&js.jsString	
	js.Flush
	set js=nothing
	'conn.close
	'set conn=nothing
	
	response.end
end if


if idord<>"" and documento<>"fattura" then
	oper="edit"
	sql1="select "&tabella&".*, utenti.trattamento_iva, utenti.tipologia, utenti_intestazioni.* FROM "&tabella&" INNER JOIN utenti ON "&tabella&".iduser = utenti.iduser inner join utenti_intestazioni on "&tabella&".idintestazione = utenti_intestazioni.id where idord="&idord
elseif idord<>"" and documento="fattura" then
	idfat=idord
	sql1="select fatture.*,utenti_intestazioni.*,utenti.tipologia from fatture inner join utenti_intestazioni on fatture.idintestazione=utenti_intestazioni.id inner join utenti on fatture.iduser=utenti.iduser where idfat="&idfat
	oper="edit"
elseif duplicaidord<>"" then
	oper="duplica"
	sql1="select utenti.*,utenti_intestazioni.* FROM utenti inner join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id   where utenti.iduser="&iduser
elseif iduser<>"" then
	oper="new"
	sql1="select utenti.*, utenti_intestazioni.*, utenti_clienti.FatturaPACodiceDestinatario FROM utenti inner join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id left join utenti_clienti on utenti.iduser = utenti_clienti.iduser  where utenti.iduser="&iduser
else
	oper="new"
	
end if
if sql1<>"" then
	nuovo_utente=false
	response.write sql1
	Set rs1 = conn.execute(sql1)
	azienda=rs1("azienda")
	indirizzo=rs1("indirizzo")
	cap=rs1("cap")
	citta=rs1("citta")
	provincia=rs1("provincia")
	piva=rs1("piva")
	cf=rs1("cf")
	val_fattura_sp=rs1("fattura_sp")
	if documento<>"fattura" then
		tipopagamento=rs1("tipopagamento")
	elseif documento="fattura" and oper="edit" then
		tipopagamento=rs1("pagamento")
	elseif documento="fattura" and oper="new" then
		tipopagamento=rs1("tipopagamento")
	end if
	trattamentoiva=rs1("trattamento_iva")
	tipologia=rs1("tipologia")
	idintestazione=rs1("idintestazione")
	'idbanca=rs1("idbanca")
'	iban=rs1("iban")

	if oper="edit" then
		if tabella<>"ordini_fornitori" and tabella<>"fatture" then
			trattamentoivaordine=rs1("trattamento_iva_ordine")
		end if
		if documento<>"fattura" then
			tipotrasporto=rs1("tipo_trasporto")
			notespedizione=rs1("note_spedizione")
		end if
		idconsegna=rs1("idconsegna")
		iduser=rs1("iduser")
		
		titolo_dialog="Modifica dati "&documento&" di "&denominazione(rs1("nome"),rs1("cognome"),rs1("azienda"))
		if documento="ordine" or documento="preventivo" then
			presagente=rs1("idagente")
		end if

		
	elseif oper="new" then
		pag_accordato=rs1("Pag_accordato")
		tipotrasporto2=rs1("tipo_trasporto_2")
		notetrasporto2=rs1("note_trasporto_2")
		fornitore=rs1("fornitore")
		titolo_dialog="Nuovo "&documento&" a "&denominazione(rs1("nome"),rs1("cognome"),rs1("azienda"))

	end if
	titolo_dialog=replace(titolo_dialog,"'","")
else
	nuovo_utente=true
	
end if

if utente_andrea then
	response.write "[soloio]idintestazione:"&idintestazione&"nuovo_utente:"&nuovo_utente&"<br>"&queryeform()&"[/soloio]"
end if

if nuovo_utente or oper="edit"	then
	tab_datiFatturazione=true
end if

fattura=""
saldato=false
if oper="edit" then 
	
	if tabella="ordini" then
		'Verifico se ci sono fatture
		sql="select fatture.idfat, fatture.nfat, fatture.data, ordini_fatture.idord FROM fatture INNER JOIN ordini_fatture ON fatture.IDfat = ordini_fatture.idfat where  ordini_fatture.idord="&idord
		set rs=conn.execute (sql)
		if not rs.eof then
			fattura=rs("nfat")&" del "&formatdatetime(rs("data"),2)
		end if
		'Verifico se è saldato
		sql="select Sum(incassi.importo) AS SommaDiimporto FROM incassi WHERE (((incassi.idord)="&idord&"));"
		set rs=conn.execute (sql)
		if not isnull(rs("SommaDiimporto")) then
			if cdbl(rs("SommaDiimporto"))>=cdbl(rs1("totale")) then
				saldato=true
			end if
		end if
		set rs = nothing
	end if

 %>

<form id="form_intestazione">
<%else %>
<form  id="form_intestazione" action="pag_adm_user.asp" method="post">
<%end if%>
<input type="hidden" name="iduser" value="<%=iduser%>">
<input type="hidden" name="cheform" value="admin">
<%if oper="edit" then %>
<input type="hidden" name="idord" value="<%=idord%>">
<input type="hidden" name="nord" value="<%=nord%>">
<input type="hidden" name="documento" value="<%=documento%>">
<%end if %>
<input type="hidden" name="duplicaidord" value="<%=duplicaidord%>">
<%if nuovo_utente then %>
<input type="hidden" name="nuovo_utente" value="si">
<%end if%>
<%if oper="new" and documento="ordine" and nuovo_utente= false then
set preventivi=conn.execute("select preventivi.* from preventivi where eliminato=0 and iduser="&iduser&" order by data")
if not preventivi.eof then %>
<div class="ui-widget-header ui-corner-all" style="padding:3px;">Copia articoli da un preventivo</div>
	<table cellpadding="2" border="0" class="tabella1" width="100%">
	<tr bgcolor="#E5E5E5">
	<td></td>
	<td>N°</td>
	<td>Data</td>
	<td align="right">Importo</td>
	<td>Stato</td>
</tr>
<%
	do while not preventivi.eof
	%>
	<tr><td>
	<input type="radio" name="idpreventivo" value="<%=preventivi("idord")%>">
	</td>
	<td><a href="pag_adm_preventivi.asp?idord=<%=preventivi("idord")%>" target="_blank"><%=preventivi("nord")%></a></td>
	 <td><%=preventivi("data")%></td>
	 <td align="right"> <%=formatnumber(preventivi("totale"),2)%></td>
	 <td><%=stato_preventivo(preventivi("stato"))%></td>
	</tr>
	<%
	preventivi.movenext
loop
%>
	<tr><td>
	<input type="radio" name="idpreventivo" value="">
	</td>
	<td colspan="4">Nessuno, ordine senza articoli</td>
	</tr>



</table>

<%
end if
end if
Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class


 %>




<div id="edit-ordine">
  <ul>
	<%if tab_datiFatturazione then %>
    <li><a href="#edit-ordine-1">Dati fatturazione</a></li>
	<%end if %>
    <li><a href="#dati-consegna">Dati consegna</a></li>
	<%if tab_datiGenerali then %>
    <li><a href="#datiGenerali">Dati generali</a></li>
	<%end if %>
    <% if tab_fattura then %>
    <li><a href="#dati-fattura">Dati fattura</a></li>
    
    <%end if %>
    <%if tab_ddt then %>
    <li><a href="#dati-ddt">Dati DDT</a></li>
    <%end if %>
	<%
	if tab_datiPa then
	%>
    <li><a href="#dati-pa">Dati PA</a></li>
    <%end if %>
  </ul>
  
  
  <%if tab_datiFatturazione then %>
  <div id="edit-ordine-1">
		<strong>Azienda</strong>:<br>
		<input name="azienda" id="azienda" type="text" value="<%=azienda%>" style="width:98%;">
		<%if nuovo_utente then %>
		<strong>Cognome</strong>:<br>
		<input name="cognome" type="text" value="" style="width:98%;">
		<strong>Nome</strong>:<br>
		<input name="nome" type="text" value="" style="width:98%;">
		<strong>Telefono</strong>:<br>
		<input name="telefono" type="text" value="" style="width:98%;">
		<strong>Email</strong>:<br>
		<input name="email" type="text" value="" style="width:98%;">
		<%end if %>
		<br>
		<strong>Indirizzo</strong>:<br>
		<input name="indirizzo" type="text" value="<%=indirizzo%>"  style="width:98%;">
		<br>
		<strong>Citt&agrave;</strong>:<br>
		<input name="citta" id="citta" type="text" value="<%=citta%>"  style="width:98%;">
		<br>
		<strong>Cap</strong>:<br>
		<input name="cap" id="cap" type="text" value="<%=cap%>"  style="width:98%;">
		<br>
		<strong>Provincia</strong>:<br>
		<input name="provincia" id="provincia" type="text"  style="width:98%;" value="<%=provincia%>" maxlength="2">
		<br>
		<strong>Partita Iva</strong>:<br>
		<input name="piva" type="text" id="piva"  style="width:98%;" value="<%=piva%>">
		<br>
		<strong>Codice Fiscale</strong>:<br>
		<input name="cf" type="text" id="cf"  style="width:98%;" value="<%=cf%>">
		<br>
				
	</div>
	<%end if %>
	<%if tab_datiGenerali then %>

	<div id="datiGenerali">
		<!-- DATI GENRALI -->
		
		<%
			visualizza_ordine=false
			if documento="sostituzione" and oper="new" then
				visualizza_ordine=true
			elseif documento="sostituzione" and oper="edit" then
				set rs_tmp=conn.execute("select ordini_sostituzioni.* from ordini_sostituzioni where idsostituzione="&idord)
				if rs_tmp.eof then visualizza_ordine=true
				
			end if
		if visualizza_ordine then %>
		
		
		<label><strong>Riferimento ordine</strong></label>
		<input  name="seleziona_ordine" id="seleziona_ordine" value="" style="width: 300px;"/><br>
		<script>
			
		$(document).ready(function () {
		$('#seleziona_ordine').select2({placeholder: '',minimumInputLength: 0,width: 'resolve',ajax: {quietMillis: 150,url: 'ajax_function.asp?select2=ordine_sostituzione&iduser=<%=iduser%>',dataType: 'json',data: function (term, page) {return {term: term};},results: function (data) {return {results: data};}}});
		});
		</script>
		<%end if %>		
	  	<%if request.querystring("carrello")<>"" then %>

		<label><input type="checkbox" name="carrello_utente" id="carrello_utente" value="SI" checked > Usa carrello utente</label>
		&nbsp;<label><input type="checkbox" name="conserva_carrello" id="conserva_carrello" value="SI" checked >Lascia gli articoli nel carrello utente</label>

		<%end if
		if oper="new" and month(date())=1 then
			anno=year(date())
			annoprecedente=anno-1
		%>
		<label><strong>Anno</strong></label><input type="radio" name="anno" value="" checked><%=anno%>		<input type="radio" name="anno" value="<%=annoprecedente%>"><%=annoprecedente%><br>
		<%
		end if
		%>
                <strong>Tipo pagamento</strong>:<br>
                <%
				if oper="new" then
					txt=pag_accordato
				else
					txt=tipopagamento
				end if
	if isnull(txt) then txt=0
%>
                <select name="tipopagamento" >
                  <option value="0" <%if txt=0 then response.write " selected"%>>Non specificato</option>
                <%call metodipagamento.stampa_option(txt)%>
                </select>
                <br />
				<%if tabella<>"ordini_fornitori" then %>
				<strong>Trattamento IVA</strong>:<br>
				<%
				modificabile=true
				testo=""
				if oper="edit" then
					val=trattamentoivaordine
					if fattura<>"" then
						modificabile=false
						testo="<br>Non modificabile perch&egrave; gi&agrave; emessa fattura "&fattura
					end if
					if saldato then
						modificabile=false
						testo=testo& "<br>Non modificabile perch&egrave; l'ordine &egrave; gi&agrave; saldato"
					end if
					if modificabile=false then
						response.write trattamento_iva(val)&testo
					end if
				else
					val=trattamentoiva
				end if
				if modificabile then
				%>
              <select name="trattamento_iva_ordine" id="trattamentoiva">
				<%for n= 0 to trattamento_iva(-1) %>
                <option value="<%=n%>" <%if val=n then response.write " selected"%> ><%=trattamento_iva(n)%></option>
				<%next %>
				</select>
				<%
				if val=10 then
					val=" checked"
				else
					val=""
				end if
				%>
				PA <input name="" type="checkbox" value="si" id="pa"  <%response.write val%>> &nbsp;&nbsp;&nbsp;  Split Payment <input name="fattura_sp" type="checkbox" value="si"   <%response.write val_fattura_sp%>>
				
				<%end if%><br>
				<%end if %>
				
				<strong>Metodo spedizione</strong>:<br>
				<%
					if oper="edit" then
						val=tipotrasporto
					else
						val=tipotrasporto2
					end if
					if val <>"" or not isnull(val) then val=cint(val) else val=0
				%>
          <select name="trasporto">
            <%
			for n=1 to tipo_trasporto(-1)
			response.write "<option value='"&n&"'"
			if val=n then response.write " selected "
			response.write ">" & tipo_trasporto(n)&"</option>"
			next
			%>
          </select><br/>
				<%
					val=""
					if oper="edit" then
						val=notespedizione
					else
						if tipotrasporto2=4 then
							val=notetrasporto2
						end if
					end if
				%>
          
          
          <strong>Note spedizione:</strong><br><input name="note-trasporto" type="text" value="<%=val%>" size="40" ><br/>
            				
            				

        <strong>Tipologia utente</strong> <br/> <%
					val=","&tipologia&","
				%>
          <select name="tipologia"  id="tipologia2" multiple style="width:400px;" <%=disabled%>>
  <%
set rsm=conn.execute("select * from utenti_tipologia order by tipologia;")
do while not rsm.eof
%>
            <option value="<%=rsm("id")%>" <%if instr(val,","&rsm("id")&",")>0 then response.write " selected"%>><%=rsm("tipologia")%></option>
            <%
rsm.movenext
loop
rsm.close
%>
          </select>
          
          
          
          
          
          <br/>
          <%if (documento="ordine" or documento="preventivo") and ha_il_permesso("Z1") then %>
        <strong>Agente</strong> <br/> <%
				%>
          <select name="idagente"  id="idagente" >
	          <option value="0">Nessuno</option>
  <%
set rsm=conn.execute("select * from agenti order by nominativo;")
do while not rsm.eof
%>
            <option value="<%=rsm("idagente")%>" <%if presagente=clng(rsm("idagente")) then response.write " selected"%>><%=rsm("nominativo")%></option>
            <%
rsm.movenext
loop
rsm.close
%>
          </select>
          
          
          
          
          
          <br/>
          <%end if %>
          
          
          
          <%if oper="edit" then %>
                Salva questi dati nel profilo utente:
                <input type="checkbox" name="salva_anagrafica" id="salva" value="SI">
                <br/>
           <%end if
	           if documento="ordine" then
	                val=""
					if oper="edit" then
						
					if rs1("no_magazzino") then val=" checked"
					end if
						
				
				%>

                Non scaricare il magazzino:
                <input type="checkbox" name="no_magazzino" value="si" <%=val%>>
                <%end if %>
                <%if documento="ordine" or documento="preventivo" then %>
                <br>
                <%
	                val=""
					if oper="edit" then
						
					if rs1("spettanze")=1 then val=" checked"
					end if
	                %>
                Spettanze: <input type="checkbox" name="spettanze" value="1" <%=val%>>
                
                <%end if
	           if documento="preventivo" then
	                val=""
					if oper="edit" then
						
					if rs1("nascondi_totali")=1 then val=" checked"
					end if
						
				
				%>

                Nascondi totali:
                <input type="checkbox" name="nascondi_totali" value="si" <%=val%>>
                <%end if%>
  </div>
  <%end if 'if tab_datiGenerali then %>
  
  <%
	call tab_dati_consegna(idconsegna,iduser)  
	if tab_ddt then 
	    idddt=0
	    if idord<>"" then idddt=idord
	    call tab_dati_ddt(idord,iduser) 
    end if 
    if tab_fattura then
	    
	    
	    
    	if oper="edit" then
    		set rs_fattura=conn.execute("select * from fatture where idfat="&idfat)
    		data=rs_fattura("data")
    		nfat=rs_fattura("nfat")
    		causale=rs_fattura("causale")
    		fine_mese=rs_fattura("fine_mese")
    		porto=rs_fattura("porto")
    		imballo=rs_fattura("imballo")
    		colli=rs_fattura("colli")
    		peso=rs_fattura("peso")
    		dimensione=rs_fattura("dimensione")
    		incaricato=rs_fattura("incaricato")
    		data_inizio=rs_fattura("data_inizio")
    		annotazioni=rs_fattura("annotazioni")
    		pagamento=rs_fattura("pagamento")
    		spese_bancarie=rs_fattura("spese_bancarie")
    		trattamento_iva_v=rs_fattura("trattamento_iva")
    		CodiceDestinatario=rs_fattura("CodiceDestinatario")
    		idbanca=rs_fattura("idbanca")
    		banca_appoggiov=rs_fattura("banca_appoggio")
    		iban=rs_fattura("iban")
    		fattura_sp=rs_fattura("fattura_sp")
    	else
	    	data=formatDateTime(date(), vbShortDate)
	    	nfat=""
	    	causale=""
    		fine_mese=0
    		porto=0
    		imballo=""
    		colli=""
    		peso=""
    		dimensione=""
    		incaricato=""
    		data_inizio=formatDateTime(date(), vbShortDate)
    		annotazioni=""
    		pagamento=tipopagamento
    		spese_bancarie=""
    		trattamento_iva_v=trattamentoiva
    		banca_appoggiov=""
	    	CodiceDestinatario=""
	    	fattura_sp=""
	    	
	    end if
    	
    	
    
    
    
     %>
   	<div id="dati-fattura">
 
	<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr>
              <td colspan="2" valign="top" class="ui-widget-header">Dati fattura  
              <%if oper="view" and ha_il_permesso("B1") then%><input name="edit" type="submit"  style="FONT: 12px;" value="Modifica" class="modifica-da-tab" data-id="<%=idfat%>" data-url="<%=questofile%>?tab=1" data-nomeid="idfat" data-oper="edit" data-form="form-fatt">
                <%end if%>
              
              </td>
            </tr>
			
			<tr>
              <td valign="top">Numero</td>
              <td>
			  <input type="text" name="new_nfat" id="new_nfat" value="<%=nfat%>">
			  </td>
            </tr>
			<tr>
              <td valign="top">Data</td>
              <td><%
				modificadata=true
				if oper="edit" then
					modificadata=false
					sql="select fatture.data, fatture.nfat, fatture.pa from fatture where idfat="&idfat
					set rs_fattura=conn.execute(sql)
					
					sql="select fatture.* from fatture where year(data)="&year(rs_fattura("data"))&" and nfat="&rs_fattura("nfat")-1&" and pa="&rs_fattura("pa")
					'response.write sql
					set rs_data=conn.execute (sql)
					if not rs_data.eof then
						data_minima=rs_data("data")
						'response.write "trovato data minima"&rs_data("data")&" di nfat:"&rs_data("nfat")&" di idfat:"&rs_data("idfat")
						if data_minima<rs_fattura("data") then
							modificadata=true
							
						end if
					end if
					modificadata=true
				else
					
				end if
				if modificadata then
				  %>
				  <input type="text" name="data" value="<%=data%>"> 
				  <input type="hidden" name="data_minima" value="<%=data_minima%>">
				<%else
					response.write "Non modificabile"
				end if %>
				</td>
            </tr>
			
            <tr>
              <td valign="top">Causale</td>
              <td><%
	              Set ccausale_ddt = new cl_causale_ddt
				%>
                <select name="causale">
					<%
	                  ccausale_ddt.stampa_option(causale)
					%>
                </select>
              <%
	            set ccausale_ddt = Nothing
	              
              %></td>
            </tr>

            <tr>
              <td valign="top">Pagamento</td>
              <td>


               <select name="pagamento">
                <%call metodipagamento.stampa_option(pagamento)%>
                </select>
                </td>
            </tr>
            <tr>
              <td valign="top">Spese bancarie</td>
              <td><%
					if spese_bancariev>0 then
					  val=formatnumber(rs1("spese_bancarie"),2)
				  else
					  val=formatnumber(0,2)
				  end if
			  
			  
			  
		
				
				%>
                <input name="spese_bancarie" type="text" value="<%=spese_bancarie%>">

                </td>
            </tr>
            <tr>
              <td valign="top">Trattamento IVA</td>
              <td><%
	              
	              
	              
	              
              val=trattamento_iva_v
			if oper="edit" then
					modificabile=true
					saldato=false
					'Verifico se è saldata
					sql="select Sum(incassi.importo) AS SommaDiimporto FROM incassi WHERE (((incassi.idfat)="&idfat&"));"
					set rs=conn.execute (sql)
					if not isnull(rs("SommaDiimporto")) then
						if cdbl(rs("SommaDiimporto"))>=cdbl(rs1("totale_fattura")) then
							saldato=true
						end if
					end if
					set rs = nothing
					if saldato then
						modificabile=false
						testo=testo& "<br>Non modificabile perch&egrave; la fattura &egrave; gi&agrave; saldata"
					end if
					if modificabile=false then
						response.write trattamento_iva(val)&testo
					end if
			end if
	              
				if modificabile then
              
	              %>
              <select name="trattamento_iva" <%=disabled%>>
				<%for n= 0 to trattamento_iva(-1) %>
                <option value="<%=n%>" <%if val=n then response.write " selected"%> ><%=trattamento_iva(n)%></option>
				<%next %>
				</select>
				<%end if 'if modificabile
				%>
              </td>
            </tr>
            <%if true then %>
            <tr>
              <td valign="top">Codice univoco</td>
              <td>
	              
	         <%
		         
		         set rs_tmp=conn.execute("select FatturaPACodiceDestinatario from utenti_clienti where iduser="&iduser)
		         if not rs_tmp.eof then
			         FatturaPACodiceDestinatario=rs_tmp("FatturaPACodiceDestinatario")
			     end if
			     set rs_tmp = Nothing
		              if not isnull(FatturaPACodiceDestinatario) and FatturaPACodiceDestinatario<>"" then
			            
		                tmp=split(FatturaPACodiceDestinatario,vbcrlf)

	              %>
	              
	              
                 <select name="CodiceDestinatario">
                  <%
	                  
				for n=0 to ubound(tmp)
					
					if instr(tmp(n),":")>0 then 'Ci sono i 2 punti
								tmp2=split(tmp(n),":")
								riga=ucase(trim(tmp2(0)))&" "&trim(tmp2(1))
								codice=tmp2(0)
							else 						'Non ci sono i 2 punti
								riga=ucase(tmp(n))
								codice=riga
							end if
							
					response.write "<option value='"&codice&"'"
					if CodiceDestinatario=codice then response.write " selected style='color:red;'"
					response.write ">" & riga&"</option>"
				next
				%>
                </select>
                <%
	            end if%></td>
            </tr>
            
            <%end if%>
            <tr>
              <td valign="top">Banca appoggio cliente</td>
              <td>
                <input name="banca_appoggio" type="text" value="<%=banca_appoggiov%>" style="width:90%;">
                </td>
            </tr>
            <tr>
              <td valign="top">IBAN cliente</td>
              <td>
                <input name="iban" type="text" value="<%=iban%>" style="width:90%;">
                
                </td>
            </tr>
            
			
            <tr>
              <td valign="top">Nostra banca</td>
              <td>
              <select name="idbanca"  <%=disabled%>>
	              <option value="1" <%if val="" then response.write " selected"%>>Non impostata</option>

                  <%
					set rsm=conn.execute("select * from banche order by nome_breve;")
					do while not rsm.eof
					%>
					<option value="<%=rsm("idbanca")%>" <%if idbanca=rsm("idbanca") then response.write " selected"%>><%=rsm("nome_breve")%></option>
					<%
					rsm.movenext
					loop
					rsm.close
					set rsm=Nothing
					%>
                </select></td>
            </tr>
        <tr>
              <td valign="top">Accompagnatoria</td>
              <td> <%val=""
	
				if fine_mese=0 then val=" checked"
%>
          <input type="checkbox" name="accompagnatoria" id="accompagnatoria" value="si"  <%response.write val%> <%=disabled%>></td>
            </tr>
        <tr>
              <td valign="top">Split Payment</td>
              <td> <%val=""
	
				if fattura_sp=1 then val=" checked"
%>
          <input type="checkbox" name="fattura_sp"  value="si"  <%response.write val%> <%=disabled%>></td>
            </tr>

            
            
            <tr class="accompagnatoria">
              <td valign="top">Porto</td>
              <td>
                <select name="porto">
                  <%
				for n=1 to porto_ddt(0)
					response.write "<option value='"&n&"'"
					if porto=n then response.write " selected style='color:red;'"
					response.write ">" & porto_ddt(n)&"</option>"
				next
				%>
                </select>
              </td>
            </tr>
            <tr class="accompagnatoria">
              <td valign="top">Imballo</td>
              <td><%if oper="view" then%>
                <%=imballo_ddt(rs1("imballo"))%>
                <%else%>
                <select name="imballo">
                  <%
for n=1 to imballo_ddt(0)
	response.write "<option value='"&n&"'"
	if imballo=n then response.write " selected style='color:red;'"
	response.write ">" & imballo_ddt(n)&"</option>"
next
%>
                </select>
              <%end if%></td>
            </tr>
            <tr class="accompagnatoria">
              <td valign="top">Numero colli</td>
              <td>
                <input name="colli" type="text" value="<%=colli%>">
              </td>
            </tr>
            <tr class="accompagnatoria">
              <td valign="top">Peso</td>
              <td>
				<input name="peso" type="text" value="<%=peso%>">
              </td>
            </tr>
            <tr class="accompagnatoria">
              <td valign="top">Dimensione</td>
              <td>
                
	                
                <input name="dimensione" type="text" value="<%=dimensione%>">
              </td>
            </tr>
           
            <tr class="accompagnatoria">
              <td valign="top">Incaricato del trasporto</td>
              <td>

              <select name="incaricato" class="richiesto">
                  <%
					for n=1 to incaricato_ddt(0)
					%>
					<option value='<%=n%>' <%if n=incaricato then%>selected<%end if%>> <%= incaricato_ddt(n)%></option>
					<%
					next
					%>
                </select>
              
              </td>
            </tr>
			<%
			'val=""
			'if rs1("data_inizio")<>"" then val=formatdatetime(rs1("data_inizio"),2)
			%>
            <tr class="accompagnatoria">
              <td valign="top">Data inizio trasporto</td>
              <td>     
                <input name="data_inizio" type="text" id="data_inizio" value="<%=data_inizio%>">
              </td>
            </tr>
            <tr class="accompagnatoria">
              <td valign="top">Annotazioni</td>
              <td>
				<input name="annotazioni" type="text" value="<%=annotazioni%>" style="width:90%;">
                </td>
            </tr>
          </table>
			<script>
			$(function() {
				
				$('#accompagnatoria').is(':checked')?$(".accompagnatoria").show():$(".accompagnatoria").hide();
				$("#accompagnatoria").click(function(){
					$('#accompagnatoria').is(':checked')?$(".accompagnatoria").show():$(".accompagnatoria").hide();
				});
				
				
				
				
    			<%if oper="edit" then%>

				
				
				
				
				var nfat=$("#nfat").val();
				$( "#new_nfat" ).blur(function(){
					var new_nfat=$(this).val();
					var idfat=$("#idfat").val();
					if(nfat!=new_nfat){
						//alert(nfat+ " diverso da "+new_nfat);
						$.ajax({
							url     : "ajax_function.asp?oper=verifica_new_nfat",
							type    : "post",
							//dataType: 'html',
							dataType: 'json',
							cache: false,
							data	: {new_nfat: new_nfat,
										idfat: idfat},
							success: function(data){
							if(data.error==true){
									$("#dialog").html(data.message);
									var dialog=$("#dialog").dialog({
										autoOpen: true,
										modal: true,
										resizable: false,
										width: "auto",
										height: "auto",
										title: "Variazione numero fattura"
									}); 
							
								//alert(data.message);
								}

							}
							,error: function(xhr, textStatus, error){
								toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
								toastr.error('Errore nel caricamento della pagina');
								console.log("xhr.statusText:"+xhr.statusText);
								console.log("xhr.responseText:"+xhr.responseText);
								console.log("textStatus:"+textStatus);
								console.log("error:"+error);
							}
						});	
					}
				});
			
            <%end if ''oper="edit" %>
            
            
			});			
            
			</script>
    </div>
    
    <%
	end if
    
    if tab_datiPa  then
  %>
				  <div id="dati-pa">
	  
	              				<%
	            				
					if oper="new" then
					
						cig=""
						mepa=""
						impegno_spesa=""
					else
						if tabella="ordini" then
							cig=rs1("cig")
							mepa_tipo_v=rs1("mepa_tipo")
							mepa_testo=rs1("mepa_testo")
							impegno_spesa=rs1("impegno_spesa")
							mepa_data=rs1("mepa_data")
						end if
					end if
				%>
                <strong>Codice CIG</strong>:<br>
                <input name="cig" type="text"   style="width:98%;" value="<%=cig%>" maxlength="12">
                <br>
                <strong>Riferimento</strong>
                <select name="mepa_tipo">
				<%for n= 0 to mepa_tipo(-1) %>
                <option value="<%=n%>" <%if mepa_tipo_v=n then response.write " selected"%> ><%=mepa_tipo(n)%></option>
				<%next %>
				</select>:<br>
                
                <input name="mepa_testo" type="text"   style="width:98%;" value="<%=mepa_testo%>" maxlength="50">
                
                <strong>Data ordine</strong>:<br>
                <input name="mepa_data" type="text" id="mepa_data" value="<%=mepa_data%>"><br>
                
                <strong>Impegno di spesa</strong>:<br>
                <input name="impegno_spesa" type="text"   style="width:98%;" value="<%=impegno_spesa%>" >
                

	  

  </div>
<%end if%>  
</div>

			<%if cancellabile="false" then%>
              <span style="text-align:left; color:red;"><strong>Attenzione:
              DDt o Fattura gi&agrave; emessi, modificando questi dati vengono modificati anche i dati di intestazione di DDt o Fattura legati a questo ordine</strong></span><br/>
              <%end if%>
            <div class="ui-dialog-buttonset" style="text-align:center;">
              <%if oper="edit" then %>
                <input name="modifica_intestazione" type="button" id="modifica_intestazione" value="Salva modifiche" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only">
                <%else 'new o duplica()
	                if oper="new" then
	                	testoBtn="Crea"
	                else
		                testoBtn="Duplica"
		            end if
		            if tabella="preventivi" then
	            		testoBtn=testoBtn&" preventivo"
	            		nomebtn="crea_preventivo"
	            	elseif tabella="ordini_fornitori" then
	            		testoBtn=testoBtn&" ordine fornitore"
	            		nomebtn="crea_ordine_fornitore"
					else
		            select case documento
		            	case "notacredito"
		            		testoBtn=testoBtn&" nota di credito"
			            	nomebtn="crea_"&documento
		            	case else
		            	testoBtn=testoBtn&" "&documento
		            	nomebtn="crea_"&documento
		            end select

	            	end if
		            
		            	
		            
		            
		            
		            
		            
		            
	                %>
                <input type="hidden" name="dialog" value="si">
                <%if oper<>"edit" then %>
                <input id="nuovo_<%=documento%>" type="button" value="Indietro" class="nuovodocumento ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only" style="float: left;">
				<%end if %>
				
                <input name="<%=nomebtn%>" type="submit" value="<%=testoBtn%>" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only">
                <%end if%>
                 </div>
              </form>
	<script src="jquery/js/jquery.validate.min.js" type="text/javascript"></script>
	<script src="jquery/js/jquery.validate.min_it.js" type="text/javascript"></script>
	<script src="jquery/ui/i18n/jquery.ui.datepicker-it.min.js"></script>
              <%
	            set rs1=Nothing
	            set rs = Nothing
	            set rsm = Nothing
	            call connclose()
	              
	              %>

<script>







$(document).on("click","#modifica_intestazione",function(e) {
	
	
		e.preventDefault();

		var dati="salva_dati=si&"+$( "#form_intestazione" ).serialize();
	        $.ajax({
            url: "pag_adm_dialog_ordine.asp",
            type: "post",
            cache: false,
            dataType: 'json',
            data	: dati,
            success: function(data) {
	            console.log("OK");
				RicaricaDocumento();
            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});

                var txt = "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                invia_errore("Errore in pulsante_modifica_intestazione()", txt);
            }
        });

});
	
$(function() {	
	
	
	
		$("#form_intestazione").validate({
			ignore: "",
			rules: {
				<%if tab_fattura then %>
				porto: {
					required: true
				},
				imballo: {required: true},
				colli: {
					required: {
					depends: function(element){
					   return $('#accompagnatoria').is(':checked');
						}
					}

				},
				peso: {
					required: {
					depends: function(element){
					   return $('#accompagnatoria').is(':checked');
						}
					}
				},
				incaricato: {
					required: {
					depends: function(element){
					   return $('#accompagnatoria').is(':checked');
						}
					}
				},

				pagamento:  {
					required: {
					depends: function(element){
					   return $('#accompagnatoria').is(':checked');
						}
					}
				}
	
				<%end if %>		
				
				} //rules
		});
	
	
	
	$('#dialog').dialog('option', 'title', '<%=titolo_dialog%>');
	
		$("#pa").change(function () {
			if ($(this).prop("checked")){
				$("#trattamentoiva").val(10);
				
				
			}else{
				$("#trattamentoiva").val(0);
			}
		});
	
	
	
	
	$("#mepa_data" ).datepicker();
	
	
	form=  $('#tipologia2');
	$(form).select2({
	formatNoMatches:function(term){
			termine=term;
            return "Voce non in elenco, <button onclick='return aggiungi();'>aggiungi "+term+"  </button>";
		}}
	);
	$( "#edit-ordine" ).tabs();

});
function RicaricaDocumento() {
	
	    if (typeof tipo_documento !== 'undefined') {
		switch (tipo_documento){
			case 'ddt':
				location.href="pag_adm_ddt.asp?idddt="+idDocumento;
				break;
			case 'ordine':
				location.href="pag_adm_ordini.asp?idord="+idDocumento;
				break;
			case 'notacredito':
				location.href="pag_adm_note_credito.asp?idord="+idDocumento;
				break;
			case 'fattura':
				location.href="pag_adm_fatture.asp?idfat="+idDocumento;
				break;
			default:
				console.log("non so dove andare, ricarico la stessa pagina");
				location.href=location.pathname+"?idord="+idDocumento;
				
		}
		}else{
			console.log("non so dove andare, ricarico la stessa pagina");
			location.href=location.pathname+"?idord="+idDocumento;
			//location.href="pag_adm_"+tabella+".asp?idord="+idDocumento;
		}
	
}
$( "#citta" ).autocomplete({
		source: "searcher_comuni.asp",
		cache: false ,
		minLength: 3,
		select: function( event, ui ) {
			$("#provincia").val(ui.item.targa);
			$("#cap").val(ui.item.cap);
		}
	});

</script>
<%
	call CheckConnChiusa()
	%>