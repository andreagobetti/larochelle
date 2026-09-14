<%
'Verifica chiusure 02_12_2015
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="JSON_latest.asp"-->

<%
idfat=request("idfat")
if idfat="" then idfat=0
idord=request("idord")
if idord="" then idord=0
nord=request("nord")
idincasso=request("idincasso")
ricarica=request("ricarica")
idscadenza=request("idscadenza")
'Preset
causale=0
importo=""
modoincasso=""

tipo_pagamento=split(Application("metodi_incasso"),vbcrlf)

if idscadenza<>"" then
	'recupero importo scadenza
	set rs = conn.execute("select * from scadenze where idscadenza="&idscadenza)
	importo=rs("importo")
	idfat=rs("idfat")
	numero_scadenze=clng(conn.execute("select count(*) from scadenze where idfat="&idfat)(0))
	if numero_scadenze=1 then
		causale=2
	else
		causale=1
	end if
	modoincasso=tipo_pagamento(0)


end if



'oper="edit"
if idincasso<>"" then
	oper="edit"
else
	oper="new"
end if
if not isnumeric(idord) then idord=0
if request("operazione")="pagamento" then
	call carica_incasso(idord,idfat)
	Set Js = jsObject()
	Js("status")="success"
	Js("message")="Dati salvati"
	session("test")= session("test") &"<br>"&js.jsString	
	js.Flush
	set js=Nothing
	call ResponseEnd()
end if
if request("operazione")="modifica" then

	tipo_pagamento=split(Application("metodi_incasso"),vbcrlf)
	Set rs_incassi = Server.CreateObject("ADODB.Recordset")
	sql="select incassi.idincasso, ordini.idord, ordini.Nord, fatture.IDfat, fatture.Nfat FROM (incassi LEFT JOIN ordini ON incassi.idord = ordini.idord) LEFT JOIN fatture ON incassi.idfat = fatture.IDfat WHERE incassi.idincasso="&idincasso&";"
	rs_incassi.Open sql, conn, 1, 1
	idord=rs_incassi("idord")
	idfat=rs_incassi("idfat")
	txt_nord=rs_incassi("nord")
	txt_nfat=rs_incassi("nfat")
	rs_incassi.Close
	set rs_incassi=nothing
	sql="select * FROM incassi where idincasso="&idincasso
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	'memorizzo situazione iniziale
	n=rs.fields.count-1
	Set Flds = rs.Fields
	dim nome_campo()
	dim valore_campo()
	redim nome_campo(n)
	redim valore_campo(n)
	'response.write "n:"&n
	'memorizzo i nomi dei campi e i valori attuali
	for i=0 to n
		'response.write "i:"&i
		nome_campo(i)=RS.Fields(i).Name
		valore_campo(i)=converti_typevar14(rs(i))
	next

	rs("data")=request.form("data")
	rs("importo")=aggiusta_decimale(request.form("importo"),"asp")
	rs("causale")=request.form("causale")
	rs("note_pagamento")=request.form("note_pagamento")
	'rs("tipo_pagamento")=ucase(tipo_pagamento(cint(request.form("tipo_pagamento"))))
	for i=0 to n
		val_rs=converti_typevar14(rs(i))
		if valore_campo(i)<>val_rs then
				modifiche_rs=modifiche_rs&"modificato campo "&nome_campo(i)&" da:"&valore_campo(i)&" a:"&rs(i)&"<br>"
		end if
	next
	rs.update
	if modifiche_rs="" then modifiche_rs="Nessuna modifica in "&cosa
	rs.close
	set rs=Nothing
	txt="Modifica [incasso="&idincasso&"] di [ordine="&idord&"]"&txt_nord&"[/ordine]"
	if not isnull(idfat) then txt=txt&" e [fattura="&idfat&"]"&txt_nfat&"[/fattura]"
	txt=txt&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&modifiche_rs&modifiche_rs2
	add2log txt&queryeform()&"LCID:"&response.LCID ,2
	Response.ContentType = "application/json; charset=utf-8"
	Set Js = jsObject()
	Js("status")="success"
	Js("message")="Dati salvati"
	session("test")= session("test") &"<br>"&js.jsString	
	js.Flush
	set js=Nothing
	call ResponseEnd()
response.end
end if

'response.write "idord:"&idord&" idfat:"&idfat&" idincasso:"&idincasso
if oper="new" and idord>0 then
	set rs=conn.execute ("select ordini.idord, ordini.nord, ordini.totale, sommadiincassi.SommaDiimporto FROM ordini LEFT JOIN (select incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idord  )  AS sommadiincassi ON ordini.idord = sommadiincassi.idord where ordini.idord="&idord)
	If Not rs.EOF Then
		arrRS = rs.GetRows()
		if isnull(arrRS(3, 0)) then
			saldo=arrRS(2, 0)
		else
			saldo=cdbl(arrRS(2, 0))-cdbl(arrRS(3, 0))
			if saldo<=0 then
				saldo=0
			end if
		end if
		importo_max=arrRS(2, 0)
		nord=arrRS(1, 0)
		Erase arrRS
	end if
	
	set rs = conn.Execute("select fatture.IDfat FROM ordini_fatture INNER JOIN fatture ON ordini_fatture.idfat = fatture.IDfat where ordini_fatture.idord="&idord)
	if not rs.eof then
		idfat=rs("idfat")
	end if
	set rs = nothing

	operazione="pagamento"
end if


if oper="new" and idfat>0 then
	set rs=conn.execute ("select fatture.idfat, fatture.nfat, fatture.totale_fattura, sommadiincassi.SommaDiimporto FROM fatture LEFT JOIN (select incassi.idfat, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idfat  )  AS sommadiincassi ON fatture.idfat = sommadiincassi.idfat where fatture.idfat="&idfat)
	If Not rs.EOF Then arrRS = rs.GetRows()
	nfat=arrRS(1, 0)
	if isnull(arrRS(3, 0)) then
		saldo=cdbl(arrRS(2, 0))
	else
		saldo=cdbl(arrRS(2, 0))-cdbl(arrRS(3, 0))
		if saldo<=0 then
			saldo=0
		end if
	end if
	set rs = nothing
	operazione="pagamento"
	
	
	
	
end if
if oper="edit" then

	operazione="modifica"
	'Recupero dati ordine
	sql="select incassi.* from incassi where idincasso="&idincasso
	set rs_incasso=conn.execute (sql)
	idfat=rs_incasso("idfat")
	idord=rs_incasso("idord")
	if idord>0 then
		set rs=conn.execute ("select ordini.idord, ordini.nord, ordini.totale, sommadiincassi.SommaDiimporto FROM ordini LEFT JOIN (select incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idord  )  AS sommadiincassi ON ordini.idord = sommadiincassi.idord where ordini.idord="&idord)
		If Not rs.EOF Then
			arrRS = rs.GetRows()
			if isnull(arrRS(3, 0)) then
				saldo=cdbl(arrRS(2, 0))
			else
				saldo=cdbl(arrRS(2, 0))-cdbl(arrRS(3, 0))
				if saldo<=0 then
					saldo=0
				end if
			end if
			importo_max=cdbl(arrRS(2, 0))
			nord=arrRS(1, 0)
			Erase arrRS
		end if
	end if
	
	
end if
saldo=cdbl(saldo)
if saldo>0 or oper="edit" then
%>
<form id="modifica_incasso">
		<input name="operazione" type="hidden" value="<%=operazione%>">
		<input name="idincasso" type="hidden" value="<%=idincasso%>">
		<input name="idord" type="hidden" value="<%=idord%>">
		<input name="idfat" type="hidden" value="<%=idfat%>">
		<input name="saldo" id="saldo" type="hidden" value="<%=saldo%>">
		<input name="importo_max" id="importo_max" type="hidden" value="<%=importo_max%>">
		<input name="idscadenza" type="hidden" value="<%=idscadenza%>">
		
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1"   >
          <tr >
            <td  style="border-bottom:1px solid;">Data</td>
            <td  style="border-bottom:1px solid;">Importo</td>
            <td  style="border-bottom:1px solid;">Causale</td>
            <td  style="border-bottom:1px solid;">Modo incasso</td>
          </tr>
          <tr >
            <td  style="border-bottom:1px solid;">
		<%
		if oper="edit" then
			val=rs_incasso("data")
		else
			val=date()
		end if
		  %>
		  <input name="data" type="text" class="richiesto" id="data" value="<%=val%>" size="10" maxlength="10"></td>
            <td  style="border-bottom:1px solid;">
			<%
		if oper="edit" then
			val=rs_incasso("importo")
		else
			val=importo
		end if
		  
		  %><input name="importo" type="text" class="richiesto" id="importo" value="<%=val%>"></td>
            <td  style="border-bottom:1px solid;">
			<%
			if oper="edit" then
				val=0
				if not isnull(rs_incasso("causale")) then val=cint(rs_incasso("causale"))
			else
				val=causale
			end if
			%>
			<select name="causale" id="causale" class="richiesto">
                <option value="">Selezionare</option>
                <%
			
				for n=1 to causale_pagamento(-1)
					response.write "<option value='"&n&"'"
					if val=n then response.write " selected"
					response.write ">" & causale_pagamento(n)&"</option>"
				next
				%>
              </select></td>
            <td  style="border-bottom:1px solid;">
			<%
			if oper="edit" then
				val=rs_incasso("tipo_pagamento")
			else
				val=modoincasso
			end if
			%>
			<select name="tipo_pagamento" id="tipo_pagamento" class="richiesto">
                <option value="">Selezionare</option>
                <%			

for n=0 to ubound(tipo_pagamento)
	response.write "<option value='"&n&"'"
					if val=tipo_pagamento(n) then response.write " selected"
	response.write ">" & tipo_pagamento(n)&"</option>"
next
%>
              </select></td>
          </tr>
          <tr >
            <td  style="border-bottom:1px solid;">Note</td>
            <td  style="border-bottom:1px solid;">&nbsp;</td>
            <td  style="border-bottom:1px solid;">&nbsp;</td>
            <td  style="border-bottom:1px solid;">Carica su:</td>
          </tr>
          <tr >
            <td colspan="3"  style="border-bottom:1px solid;">
			<%
		if oper="edit" then
			val=rs_incasso("note_pagamento")
		else
			val=""
		end if
		  
		  %>
			<input name="note_pagamento" type="text" size="40" value="<%=val%>"></td>
            <td colspan="3"  style="border-bottom:1px solid;">
            <%
			if oper="new" and idfat>0 then
            
	              response.write "Fattura "&nfat
	              
	          else
              
              response.write "Ordine "&nord

			  end if
			  
			  %>
              
              </td>
          </tr>
        <%
	        if oper="edit" then
	        	set rs_incasso=Nothing
	        end if
			call connclose()
	          
	          
	    %>
          <tr >
            <td colspan="4"  style="border-bottom:1px solid;">
			<%if oper="edit" then%>
				<input name="modifica" type="button" value="Salva modifiche" id="salva_incasso">
			<%else %>
				<input name="aggiungi_pagamento" type="button" value="Aggiungi pagamento" onClick="Valida_pagamento()">
				<%if idfat>0 then
					 txt="fattura"
				 else
					 txt="ordine"
				 
				 end if %>
				<input name="" type="button" value="Salda <%=txt%> per € <%=saldo%>" onClick="Valida_saldaordine()">
				
			<%end if%>
			</td>
          </tr>
        </table>
        </form>
<script>
function Valida_saldaordine() 
{	
	
	if ($("#tipo_pagamento").val() == '' ) {
		alert("Selezionare il tipo pagamento");
		$("#tipo_pagamento").focus();
		return false;
	}
	if ($("#importo").val()>$("#saldo").val()){
		alert("Importo superiore al saldo");
		return;
	}
	$("#importo").val($("#saldo").val());
	$("#causale").val("2");
	invia_dati();
};
function Valida_pagamento() 
{	var str;
	str=$("#importo").val().replace(',','.');
	if (isNaN(str) ||(str == '') ) {
		alert("Inserire un valore numerico per l\'importo");
		$("#importo").focus();
		return false;
	}
	if (parseFloat(str)>parseFloat(str.replace(',','.'))  ) {
		alert("L\'importo è superiore al saldo");
		document.form1.causale_pagamento.focus();
		return false;
	}
	if ($("#causale").val() == '' ) {
		alert("Selezionare la causale");
		$("#causale").focus();
		return false;
	}
	if ($("#tipo_pagamento").val() == '' ) {
		alert("Selezionare il tipo pagamento");
		$("#tipo_pagamento").focus();
		return false;
	}
	invia_dati();
} 
function invia_dati(){	
	var dati=$( "#modifica_incasso" ).serialize();
	$.ajax({
		url     : "pag_adm_dialog_incasso.asp",
		type    : "post",
		cache: false,
		dataType: 'json',
		data	: dati,
		
		success: function(data){
			console.log ("Chiamo aggiorna_tabella_incassi");
			//aggiorna_tabella_incassi() in pag_adm_ordini .js
			aggiorna_tabella_incassi(<%=idord%>,<%=idfat%>);
			<%if idscadenza<>"" then %>
			aggiorna_tabella_scadenze(<%=idfat%>,0)
			<%end if %>
			$("#dialog").dialog("close");
		}
		,error: function(xhr, textStatus, error){
			toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
			toastr.error("Errore nell'aggiornamento dei dati");
			var txt="Errore in <%=questofile%>: invia_dati()";
			txt+="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
			txt+="Errore in "+location.href+"<br>";
			txt+="<br>"+xhr.statusText;
			txt+="<br>"+xhr.responseText;
			txt+="<br>textStatus:"+textStatus;
			txt+="<br>error:"+error;
			invia_errore("Errore in invia_dati()",txt );
		}
	}); 
}
$(document).on("click","#salva_incasso",function(e) {
	e.preventDefault();
	var importo=parseFloat($("#importo").val());
	var importo_max=parseFloat($("#importo_max").val());
	if (importo>importo_max){
		alert("Importo ("+importo+") superiore al totale ordine ("+importo_max+")");
		return;
	}
	if ($("#causale").val()==""){
		alert("Selezionare la causale");
		$("#causale").setfocus();
		return;
	}
	invia_dati();

});
	

function Ricarica() {
	<%if ricarica="fattura" then %>
	location.href="pag_adm_fatture.asp?idfat=<%=idfat%>";
	<%else %>
	location.href="pag_adm_"+tabella+".asp?idord=<%=idord%>";
	
	<%end if%>
}

</script>

<% else %>
<div style="text-align: center; background: #99FF99; height: 40px; line-height: 40px;"><span style="vertical-align: middle; font-weight: bold;">Ordine / fattura già saldato</span></div>

<% end if%>
