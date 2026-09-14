<!--#include virtual="/setup.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="JSON_latest.asp"-->
<%
'Edita le spettanze di singola riga ordine
if session("idadmin") = "" then call login()
nord=request("nord")
idord=request("idord")
oper=request("oper")
m_tabella=request("tabella")
iddett=request("iddett")
dim quantita_tot
if oper="salva" then
	quantita_tot=0
	sql="UPDATE "&m_tabella&" SET calcolato = 0 WHERE idord = "&idord&";"
	conn.execute (sql)
	
	'ciclo su ordini_dett_spettanze
	sql="select * from "&m_tabella&"_dett_spettanze  where iddett="&iddett
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql, conn,3,3
	do while not rs.eof
		newq=request.form("quantita_"&rs("id"))
		if newq="" or newq="0" then
			rs.delete
		else
			if isnumeric(newq) then
				rs("quantita")=newq
			end if
			quantita_tot=quantita_tot+clng(rs("quantita"))
		end if
		rs.MoveNext
	loop
	avanzo=request.form("avanzo")
	if avanzo<>"" and isnumeric(avanzo)  then
		if avanzo>0 then
			'Creo record per avanzo con iddip=0
			conn.execute("INSERT INTO ordini_dett_spettanze (iddett, iddip, quantita, taglia_misura) VALUES ( "&iddett&", '0', "&avanzo&",'"&request.form("avanzo_taglia")&"');")
			
			quantita_tot=quantita_tot+avanzo
		end if
	end if
	rs.Close
	
	iddip=request("iddip_aggiungi")
	quantita_aggiungi=request("quantita_aggiungi")
	if quantita_aggiungi<>"" and iddip<>"" then
		taglia_misura=""
		set rs_prodotto=conn.execute("select prodotti.idpro, iddett, tipo_taglia from ("&m_tabella&"_dett left join prodotti on "&m_tabella&"_dett.idpro=prodotti.idpro) where iddett="&iddett)
		if not isnull(rs_prodotto("idpro")) then
			taglia_misura=recupera_taglia(rs_prodotto("tipo_taglia"),iddip)
		end if
	
	

			conn.execute("INSERT INTO "&m_tabella&"_dett_spettanze (iddett, iddip, quantita, taglia_misura) VALUES ( "&iddett&","&iddip&", "&quantita_aggiungi&",'"&taglia_misura&"');")
			
			quantita_tot=quantita_tot+quantita_aggiungi
		
		
	end if
	
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "&m_tabella&"_dett where iddett="&iddett
	rs.open sql, conn,3,3
	rs("quantita")=quantita_tot
	m_totale_riga=totale_riga(rs("prezzo"),rs("sconto_prodotto"),quantita_tot)
	rs("totale_riga")=m_totale_riga
	rs.update
	rs.Close
	
		
	
	
	
	set rs=Nothing
	
	
	Set Js = jsObject()
	Js("success")=true
	
	
	call add2log("Aggiornato spettanze riga "&iddett&"  [ordine="&idord&"]"&nord&"[/ordine]"&vbcrlf&txt_log,2)
	js.Flush
	set js=Nothing
	call ResponseEnd()
elseif oper="aggiungi" then	
	
end if
select case oper
	case "annulla"
		oper="view"
	case "new"
		oper="new"
	case "aggiungi"
		oper="add"
	case "modifica"
		oper="update"
	case else
		oper="new"
		if request("idpro_a")<>"" then oper="view"
end select
tabellaid="idord"
%>
<script Language="JavaScript"> 
$(function() {
	$( "input" ).tooltip();
});
</script>
<%'response.write "ordine:"&ordine%>
<form id="modifica_spettanze" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" >
  <input type="hidden" name="tabella" value="<%=m_tabella%>" />
  <input type="hidden" name="iddett" value="<%=iddett%>" />
  <input type="hidden" name="idord" value="<%=idord%>" />
  <input type="hidden" name="nord" value="<%=nord%>" />
    <%
	    sql="select ordini_dett.*, magazzino.idmag, magazzino.db_ven from (ordini_dett LEFT JOIN prodotti ON ordini_dett.idpro = prodotti.idpro) left join magazzino on ( ordini_dett.idvarb = magazzino.idvarb) AND ( ordini_dett.idvara = magazzino.idvara) AND ( ordini_dett.idpro = magazzino.idpro) where iddett="&iddett
	    
	    sql=replace(sql,"ordini",m_tabella)
	    
	    
	    
    set rs = conn.Execute(sql)
    response.write "<b>"&rs("articolo_ordine")&"</b><br>"
    quantita_tot=cdbl(rs("quantita"))
    set rs = nothing
    
	sql="select "&m_tabella&"_dett_spettanze.* , dipendenti.cognome, dipendenti.nome, dipendenti.sesso, dipendenti.matricola, utenti_gradi.nome_grado, utenti_gradi.colore_nominativo from (("&m_tabella&"_dett_spettanze left join dipendenti on "&m_tabella&"_dett_spettanze.iddip=dipendenti.iddip) left join utenti_gradi on dipendenti.grado = utenti_gradi.id) where iddett="&iddett &" order by sesso desc, cognome, nome"
	
	
	
	scorta_tecnica=false
	'response.write sql
    set rs_spettanze=conn.execute (sql)
    do while not rs_spettanze.eof
        quantita_tot=quantita_tot-cdbl(rs_spettanze("quantita"))
%>
	<input type="text" name="quantita_<%=rs_spettanze("id")%>" size="2"  value="<%=rs_spettanze("quantita")%>" style="text-align:right;">x 
	<%
		if isnull(rs_spettanze("cognome")) then
			response.write "Scorta tecnica"
			scorta_tecnica=true
		else
			response.write dipendente_colorato(  rs_spettanze("cognome")&" "&rs_spettanze("nome"),rs_spettanze("sesso"), rs_spettanze("colore_nominativo") )
		end if
		%>
		<br>
		<%
	    rs_spettanze.MoveNext
	    Loop
	    set rs_spettanze = nothing
	    if quantita_tot<0 then
		    quantita_tot=0
	    end if
	    if not scorta_tecnica then
	    %>
    	<input type="text" name="avanzo" size="2" value="<%=quantita_tot%>" style="text-align:right;">Scorta tecnica taglia-misura <input type="text" name="avanzo_taglia" value="" ><br>
		<%
		end if
		%>
	    <input type="text" name="quantita_aggiungi" size="2"  value="" style="text-align:right;">x <input type='hidden' name='iddip_aggiungi' id='iddip' style="width:350px;" tabindex="6" />
		
		<div class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix">
			<div class="ui-dialog-buttonset" >
				<input type="button" id="aggiorna_quantita" value="Salva modifiche" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only">
			</div>
		</div>
		<%
    	sql="select iduser from "&m_tabella&" where idord="&idord
		iduser=conn.execute(sql)(0)
		call connclose()
	  
	  %>
  <script type="text/javascript">
  $(document).ready(function () {
	//Inizializzo select articolo
		
	
	
	
	$("#aggiorna_quantita").click(function(e){
		var dati=$( "#modifica_spettanze" ).serialize();
		console.log("dati:"+dati);
		$.ajax({
			url     : "pag_adm_dialog_ordini_edit.asp?oper=salva",
			type    : "post",
			dataType: "json",
			cache	: false,
			data	: dati,
			success: function(data){
				if (data.success){
					aggiorna_tabella("");		
					$("#dialog").dialog("close");
				}
				else
				{
					alert(data.message);
				}
			}
			,error: function(xhr, textStatus, error){
				console.log("Errore aggiorna_quantita");
				invia_errore("Errore in quantitamagazzino()",xhr.statusText+xhr.responseText+textStatus+error );
			      
				  }
		});
	});
 	var iduser=<%=iduser%>;
	$('#iddip').select2({
		dropdownCssClass: 'ui-dialog',
		placeholder: 'Seleziona un dipendente',
		minimumInputLength: 0,
		allowClear: true,
		ajax: {
			quietMillis: 150,
			url: "ajax_function.asp?select2=dipendenti",
			dataType: 'json',
			data: function (term, page) {
				return {
					term: term
					,iduser: iduser
				};
			},
			results: function (data) {
				return {
					results: data
				};
			}
		}
	});
	
	
});
</script>
</form>
<%
call CheckConnChiusa()
%>