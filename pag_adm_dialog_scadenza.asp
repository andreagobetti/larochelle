<%
'Verifica chiusure 02_12_2015
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="JSON_latest.asp"-->

<%
idscadenza=request("idscadenza")
if idfat="" then idfat=0
ricarica=request("ricarica")
'oper="edit"
if idscadenza<>"" then
	oper="edit"
else
	oper="new"
end if

if request("operazione")="modifica" then

	sql="select * FROM scadenze where idscadenza="&idscadenza
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
	rs("note_pagamento")=request.form("note_pagamento")
	for i=0 to n
		val_rs=converti_typevar14(rs(i))
		if valore_campo(i)<>val_rs then
				modifiche_rs=modifiche_rs&"modificato campo "&nome_campo(i)&" da:"&valore_campo(i)&" a:"&rs(i)&"<br>"
		end if
	next
	idfat=rs("idfat")
	rs.update
	if modifiche_rs="" then modifiche_rs="Nessuna modifica in "&cosa
	rs.close
	set rs=Nothing
	txt="Modifica [scadenza="&idscadenza&"] "
	set rs_fattura=conn.Execute("select idfat, nfat, pa, fattura_ce from fatture where idfat="&idfat)
	txt_nfat=rs_fattura("nfat")&get_fattura_pa(rs_fattura("pa"),rs_fattura("fattura_ce"))
	set rs_fattura = Nothing
	if not isnull(idfat) then txt=txt&" e [fattura="&idfat&"]"&txt_nfat&"[/fattura]"
	txt=txt&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&modifiche_rs&modifiche_rs2
	add2log txt&queryeform()&"LCID:"&response.LCID ,2
	Response.ContentType = "application/json; charset=utf-8"
	Set Js = jsObject()
	Js("status")="success"
	Js("message")="Dati salvati"
	js.Flush
	set js=Nothing
	call ResponseEnd()
response.end
end if


if oper="edit" then

	operazione="modifica"
	'Recupero dati ordine
	sql="select * from scadenze where idscadenza="&idscadenza
	set rs_incasso=conn.execute (sql)
	idfat=rs_incasso("idfat")
	idord=rs_incasso("idord")
	
end if
if oper="edit" then
%>
<form id="modifica_scadenza">
		<input name="operazione" type="hidden" value="<%=operazione%>">
		<input name="idscadenza" type="hidden" value="<%=idscadenza%>">
		<input name="idfat" type="hidden" value="<%=idfat%>">
		<input name="idfat" type="hidden" value="<%=idord%>">
		
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1"   >
          <tr >
            <td  style="border-bottom:1px solid;">Data</td>
            <td  style="border-bottom:1px solid;">Importo</td>
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
			val=""
		end if
		  
		  %><input name="importo" type="text" class="richiesto" id="importo" value="<%=val%>"></td>

          </tr>
          <tr >
            <td  style="border-bottom:1px solid;">Note</td>
          </tr>
          <tr >
            <td   style="border-bottom:1px solid;">
			<%
		if oper="edit" then
			val=rs_incasso("note_pagamento")
		else
			val=""
		end if
		  
		  %>
			<input name="note_pagamento" type="text" size="40" value="<%=val%>"></td>
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
	var dati=$( "#modifica_scadenza" ).serialize();
	$.ajax({
		url     : "pag_adm_dialog_scadenza.asp",
		type    : "post",
		cache: false,
		dataType: 'json',
		data	: dati,
		
		success: function(data){
			console.log ("Chiamo aggiorna_tabella_incassi");
			//aggiorna_tabella_incassi() in pag_adm_ordini .js
			aggiorna_tabella_scadenze(<%=idfat%>,<%=idord%>);
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


<% end if%>
