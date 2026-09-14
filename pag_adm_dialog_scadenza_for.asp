<%
'Verifica chiusure 02_12_2015
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="JSON_latest.asp"-->

<%
idscadenza=request("idscadenza")
idfat=request("idfat")
if idfat="" then idfat=0
ricarica=request("ricarica")
'oper="edit"
if idscadenza<>"" then
	oper="edit"
else
	oper="new"
end if
operazione=request("operazione")
if operazione="modifica" then

	sql="select * FROM scadenze_for where id="&idscadenza
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

	rs("scadenza")=request.form("data")
	rs("importo")=aggiusta_decimale(request.form("importo"),"asp")
	for i=0 to n
		val_rs=converti_typevar14(rs(i))
		if valore_campo(i)<>val_rs then
				modifiche_rs=modifiche_rs&"modificato campo "&nome_campo(i)&" da:"&valore_campo(i)&" a:"&rs(i)&"<br>"
		end if
	next
	rs.update
	idfat=rs("idfat")
	if modifiche_rs="" then modifiche_rs="Nessuna modifica in "&cosa
	rs.close
	
elseif operazione="aggiungi" then
	sql="select * FROM scadenze_for"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	rs.addnew
	rs("idfat")=idfat
	rs("scadenza")=request.form("data")
	rs("importo")=aggiusta_decimale(request.form("importo"),"asp")
	rs.update	
	rs.close
	
end if
if operazione<>"" then
	
	
	metodo=request.form("metodo")
	totale_fattura=cdbl(conn.execute("select totale_fattura from fatture_for where idfat="&idfat)(0))

	
	
	
	call elabora_scadenze()
	
	
	
	set rs=Nothing
	txt=ucase(operazione)&" [scadenza_for="&idincasso&"] "
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
	sql="select * from scadenze_for where id="&idscadenza
	set rs_incasso=conn.execute (sql)
	idfat=rs_incasso("idfat")
else
	
	operazione="aggiungi"
end if
%>
<form id="modifica_scadenza">
		<input name="operazione" type="hidden" value="<%=operazione%>">
		<input name="idscadenza" type="hidden" value="<%=idscadenza%>">
		<input name="idfat" type="hidden" value="<%=idfat%>">
		
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1"   >
          <tr >
            <td  style="border-bottom:1px solid;">Data</td>
            <td  style="border-bottom:1px solid;">Importo</td>
          </tr>
          <tr >
            <td  style="border-bottom:1px solid;">
		<%
		if oper="edit" then
			val=rs_incasso("scadenza")
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
        <%
		    %>    
		    <tr>
			    <td colspan="2">
				    Scadenze successive:  <input name="metodo" type="radio" value="1" checked> Conguaglio su ultime <input name="metodo" type="radio" value="2"> Ridistribuisci				    
			    </td>
		    </tr>
		        
		        
		        
		      <%  
	        	set rs_incasso=Nothing
			call connclose()
	          
	          
	    %>
          <tr >
            <td colspan="4"  style="border-bottom:1px solid;">
			<%if oper="edit" then%>
				<input name="modifica" type="button" value="Salva modifiche" id="salva_incasso">
			<% else %>
				<input name="aggiungi" type="button" value="Aggiungi" id="salva_incasso">
			
			<%end if%>
			</td>
          </tr>
        </table>
        </form>
<script>
function invia_dati(){	
	var dati=$( "#modifica_scadenza" ).serialize();
	$.ajax({
		url     : "pag_adm_dialog_scadenza_for.asp",
		type    : "post",
		cache: false,
		dataType: 'json',
		data	: dati,
		
		success: function(data){
			console.log ("Chiamo aggiorna_tabella_incassi");
			//aggiorna_tabella_incassi() in pag_adm_ordini .js
			aggiorna_tabella_scadenze_for(<%=idfat%>);
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
	e.stopImmediatePropagation();

});
	

function Ricarica() {
	<%if ricarica="fattura" then %>
	location.href="pag_adm_fatture.asp?idfat=<%=idfat%>";
	<%else %>
	location.href="pag_adm_"+tabella+".asp?idord=<%=idord%>";
	
	<%end if%>
}

</script>
<%
	
sub elabora_scadenze()
		totale_fattura=cdbl(conn.execute("select totale_fattura from fatture_for where idfat="&idfat)(0))

		if metodo=1 then 'Conguaglio
		sql="select * from scadenze_for where idfat="&idfat&" order by scadenza"
		rs.open sql, conn, 3,3
		totale_progressivo=0
		elimina=false
		do while not rs.EOF
			importo=cdbl(rs("importo"))
			if totale_progressivo+importo=>totale_fattura then
				rs("importo")=totale_fattura-totale_progressivo
				rs.update
				elimina=true
				
			elseif elimina then
				
				rs.delete
			elseif totale_progressivo=totale_fattura then
				elimina=true
			else
				totale_progressivo=totale_progressivo+importo
			end if
				
		
			rs.MoveNext
		loop
		if totale_progressivo<totale_fattura and not elimina then
			rs.movelast
			
			rs("importo")=cdbl(rs("importo"))+(totale_fattura-totale_progressivo)
			rs.update
		end if
		
	elseif metodo=2 then	'Ridistribuisci
		n_rate_restanti=cint(conn.execute("select count(*) from scadenze_for where idfat="&idfat&" and scadenza>STR_TO_DATE('"&request.form("data")&"', '%d/%m/%Y') ")(0))
		'response.write "n_rate_restanti:"&n_rate_restanti
		totale_rate_precedenti=cdbl(conn.execute("select sum(importo) from scadenze_for where idfat="&idfat&" and scadenza<=STR_TO_DATE('"&request.form("data")&"', '%d/%m/%Y') ")(0))
		'response.write "totale_rate_precedenti:"&totale_rate_precedenti
		
		n_rata=0
		importo=(totale_fattura-totale_rate_precedenti)/n_rate_restanti
		importo=RoundUp(importo, 2)
		sql="select * from scadenze_for where idfat="&idfat&" and scadenza>STR_TO_DATE('"&request.form("data")&"', '%d/%m/%Y')  order by scadenza"
		rs.open sql, conn, 3,3
		totale_progressivo=0
		do while not rs.EOF
			'call add2log("processo "&rs("scadenza"),0)
			if n_rata<n_rate_restanti then
				rs("importo")=importo
				totale_progressivo=totale_progressivo+importo
			else
				rs("importo")=totale_fattura-totale_rate_precedenti-totale_progressivo
			end if
		
			rs.update
		
			n_rata=n_rata+1
			rs.MoveNext
		loop
	
	end if

	
	
end sub	
	
	
	
	
	%>

