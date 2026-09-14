<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/paginazione.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->

<%
if session("idadmin") = "" then call login()
idmulti=request.form("cercamulti")
selezione=request("selezione")
if request.form("stato")<>"" then
    sql_stato=""
    Astato=split (request.form("stato"),",")
    for n=0 to UBound(Astato)
        call concatena_stringa(sql_stato," or ","stato="&trim(Astato(n)))
        
    next
    sql_stato=" and ("&sql_stato&")"
    
end if

%>
<!--#include virtual="/sub_head_adm.asp" -->
<style>
	
	.padding{
		padding:2px;
	}
	.center{
	text-align:center;
	}
	</style>
</head>
<body> 
<div id="wrap">
<div id="header">
<%=titolo_top%>
        <!-- Div Content INIZIO-->      <%barra=0%>
	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Cerca articoli in ordini</a></div> 
                <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin:0px 0px 0px 0px;" >
               <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr > 
            <td colspan="4" style="border-bottom:1px solid;">
				<select name="selezione" id="selezione">
					<option value="A" <%if selezione="A" then%>selected<%end if%>>Articolo</option>
					<option value="F" <%if selezione="F" then%>selected<%end if%>>Fornitore</option>
				</select>

		<input type='hidden' name='cercamulti' id='cercamulti' style="width:350px;" tabindex="1" value=""/><br>
		<%
			if request.form("stato")<>"" then
				stato=request.form("stato")
			else
				stato="235"
			end if
			%>
		
        <span class="stato_1 padding"><input type="checkbox" name="stato" value="1" <%if instr(stato,"1")>0 then response.write "checked" %>><%=stato_ordine(1)%></span>
		<span class="stato_2 padding"><input type="checkbox" name="stato" value="2" <%if instr(stato,"2")>0 then response.write "checked" %>><%=stato_ordine(2)%></span>
		<span class="stato_3 padding"><input type="checkbox" name="stato" value="3" <%if instr(stato,"3")>0 then response.write "checked" %>><%=stato_ordine(3)%></span>
		<span class="stato_5 padding"><input type="checkbox" name="stato" value="5" <%if instr(stato,"5")>0 then response.write "checked" %>><%=stato_ordine(5)%></span>
		<span class="stato_6 padding"><input type="checkbox" name="stato" value="6" <%if instr(stato,"6")>0 then response.write "checked" %>><%=stato_ordine(6)%></span>
<input type="submit" name="submit" value="Cerca" style="FONT: 12px;"  id="puls_cerca">       
       
              </td>
          </tr></table></form>
        <%if selezione="A" and idmulti<>"" then
	        set rs=conn.execute("select codice,articolo from prodotti where idpro="&idmulti)
	        articolo=rs("articolo")
	        codice=rs("codice")
	        
	        %>
        <div class="ui-widget-header ui-corner-all " style="padding: 5px;"><%=codice%>&nbsp;<%=articolo%></div> 
	        <%
		end if
			 %>
			<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
				<tr>
					<td>Ordine</td>
					<td>Cliente</td>
					<td>Quantit&agrave;</td>
					<td>Spettanze</td>
					<td>Prezzo</td>
				</tr>
				
			<%
			
			if selezione="A" then
			
			
				sql="select ordini_dett.iddett, quantita, prezzo, ordini.idord, ordini.nord, ordini.stato, ordini.data, utenti_intestazioni.nome, utenti_intestazioni.cognome, utenti_intestazioni.azienda, ordini.tipo_documento, ordini_dett_note.nota from ordini_dett inner join ordini on ordini_dett.idord = ordini.idord inner join utenti_intestazioni on ordini.idintestazione = utenti_intestazioni.id left join ordini_dett_note on ordini_dett.iddett=ordini_dett_note.iddett where (tipo_documento='ordine' or tipo_documento='sostituzione') and eliminato=0 and idpro="&idmulti&sql_stato&" order by ordini.data desc"
				if utente_andrea then response.write "sql:"&sql
				
				set rs=conn.execute (sql)
				call elenca_articoli(rs)
				
			elseif selezione="F" then
				set rs_articoli=conn.execute("select idpro, codice, articolo from prodotti where idfor="&idmulti)
				
				do while not rs_articoli.eof

				
					idpro=rs_articoli("idpro")
					sql="select ordini_dett.iddett, quantita, prezzo, ordini.idord, ordini.nord, ordini.stato, ordini.data, utenti_intestazioni.nome, utenti_intestazioni.cognome, utenti_intestazioni.azienda, ordini_dett_note.nota from ordini_dett inner join ordini on ordini_dett.idord = ordini.idord inner join utenti_intestazioni on ordini.idintestazione = utenti_intestazioni.id left join ordini_dett_note on ordini_dett.iddett=ordini_dett_note.iddett where tipo_documento='ordine' and eliminato=0 and idpro="&idpro&sql_stato&" order by ordini.data desc"
				
					set rs=conn.execute (sql)
					if not rs.eof then
						totale=0
									%>
				<tr><td colspan="5" style="background:rgba(33, 33, 33, 0.5);"><strong>
				<%=rs_articoli("codice")&" "&rs_articoli("articolo")%></strong>
				
				</td></tr>
				
				<%

					end if
				call elenca_articoli(rs)

				
				
				
					rs_articoli.movenext
				loop
				
			
			end if

				%>	
			</table>
			<%
		set rs = Nothing


	
	
		call connclose()

	
%>
        <!-- FINE BLOCCO CENTRALE -->

</div>

<script>
	function Valida_ricerca() 
{	posizione_errore="Valida_ricerca";
	if ((document.form1.idpro.value == '' )) {
		alert("Selezionare un articolo");
		document.form1.idpro.focus();
		return false;
	}

	document.form1.submit();
} 
	
	$(function() {
	
	$(document).ready(function () {
	function attiva_select2(Url,Placeholder){
	
	  	$('#cercamulti').select2({
				placeholder: Placeholder,
				minimumInputLength: 1,
				//allowClear: true,
				ajax: {
					quietMillis: 150,
	  			url: Url,
					dataType: 'json',
					data: function (term, page) {
						return {
							term: term
						};
					},
					results: function (data) {
						
						return {
							results: data
						};
					}
				}
			});
		}

    $("#selezione").on('change', function() {
	    
		impostaselect();
		
    });
	
	impostaselect();
    		function impostaselect(){
	if($("#selezione").val()=="A")
	{
	  	attiva_select2("ajax_function.asp?select2=articoli_sel2","Cerca singolo articolo");
	}
	else if($("#selezione").val()=="F")
	{
	  	attiva_select2("ajax_function.asp?select2=fornitori_sel2","Cerca fornitore");
	}
	else
	{
	  	attiva_select2("ajax_function.asp?select2=articoli_sel2","Cerca singolo articolo");
    	
	}   }

});


	
	

	});
</script>	

<div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
<%


sub elenca_articoli(rs)
	dim totale
	totale=0

				do while not rs.EOF
				stile_cella=""
				txt_spettanze=mid(dettaglio_spettanze("ordini", rs("iddett"), rs("quantita"), false, stile_cella ),5)
				%>
				
				<tr style="border-top: 1px solid black;">
					<%if rs("tipo_documento")="ordine" then
					
				
					
					 %>
					
					<td class="stato_<%=rs("stato")%> center" >ordine<br><strong><a href="pag_adm_ordini.asp?idord=<%=rs("idord")%>" target="_blank"><%=rs("nord")%></strong></a><br><%=formatdatetime(rs("data"),2)%></td>
					
					<% else %>
					<td class="center" ><strong>sostituzione<br><a href="pag_adm_ordini.asp?idord=<%=rs("idord")%>" target="_blank"><%=rs("nord")%></strong></a><br><%=formatdatetime(rs("data"),2)%></td>					
					
					<%end if %>
					
					
					<td><%=denominazione(rs("nome"),rs("cognome"),rs("azienda"))%></td>
					<td class="<%=stile_cella%>" style="text-align: center;"><%=rs("quantita")%></td>
					<td class="<%=stile_cella%>"><%=txt_spettanze%>
					<%nota=rs("nota")
						
						
					if  nota<>"" then %>
					
					
					<br><span class="notaarticolo"><%=nota%></span>
					
					
					<%end if %>
					 </td>
					<td class="<%=stile_cella%>"><%=formatnumber(rs("prezzo"),2)%></td>
				</tr>
				<%
					totale=totale+cdbl(rs("quantita"))
				rs.MoveNext
				loop
				%>
				<tr>
					<td colspan="2" style="text-align: right;"><strong>Totale</strong></td><td style="text-align: center;"><strong><%=totale%></strong></td><td colspan="2"> </td>
				</tr>
				<%

end sub


%>
