<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/regioni_inc.asp" -->
<!--#include file="JSON_latest.asp"-->
<!--#include virtual="/ClasseOrdine.asp" -->
<%
iduser=request("iduser")
idord=request("idord")
cosa=request("cosa")
parziale=request("parziale")
dim mostra_ddt
dim mostra_fattura
mostra_ddt=true
mostra_fattura=true
documento=request.form("documento")

mostra_fattura=false
mostra_ddt=false

ddt_o_fattura=request.form("ddt_o_fattura")
if instr(ddt_o_fattura,"D")>=1 then
	mostra_ddt=true
end if
if instr(ddt_o_fattura,"F")>=1 then
	mostra_fattura=true
	tipofattura="accompagnatoria"

end if

'if documento="ddt" then
'	
'	cosa="D"
'	mostra_ddt=true
'	mostra_fattura=false
'end if
'if cosa="D" then
'	mostra_ddt=false
'	mostra_fattura=true
'	tipofattura="fine_mese"
'end if
'if cosa="F" then
'	mostra_ddt=false
'	mostra_fattura=true
'	tipofattura="fine_mese"
'end if




if idord<>"" then
	set rs_ordine=conn.execute ("select * from ordini where idord="&idord)
	iduser=rs_ordine("iduser")
	peso=rs_ordine("peso")
	dimensione=rs_ordine("dimensione")
	tipopagamento=rs_ordine("tipopagamento")
	trattamento_iva_ordine=rs_ordine("trattamento_iva_ordine")
	trasporto=rs_ordine("trasporto")
	idconsegna=rs_ordine("idconsegna")
	set rs_ordine = Nothing
else
	iduser=request("iduser")
end if
operazione=request("operazione_ddt")
if operazione<>"" then

	Set Js = jsObject()
	if operazione="crea_fattura" then
		set rs_ddt=conn.execute("select idddt from ddt where idord="&idord)
		if rs_ddt.eof then
			tipo="accompagnatoria"
			
		else
			tipo="fine_mese"
		end if
		
		
		idfat_creata=crea_fattura(tipo)
		
		
		
		if idfat_creata<>"-1" then
			idfat_creata=split(idfat_creata,"|")
			Js("success")=true
			Js("message")="Fattura creata"
			js("cosa")="F"
			js("idfat")=idfat_creata(0)
			js("nfat")=idfat_creata(1)
			js("data")=FormatDateTime(date(),2)
		else
			Js("success")=false
			Js("message")="Esiste già una fattura per questo ordine"
		end if
	end if
	
	
	if operazione="crea_ddt" then
		call add2log(queryeform(),0)
		
		set ordine= (new ClasseOrdine)(array("nuovoddt",request.form("iduser"),""))

		idddt_creato=ordine.idord()
		nddt_creato=ordine.nord()
		
		'idddt_creato=crea_ddt()
		
		if idddt_creato<>"-1" then
			'if request.form("spettanze")<>"" then
				
			'	call add2log("aggiorno spettanze",0)
			'	m_tabella="ordini"
			'	set ordine= (new ClasseOrdine)(array("aggiorna_quantita_spettanze",m_tabella,idord,idddt_creato))
				
			'end if
			
			Js("success")=true
			Js("message")="DDT creato"
			js("cosa")="D"
			Js("idddt")=idddt_creato
			Js("nddt")=nddt_creato
			js("data")=FormatDateTime(date(),2)
		else
			Js("success")=false
			Js("message")="Esiste gi&agrave; una DDT per questo ordine"
		end if
	end if




	js.Flush
	set js=Nothing
	call ResponseEnd()


end if
dim ccausale_ddt
if utente_andrea then
	response.write "[soloio]"& queryeform()
	response.write "mostra_ddt:"&mostra_ddt&"mostra_fattura:"&mostra_fattura
end if

	%>
<!-- CREA DDT-->
<form id="form_crea_ddtfattura" name="form_ddt" method="post" action="<%=questofile%>" style="margin-top:0px;">


<div id="tabs">
  <ul>
    <li><a href="#finestra_ddt">Dati documento</a></li>
    <li><a href="#dati-consegna">Dati consegna</a></li>
  </ul>

	<div id="finestra_ddt" style="<%=display_none%> border: 0px; padding-bottom:0px;">
			
		<input type="hidden" name="tipofattura" value="">	
		<input type="hidden" name="nord" value="<%=nord%>">
		<input type="hidden" name="parziale" id="parziale" value="<%=parziale%>">
			<table width="100%" border="0" cellpadding="" cellspacing="0" class="">
			<%if month(date())=1 then %>
			<tr>
			<td>
			Anno
			</td>
			<td>
			<%
				anno=year(date())
				annoprecedente=anno-1
			%>
			<input type="radio" name="anno" value="" checked><%=anno%>		<input type="radio" name="anno" value="<%=annoprecedente%>"><%=annoprecedente%>
	
			</td>
			</tr>
			
			<%end if 
				
				sql="select utenti.*,utenti_intestazioni.*, utenti_clienti.FatturaPACodiceDestinatario,utenti_clienti.tipo, utenti_clienti.stato_estero  FROM utenti LEFT JOIN utenti_clienti ON utenti.iduser = utenti_clienti.iduser inner join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id where utenti.iduser="&iduser
				set rs_user=conn.execute(sql)
				FatturaPACodiceDestinatario=rs_user("FatturaPACodiceDestinatario")
				vtipocliente=rs_user("tipo")
				idintestazione=rs_user("idintestazione")
				da_fatturare=1
				if valore_vuoto(vtipocliente) then
					'Se non ancora impostato determino possibile tipo cliente
					vtipocliente=determina_tipo_cliente(rs_user("azienda"),rs_user("cf"),rs_user("piva"))
					testo_tipo="<span class=""rosso"">Tipo cliente "&TipoCliente(vtipocliente)&" non impostato ma determinato in automatico in base ai dati anagrafici</span>"
				end if
				if mostra_ddt then 
				%>
				<tr>
					<td valign="top" colspan="2" bgcolor="#E5E5E5">Dati DDT</td>
	            </tr>
	            <tr>
					<td valign="top">Causale</td>
					<td><select name="causale" class="richiesto" id="causale">
						<%Set ccausale_ddt = new cl_causale_ddt
	
		                  ccausale_ddt.stampa_option(1)
		                  set ccausale_ddt = Nothing
					%>
					</select>
					<input name="iduser" type="hidden" value="<%=iduser%>"></td>
	            </tr>
	            <tr>
	              <td valign="top">Da fatturare</td>
	              <td>
		              <input type="checkbox" name="da_fatturare" id="da_fatturare" value="1" <%if da_fatturare=1 then response.write " checked"%>>
		              
	                </td>
	            </tr>
	            <tr>
	              <td valign="top">Porto</td>
	              <td><select name="porto" id="porto" class="richiesto">
	                  <%
						for n=1 to porto_ddt(0)
						response.write "<option value='"&n&"'"
						if n=3 then response.write " selected"
						response.write ">" & porto_ddt(n)&"</option>"
						next
						%>
	                </select></td>
	            </tr>
	            <tr>
	              <td valign="top">Imballo</td>
	              <td><select id="imballo" name="imballo" class="richiesto">
	                  <%
						for n=1 to imballo_ddt(0)
						response.write "<option value='"&n&"'>" & imballo_ddt(n)&"</option>"
						next
						%>
	                </select></td>
	            </tr>
	            <tr>
	              <td valign="top">Numero colli</td>
	              <td><input type="text" id="colli" name="colli" size="8" maxlength="3" value="1" class="richiesto"></td>
	            </tr>
	            <tr>
	              <td valign="top">Peso</td>
	              <td><%
				  val=""
				  if peso<>"" then val=peso
				  %>
	                <input type="text" id="peso" name="peso" size="8" maxlength="5" value="<%=val%>" class="richiesto"></td>
	            </tr>
	            <tr>
	              <td valign="top">Dimensione</td>
	              <td><%
				  val=""
				  if dimensione<>"" then val=dimensione
				  %>
	                <input type="text" id="dimensione" name="dimensione" size="50" maxlength="50" value="<%=val%>"></td>
	            </tr>
	            <%end if%>
	            <tr>
	              <td valign="top">Annotazioni</td>
	              <td><input type="text" name="annotazioni" size="50"  value=""></td>
	            </tr>
	            <%
				if mostra_ddt then 
		            %>
	            <tr>
	              <td valign="top">Incaricato del trasporto</td>
	              <td><select name="incaricato" class="richiesto">
	                  <%
						for n=1 to incaricato_ddt(0)
						response.write "<option value='"&n&"'>" & incaricato_ddt(n)&"</option>"
						next
						%>
	                </select>
	               </td>
	            </tr>
				 <tr>
	              <td valign="top">Vettore</td>
	              <td>
	                <input type="text" name="vettore" size="50" value="<%=rs_user("note_trasporto_2")%>"></td>
	            </tr>
	            
	            <tr>
	              <td valign="top">Data inizio trasporto</td>
	              <td><input type="text" name="data_inizio" size="12" maxlength="10" value="<%= formatDateTime(date(), vbShortDate)%>"></td>
	            </tr>
	            <%end if %>
	            <%if mostra_fattura then
		            accompagnatoria=""
	            	if mostra_ddt then
	            		accompagnatoria="accompagnatoria"
	            		
	            	else
		            	'C'e ddt, cerco i dati nell'ordine
		            	set rs_ddt=conn.execute("select * from ddt where idord="&idord)
		            	if rs_ddt("tipo_ddt")<>"no_ordine" then
			            	sub_idord=rs_ddt("sub_idord")
			            	set rs_ordine=conn.execute("select * from ordini where idord="&sub_idord)
			            	tipopagamento=rs_ordine("tipopagamento")
			            	trattamento_iva_ordine=rs_ordine("trattamento_iva_ordine")
			            	da_ordine="Dati da ordine "&rs_ordine("nord")
			            	
			            end if
	            		
	            
	            	end if
	             %>
	            <tr>
	              <td valign="top" colspan="2" bgcolor="#E5E5E5">Dati per fattura <%=accompagnatoria%></td>
	            </tr>
	            <%
		         if   da_ordine<>"" then %>
		         
	            <tr>
	              <td valign="top" colspan="2" bgcolor=""><b><%=da_ordine%></b></td>
	            </tr>
		         
		         
		         <%end if  
		            
		            
		            
		            
				if testo_tipo<>"" then
					%>
					
	            <tr>
	              <td valign="top" colspan="2" ><%=testo_tipo%></td>
	            </tr>
					<%end if
				italia=true
				
				if rs_user("stato_estero")=1 then  italia=false
				if tipopagamento="0" then
					pagamento=rs_user("pag_accordato")
				else
					if tipopagamento<>"" then
						pagamento=cint(tipopagamento)
					else
						pagamento=0
					end if
				end if
				%>
	            <tr>
	              <td valign="top">Pagamento</td>
	              <td>
		              <select name="pagamento" id="pagamento" class="richiesto">
			              <option value="">Selezionare</option>
	                <%=metodo_pagamento(-1,pagamento)%>
		              </select>
	                </td>
	            </tr>
	            <tr>
	              <td valign="top">Fattura PA</td>
	              <td>
	                <%
		            if  trattamento_iva_ordine=10 then 
			            checkedsi=" checked"
			            
			        else
				        checkedno=" checked"
			        end if
					if isnull(FatturaPACodiceDestinatario) or FatturaPACodiceDestinatario="" then
						errore_pa="Manca il codice destinatario"
						checkedsi=""
				        checkedno=" checked"
				        disabledsi="disabled"
					end if
	            %>SI <input type="radio" name="fattura_pa" value="1" <%=checkedsi%> <%=disabledsi%>>  NO<input type="radio" name="fattura_pa" value="0" <%=checkedno%>> <%=errore_pa%>
	                </td>
	            </tr>
	            <%
		            
	            if  errore_pa="" and trattamento_iva_ordine=10 then %>
	            
	            
		        <tr>
	              <td valign="top">Codice univoco</td>
	              <td>              
	              <%
							if errore_pa="" then
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
						response.write ">" & riga&"</option>"
					next
					%>
	                </select>
					
					<%end if%>
	
	              
	              </td>
	            </tr>
	            <%end if
%>
	            
	            
	            
	            
	            <tr>
	              <td valign="top">Pagamento bollettino</td>
	              <td><select name="pagamentobollettino">
	                  <option value="">Selezionare</option>
	                  <option value="cont">Contante</option>
	                  <option value="ab">Assegno bancario</option>
	                  <option value="asr">Assegno come rilasciato</option>
	                  <option value="asrp">Assegno come rilasciato no poste</option>
	                </select></td>
	            </tr>
	            <tr>
	              <td valign="top">Spese bancarie</td>
	              <td><input type="text" name="spese_bancarie" size="50"  value="" autocomplete=off></td>
	            </tr>
				<%if italia and (vtipocliente = "S" or vtipocliente = "D") then %>
	            <tr>
	              <td valign="top">Partita IVA</td>
	              <td><input type="text" name="piva" size="50" value="<%=rs_user("piva")%>" autocomplete=off class="richiesto"></td>
	            </tr>
				<%end if %>
	            <tr>
	              <td valign="top">Codice fiscale</td>
	              <td><input type="text" name="cf" size="50"  value="<%=rs_user("cf")%>" autocomplete=off class="richiesto"></td>
	            </tr>
	            <tr>
	              <td valign="top">Banca appoggio cliente</td>
	              <td><input type="text" name="banca_appoggio" size="50" value="<%=rs_user("banca_appoggio")%>" autocomplete=off></td>
	            </tr>
	            <tr>
	              <td valign="top">IBAN cliente</td>
	              <td><input type="text" name="iban" size="50" maxlength="27" value="<%=rs_user("iban")%>" autocomplete=off></td>
	            </tr>
	            <tr>
	              <td valign="top">Salva questi dati nel profilo utente</td>
	              <td>
	                <input type="checkbox" name="aggiorna_dati_utente" ></td>
	            </tr>
	            <%
		            
			end if

            %>
	            
		<tr>
		  <td colspan="2" align="center" valign="top">
			  
		  	<div class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix">
				<div class="ui-dialog-buttonset">
				

			  
			  
			  
			  
			  
			  
			<%  if mostra_ddt then %>
					<button type="button" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only"  onClick="Valida_ddt_fattura('ddt')">
					<span class="ui-button-text">Crea D.d.T.</span>
					</button>
				
				
				
				<%
				end if
					
				if mostra_fattura  then %>
					<button type="button" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only"  onClick="Valida_ddt_fattura('fattura')">
					<span class="ui-button-text">Crea fattura <%=accompagnatoria%></span>
					</button>
				
				
			    <%end if%>
					<button type="button" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only"  onClick="Valida_ddt_fattura('bollettino')">
					<span class="ui-button-text">Crea bollettino</span>
					</button>
			    
			    

				</div>
			</div>
				 
				 
		    </td>
		</tr>
		</table>
		<input type="hidden" name="operazione_ddt"/>
		<input type="hidden" name="idord" value="<%=idord%>"/>
		<input type="hidden" name="idintestazione" value="<%=idintestazione%>">
		<input type="hidden" name="tipo-ddt" value="<%=tipo_ddt%>">
	</div>
		<script>
		$(function() {
			$("#causale").change(function(){
				var fattura=$( "#causale option:selected" ).attr("data-fattura");
				console.log("fattura:"+fattura);
				$("#da_fatturare").prop('checked', fattura=="True");
				
			})
		});
		
		
		
		</script>

	
<%
	call tab_dati_consegna(idconsegna, iduser)

 %>
	
		



	 
	
</div>

	</form>


              <script src="jquery/js/jquery.validate.min.js" type="text/javascript"></script>
		<script src="jquery/js/jquery.validate.min_it.js" type="text/javascript"></script>
		<script type="text/javascript" src="chk_piva_cf.js"></script>
        <script>
	      function enableSaveBtn(t_this){
	//alert("enableSaveBtn");
		console.log("enableSaveBtn");
	}

$(function() {	
	
$( "#tabs" ).tabs();
checkAnyFormFieldEdited();

$.validator.addMethod("iban", function (value, element) {
	
	if (!(/^([a-zA-Z0-9]{4} ){2,8}[a-zA-Z0-9]{1,4}|[a-zA-Z0-9]{12,34}$/.test(value))) {
		return false;
	}

	// check the country code and find the country specific format
	var iban = value.replace(/ /g, '').toUpperCase(); // remove spaces and to upper case
	var countrycode = iban.substring(0, 2);
	var bbancountrypatterns = {
		'AL': "\\d{8}[\\dA-Z]{16}",
		'AD': "\\d{8}[\\dA-Z]{12}",
		'AT': "\\d{16}",
		'AZ': "[\\dA-Z]{4}\\d{20}",
		'BE': "\\d{12}",
		'BH': "[A-Z]{4}[\\dA-Z]{14}",
		'BA': "\\d{16}",
		'BR': "\\d{23}[A-Z][\\dA-Z]",
		'BG': "[A-Z]{4}\\d{6}[\\dA-Z]{8}",
		'CR': "\\d{17}",
		'HR': "\\d{17}",
		'CY': "\\d{8}[\\dA-Z]{16}",
		'CZ': "\\d{20}",
		'DK': "\\d{14}",
		'DO': "[A-Z]{4}\\d{20}",
		'EE': "\\d{16}",
		'FO': "\\d{14}",
		'FI': "\\d{14}",
		'FR': "\\d{10}[\\dA-Z]{11}\\d{2}",
		'GE': "[\\dA-Z]{2}\\d{16}",
		'DE': "\\d{18}",
		'GI': "[A-Z]{4}[\\dA-Z]{15}",
		'GR': "\\d{7}[\\dA-Z]{16}",
		'GL': "\\d{14}",
		'GT': "[\\dA-Z]{4}[\\dA-Z]{20}",
		'HU': "\\d{24}",
		'IS': "\\d{22}",
		'IE': "[\\dA-Z]{4}\\d{14}",
		'IL': "\\d{19}",
		'IT': "[A-Z]\\d{10}[\\dA-Z]{12}",
		'KZ': "\\d{3}[\\dA-Z]{13}",
		'KW': "[A-Z]{4}[\\dA-Z]{22}",
		'LV': "[A-Z]{4}[\\dA-Z]{13}",
		'LB': "\\d{4}[\\dA-Z]{20}",
		'LI': "\\d{5}[\\dA-Z]{12}",
		'LT': "\\d{16}",
		'LU': "\\d{3}[\\dA-Z]{13}",
		'MK': "\\d{3}[\\dA-Z]{10}\\d{2}",
		'MT': "[A-Z]{4}\\d{5}[\\dA-Z]{18}",
		'MR': "\\d{23}",
		'MU': "[A-Z]{4}\\d{19}[A-Z]{3}",
		'MC': "\\d{10}[\\dA-Z]{11}\\d{2}",
		'MD': "[\\dA-Z]{2}\\d{18}",
		'ME': "\\d{18}",
		'NL': "[A-Z]{4}\\d{10}",
		'NO': "\\d{11}",
		'PK': "[\\dA-Z]{4}\\d{16}",
		'PS': "[\\dA-Z]{4}\\d{21}",
		'PL': "\\d{24}",
		'PT': "\\d{21}",
		'RO': "[A-Z]{4}[\\dA-Z]{16}",
		'SM': "[A-Z]\\d{10}[\\dA-Z]{12}",
		'SA': "\\d{2}[\\dA-Z]{18}",
		'RS': "\\d{18}",
		'SK': "\\d{20}",
		'SI': "\\d{15}",
		'ES': "\\d{20}",
		'SE': "\\d{20}",
		'CH': "\\d{5}[\\dA-Z]{12}",
		'TN': "\\d{20}",
		'TR': "\\d{5}[\\dA-Z]{17}",
		'AE': "\\d{3}\\d{16}",
		'GB': "[A-Z]{4}\\d{14}",
		'VG': "[\\dA-Z]{4}\\d{16}"
	};
	var bbanpattern = bbancountrypatterns[countrycode];
	// As new countries will start using IBAN in the
	// future, we only check if the countrycode is known.
	// This prevents false negatives, while almost all
	// false positives introduced by this, will be caught
	// by the checksum validation below anyway.
	// Strict checking should return FALSE for unknown
	// countries.
	if (typeof bbanpattern !== 'undefined') {
		var ibanregexp = new RegExp("^[A-Z]{2}\\d{2}" + bbanpattern + "$", "");
		if (!(ibanregexp.test(iban))) {
			return false; // invalid country specific format
		}
	}

	// now check the checksum, first convert to digits
	var ibancheck = iban.substring(4, iban.length) + iban.substring(0, 4);
	var ibancheckdigits = "";
	var leadingZeroes = true;
	var charAt;
	for (var i = 0; i < ibancheck.length; i++) {
		charAt = ibancheck.charAt(i);
		if (charAt !== "0") {
			leadingZeroes = false;
		}
		if (!leadingZeroes) {
			ibancheckdigits += "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".indexOf(charAt);
		}
	}

	// calculate the result of: ibancheckdigits % 97
	var cRest = '';
	var cOperator = '';
	for (var p = 0; p < ibancheckdigits.length; p++) {
		var cChar = ibancheckdigits.charAt(p);
		cOperator = '' + cRest + '' + cChar;
		cRest = cOperator % 97;
	}
	return cRest === 1;
}, "Inserire un IBAN valido");

		$.validator.addMethod("codice_fiscale", function (value, element) {
		result = true;
		chk=ControllaCF(value);
		if (chk != '' ) { 
			$.validator.messages.codice_fiscale = chk;
			result = false; 
		}
		if (value=="")
		{
			$.validator.messages.codice_fiscale = "Inserire il codice fiscale";
			result = false; 
		}
		return result;
	}, "");
	$.validator.addMethod("partita_iva", function (value, element) {
		result = true;
		chk=ControllaPIVA(value);
		if (chk != '' ) { 
			$.validator.messages.partita_iva = chk;
			result = false; 
		}
		if (($("#azienda").val()!="") && (value=="")){
			$.validator.messages.partita_iva = "Inserire la partita iva";
			result = false; 
		}
		return result;
	}, "");
	$.validator.addMethod("dateITA", function(value, element) {
    var check = false;
    var re = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
    if( re.test(value)) {
        var adata = value.split('/');
        var gg = parseInt(adata[0],10);
        var mm = parseInt(adata[1],10);
        var aaaa = parseInt(adata[2],10);
        var xdata = new Date(aaaa,mm-1,gg);
        var currentdata = new Date();   

        if (xdata.getFullYear() >= currentdata.getFullYear()-3) {
            check = false;
        }   
        else if ( ( xdata.getFullYear() === aaaa ) && ( xdata.getMonth() === mm - 1 ) && ( xdata.getDate() === gg ) ){
            check = true;
        } else {
            check = false;
        }
    } else {
        check = false;
    }
    return this.optional(element) || check;
}, "Inserisci una data corretta");

});

function Valida_ddt_fattura(cosa) {
	if (cosa == 'fattura') {
		$("#form_crea_ddtfattura").validate({
			ignore: "",
			rules: {
				porto: {required: true},
				imballo: {required: true},
				colli: {required: true},
				peso: {required: true},
				incaricato: {required: true},
				pagamento: {required: true},
				<%if italia then%>
				cf: {codice_fiscale: true},
				<%else%>
				cf: {required: true},
				<%end if%>
				vettore:{
				   required:{
					   depends: function(element){
					   var status = false;
						if (document.form_ddt.incaricato.value == '3' ) { 
							console.log("incaricato");
							   var status = true;
						   }
						   return status;
					   }
				   }
			   },
				iban:{
				   iban:{
					   depends: function(element){
					   var status = false;
					   var testo=$( "#pagamento option:selected" ).text();
						if(testo.indexOf('RI.BA.')>-1) { 
							console.log("riba");
							   var status = true;
						   }
						   return status;
					   }
				   }
			   },
				banca_appoggio:{
				   required:{
					   depends: function(element){
					   var status = false;
					   var testo=$( "#pagamento option:selected" ).text();
						if(testo.indexOf('RI.BA.')>-1) { 
							console.log("riba");
							   var status = true;
						   }
						   return status;
					   }
				   }
			   }
				<%if italia and (vtipocliente = "S" or vtipocliente = "D") then %>
				,piva: {partita_iva: true}
				 <%end if %>
			}
		});
	}
	if (cosa == 'ddt') {
		$("#form_crea_ddtfattura").validate({
			ignore: "",
			rules: {
				porto: {required: true},
				imballo: {required: true},
				colli: {required: true},
				peso: {required: true},
				incaricato: {required: true},
				<%if rs_user("stato_estero")<>1  then %>
				cf: {codice_fiscale: true},
				<%end if%>
				vettore:{
				   required:{
					   depends: function(element){
					   var status = false;
						if (document.form_ddt.incaricato.value == '3' ) { 
							   var status = true;
						   }
						   return status;
					   }
				   }
			   }
			}
		});
	}
	if ($("#form_crea_ddtfattura").valid()) {
		if (cosa == 'bollettino') {
			if (document.form_ddt.porto.value == '2') {
				alert("Impossibile generare bollettino.");
				return false;
			} <%
			if cdbl(trasporto) = 0 then %>
			var r = confirm('SEI SICURO CHE NON CI SONO SPESE TRASPORTO ?');
			if (r == false) {
				return;
			}

			<% end if %>
			dati="&colli="+$("#colli").val()+"&peso="+$("#peso").val()+"&dimensione="+$("#dimensione").val()+"imballo="+$("#imballo").val()+"&porto="+$("#porto").val();
			top.location.href="bollettino.asp?idord=<%=idord%>"+dati;
			return true;
		}
		
		document.form_ddt.operazione_ddt.value = 'crea_' + cosa;
		invia_dati();
	}		

}

function invia_dati(){	
	var dati=$( "#form_crea_ddtfattura" ).serialize();
	var spettanze="";
	//var parziale='<%=parziale%>';
	var parziale=$("#parziale").val();
	if ( $( "#form-avanz-spettanze" ).length && parziale=='parziale') {
			
			dati = dati.replace(/(&idord=)\d+/, '');
			console.log("esiste");
			//var spettanze = $("#form-avanz-spettanze").serialize();
			
			spettanze = $("#form-avanz-spettanze").serialize().replace(/[^&]+=&/g, '').replace(/&[^&]+=$/g, '')
			console.log("serialize SPETTANZE:"+spettanze);
			dati+="&spettanze=si&"+spettanze;
			console.log("ALLEGASTO SPETTANZE:");
	}
	console.log(dati);
	
	$.ajax({
		url     : "pag_adm_dialog_creafattura.asp",
		type    : "post",
		cache: false,
		dataType: 'json',
		data	: dati,
		
		success: function(data){
			
			
					if (data.cosa=="F"){
						location.href="pag_adm_fatture.asp?idfat="+data.idfat;
	
				}
				else{
						location.href="pag_adm_ddt.asp?idddt="+data.idddt;
					}

			
			
			console.log (data.success+data.message);
			if ($( "#multitabs" ).length|| false){
				console.log("Gestisco tabs");
				$("#pulsante_crea_ddt_fattura").hide();
				var tabs = $( "#multitabs" ).tabs();
				var ul = tabs.find( "ul" );
				if (data.cosa=="F"){
					$( "<li><a href='pag_adm_fatture.asp?tab=1&idfat="+data.idfat+"'><strong>Fattura "+data.nfat+"</strong> del "+data.data+"</a></li>" ).appendTo( ul );
				}
				else{
					$( "<li><a href='pag_adm_ddt.asp?tab=1&idddt="+data.idddt+"'><strong>DDT "+data.nddt+"</strong> del "+data.data+"</a></li>" ).appendTo( ul );
				}
				tabs.tabs( "refresh" );
				var last = $(".ui-tabs-nav li").last().index();
				//Imposto tab creato come attivo
				$("#multitabs").tabs("option", "active", last);
				
				$("#dialog").dialog("close");
			}else{
				console.log("location.href");
				if (data.cosa=="F"){
						location.href="pag_adm_fatture.asp?idfat="+data.idfat;
	
				}
				else{
						location.href="pag_adm_ddt.asp?idddt="+data.idddt;
					}
					
				}
				
			
			//Nascondo pulsante crea ddt e fattura
			//Creo nuovo tab
		}
		,error: function(xhr, textStatus, error){
			toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
			toastr.error("Errore nell'aggiornamento dei dati");
			var txt="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
			txt+="Errore in "+location.href+"<br>";
			txt+="<br>"+xhr.statusText;
			txt+="<br>"+xhr.responseText;
			txt+="<br>textStatus:"+textStatus;
			txt+="<br>error:"+error;
			invia_errore("Errore in creazione ddt o fattura()",txt );
		}
	});
}


		</script>

<%
rs_user.close
set rs_user=nothing
	
	set rs_ordine=Nothing
	call connclose()
	call CheckConnChiusa()

%>