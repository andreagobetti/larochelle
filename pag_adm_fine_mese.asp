<%
'Verifica chiusure 04_12_2015
%>
<%t_inizio=timer()%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/chk_piva_cf.asp" -->
<!--#include virtual="/paginazione.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->

<%
testo_esito=""

'add2log queryeform(),0
if session("idadmin") = "" then call login()
if request("idord")<>"" then idord=request("idord")
mese=request.form("mese")
cercain=request("cercain")
cerca=request("cerca")


if request.form("operazione")="fatture_fm" then

	call fatture_fm()
end if
Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class

if oper="" then
	oper=request("oper")
	select case oper
		case "annulla"
			oper="list"
		case "new"
			oper="new"
		case "aggiungi"
			oper="add"
		case "modifica"
			oper="update"
		case else
			oper="list"
			if request.form("modifica")<>"" then
				oper="view"
				idord=request.form("modifica")
			end if
			if request("idord")<>"" then
				 oper="view"
				 idord=request("idord")
			end if
			if isnumeric(request.form("cerca")) and (request.form("cercain")="" or request.form("cercain")="idord") and request.form("cerca")<>"" then
				oper="view"
				nord=request.form("cerca")
			end if
	end select
end if
if oper="add" or oper="update" then
	'controlli
end if


display_none="display:none;"
if request("anno")="" then
	anno=year(date())
else
	anno=request("anno")
end if
%>
<!--#include virtual="/regioni.inc" -->
<!--#include virtual="/sub_head_adm.asp" -->

  <style type="text/css">
span.red {
	color: red;
}
  </style>
  </head>
  <body>
  <div id="wrap">
    <div id="header"> <%=titolo_top%> 
      <!-- Box CORPO INIZIO-->
      <%barra=4%>
      <div id="barra_fissa"> 
        <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">DDT per fatturazione a fine mese</a></div>
      </div>
      <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" onSubmit="return Validator(this)">
        <!-- Box CORPO INIZIO-->
        <input type="hidden" name="operazione">
        
        <%
	        
call report_errore()


if testo_esito<>"" then
	
	
	response.write testo_esito
	
	
end if
if oper="list" then



				

sql="select oddt.idord as idddt, oddt.nord as nddt, oddt.iduser, oddt.data, ordini_fatture.idord, ddt.causale, ddt.da_fatturare, ddt.tipo_ddt, ddt.sub_idord, i.* , ordini.tipopagamento, utenti.spese_0, utenti.pag_accordato, utenti_clienti.FatturaPACodiceDestinatario, utenti_clienti.stato_estero, utenti.banca_appoggio, utenti.banca_appoggio, utenti.iban, utenti_clienti.tipo , ordini.idord, ordini.nord,ordini.anno as annoordine, ordini.trattamento_iva_ordine, ordini.trasporto from (select idord,nord,data,iduser,tipo_documento from ordini where tipo_documento='ddt') as oddt left join ordini_fatture on oddt.idord = ordini_fatture.idord inner join ddt on oddt.idord = ddt.idord inner join utenti on oddt.iduser = utenti.iduser inner join utenti_intestazioni i on utenti.idintestazione = i.id left join ordini on ddt.sub_idord = ordini.idord left join utenti_clienti on utenti.iduser = utenti_clienti.iduser  where ordini_fatture.idord is null and ddt.da_fatturare=1  ORDER BY oddt.iduser, ordini.tipopagamento, ordini.nord, oddt.data"


				where=""
				order=""
				PaginazioneString=""
				'response.write "finemese"
				strsql=sql
				
		PaginazioneString="&cercain="&cercain&"&cerca="&cerca&"&anno="&anno&"&cerca_hidden="&request("cerca_hidden")


				strsql=sql
				'giorno=day(date())
				'if giorno>20 then
				'	data_fatture=date()
				'else
				'	data_fatture = DateSerial(Year(date()),Month(date()),1) 
				'	data_fatture = data_fatture - 1 
				'end if
				data_fatture=date()
		%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1">

          <tr>
            <td colspan="6">
              <label><input type="checkbox" id="checkAll" /> Seleziona tutti</label> <span style="float: right;"><input type="button" name="Genera_ffm" value="Genera fatture" style="FONT: 12px;" onClick="conferma1()"> in data <input type="text" value="<%=data_fatture%>" name="data_fatture" id="data_fatture"></span>
              </td>
          </tr>
          <%
call paginazione_start(sql,100,"access")

if objPagingRS.eof then %>
<tr>
	<td colspan="6" align="center"><strong>Nessun ddt valido per fattura a fine mese</strong></td></tr>


<%
	end if


iduser=0
v_tipopagamento=0
totale_ordini=0


totale_imponibile=0

idord_tmp=0
Do While iRecordsShown < iPageSize And Not objPagingRS.EOF


		if (objPagingRS("iduser")<>iduser) or (objPagingRS("tipopagamento")<>v_tipopagamento) then	'E' cambiato l'utente- creo nuova fattura
			fattura_ok=true
  	        v_tipo_cliente=objPagingRS("tipo")
			%>
          <tr bgcolor="#E5E5E5" style="border-top:2px solid;">
	          <td colspan="1"><strong>Fattura a:</strong></td>
  	          <td colspan="4">
	  	          <b><a href="pag_adm_user.asp?iduser=<%=objPagingRS("iduser")%>"><%=denominazione(objPagingRS("nome"),objPagingRS("cognome"),objPagingRS("azienda"))%></a></b><br>
	              <%=objPagingRS("Indirizzo")%> <br>
	              <%=objPagingRS("citta")%> (<%=objPagingRS("provincia") & ") - " & regione(objPagingRS("regione"))%>
              </td>
  	          <td colspan="">
	  	          <%
		  	        chk_iban=true
		  	        tipopagamento_txt=""
		  	        'v_tipo_cliente=objPagingRS("tipo")
		  	        
		  	        %>
		  	    Tipo cliente: <%'=TipoCliente(v_tipo_cliente)%><br>
		  	        <%
				v_tipopagamento=objPagingRS("tipopagamento")
				chkpiva="" 
		  	    chkCF=""
				if objPagingRS("spese_0")=1 or isnull(v_tipopagamento) then
					spese_bancarie_v=0
				else
					spese_bancarie_v=metodipagamento.spese_bancarie(v_tipopagamento)
				end if
				
				if v_tipopagamento=0 or isnull(v_tipopagamento) then
					
					tipopagamento_txt="<span class='red'>Manca tipo pagamento</span>"
					if objPagingRS("pag_accordato")>0 then
						v_tipopagamento=objPagingRS("pag_accordato")
						tipopagamento_txt=tipopagamento_txt&" applicato pagamento concordato <br>"&metodipagamento.descrizione(v_tipopagamento)

					else
						tipopagamento_txt=tipopagamento_txt&"<span class='red'> Pagamento concordato in anagrafica cliente non specificato</span>"
						fattura_ok=false
					end if
					tipopagamento_txt=tipopagamento_txt&"<br>"
				else
					tipopagamento_txt=metodipagamento.descrizione(v_tipopagamento)
				end if		
				
				if metodipagamento.campo(v_tipopagamento,"richiedi_banca")=1 then
				'if instr(lcase(v_tipopagamento),"bonifico")>0 then
					chk_iban=true
				else
					chk_iban=false
				end if
				
				  	          
				FatturaPACodiceDestinatario=objPagingRS("FatturaPACodiceDestinatario")
				if not isnull(FatturaPACodiceDestinatario) or FatturaPACodiceDestinatario<>"" then
					codicedestinatario=FatturaPACodiceDestinatario
				else
					codicedestinatario="<span class=""rosso"">CODICE DESTINATARIO MANCANTE</span>"
					'fattura_ok=false
				end if 
	             
				banca_appoggio_v=objPagingRS("banca_appoggio")
				if (banca_appoggio_v="" or isnull(banca_appoggio_v)) and chk_iban then
					banca_appoggio_v= "<span class='red'>Manca banca appoggio</span>"
					fattura_ok=false
				end if
				
				iban=objPagingRS("iban")
				if (iban="" or isnull(iban)) and chk_iban then
					iban= "<span class='red'>Manca IBAN</span><br>"
					fattura_ok=false
				end if
				if objPagingRS("stato_estero")=1 then
				
				
				elseif (v_tipo_cliente<>"A" and v_tipo_cliente<>"P") then
					if objPagingRS("piva")="" or isnull(objPagingRS("piva")) then
						response.write "<span class='red'>Manca partita iva</span><br>"
						fattura_ok=false
					else	
						chkpiva=ControllaPIVA(objPagingRS("piva"))
						response.write "Partita iva: "&objPagingRS("piva")&"<br>"
						if chkpiva<>"" then
							response.write chkpiva&"<br>"
							fattura_ok=false
						end if
					end if
				end if
				if objPagingRS("stato_estero")=1 then
				
				elseif objPagingRS("cf")="" or  isnull(objPagingRS("cf")) then
					response.write "<span class='red'>Manca codice fiscale</span><br>"
					fattura_ok=false
				else
					response.write "Codice fiscale: "&objPagingRS("cf")&"<br>"
					if v_tipo_cliente="A" then
						
					elseif objPagingRS("cf")<>objPagingRS("piva") then
						chkcf=ControllaCF(objPagingRS("cf"))
						if chkCF<>"" then 
							response.write chkCF&"<br>"
							fattura_ok=false
						end if							
					end if
				end if
		  	          %>
				Pagamento:<%=tipopagamento_txt%><br>
				Spese bancarie:<%=spese_bancarie_v%><br>
				Trattamento IVA:<%=trattamento_iva(objPagingRS("trattamento_iva_ordine"))%><br>
				<%if objPagingRS("trattamento_iva_ordine")=10 then%>
				Codice univoco:<%=codicedestinatario%><br>
				<%end if %>
				<%if chk_iban then %>
				Banca appoggio cliente:<%=banca_appoggio_v%><br>
				IBAN cliente:<%=iban%>
				<%end if %>
	              
	              
	              <%if fattura_ok=false then response.write "<br><span class=""rosso""><b>FATTURA NON GENERABILE PER DATI MANCANTI</b></span>"%>
			</td>
		</tr>
 		<tr style="border-top:1px solid;">
            <td align="center" bgcolor="#E5E5E5" colspan="2">DDT</td>
            <td align="center" bgcolor="#E5E5E5" colspan="1">Data</td>
            <td align="center" bgcolor="#E5E5E5" colspan="1">N&deg; Ordine</td>
            <td bgcolor="#E5E5E5" colspan="2">Tot. ordine</td>
		</tr>
	          <%
		end if
		'totale_ordini=totale_ordini+cdbl(objPagingRS("totale"))
		tipo_ddt=objPagingRS("tipo_ddt")
    	if tipo_ddt="parziale" then
        	sub_idord=objPagingRS("idddt")	
        else
        	sub_idord=objPagingRS("sub_idord")
		end if
		
		'Recupero dati singolo ddt/ordine
		set rs_ordine=conn.execute("select * from ordini where idord="&sub_idord)
		totale_merce_ordine=cdbl(rs_ordine("totale_merce_ordine"))
		if objPagingRS("idord")<>idord_tmp then
			if tipo_ddt="parziale" then
				trasporto=cdbl(objPagingRS("trasporto"))
			else
				trasporto=cdbl(rs_ordine("trasporto"))
			end if
		else
			trasporto=0
		end if
		idord_tmp=objPagingRS("idord")
		'pagamento_ordine=rs_ordine("tipopagamento")
		if utente_andrea then %>
		
		
		<tr>
			<td colspan="6">
				tipo_ddt:<%=tipo_ddt%>  sub_idord:<%=sub_idord%> 
				
			</td>
		</tr>
		<%end if%>
		<tr class="coprobox" <%=colore%> style="border-top:1px solid;">
	        <%
		      if fattura_ok=true then
		      	readonly=""
		      	v_class=""
		      else
		      	v_readonly="disabled"
			  end if
		    %>
            <td align="center" colspan="1">
	        <%
		      if fattura_ok=true then
			      %>
			       <input type="checkbox" class="checkboxenabled" name="includi" value="|<%=objPagingRS("idddt")%>|" >
			      <%
		      else
			      %>
			       <input type="checkbox" disabled >
			      <%
			  end if
		    %>
			</td>
			<td>
             <%if objPagingRS("idddt")>0 then
            	
				
				totale_imponibile=totale_imponibile+totale_merce_ordine+trasporto
            %>
	            <a href="pag_adm_ddt.asp?idddt=<%=objPagingRS("idddt")%>">D <%=objPagingRS("nddt")%></a> <%=get_tipo_ddt(tipo_ddt)%>
	        <%end if%>
		        
	        </td>    
            <td align="center"   colspan="1">
	            <%=formatdatetime(objPagingRS("data"),2)%>
            </td>
            <td align="center" style="border-top:1px solid;" colspan="2">
	            <%if tipo_ddt<>"no_ordine" then %>
	            <a href="pag_adm_ordini.asp?idord=<%=objPagingRS("idord")%>"><%=objPagingRS("nord")&anno_se_passato(objPagingRS("annoordine"))%></a>
	            <%end if %>
            </td>
            <td align="right" colspan="2">
	            <%="Totale merce: "&formatcurrency(totale_merce_ordine)%><br>
				
	            <%if trasporto>0 then response.write "Trasporto: "&formatcurrency(trasporto)&"<br>" end if%>
	            <%="tipo pagamento: "&metodipagamento.descrizione(rs_ordine("tipopagamento"))%>
	            </td>
          </tr>

		  
		<% 	
		
  		iduser=objPagingRS("iduser")
		v_tipopagamento=objPagingRS("tipopagamento")

		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
		visualizza_totale=false
		if objPagingRS.eof then
			visualizza_totale=true
		else
			if (objPagingRS("iduser")<>iduser) or (objPagingRS("tipopagamento")<>v_tipopagamento) then			
				visualizza_totale=true
			end if	
		end if
		
		
		
		if visualizza_totale then
			totale_ordini=totale_ordini+spese_bancarie_v*(1+iva(date())/100)
			%>
          <tr bgcolor="#E5E5E5" style="border-bottom:2px solid; border-top:1px solid;">
	          <td colspan="3" >
		          <b>Totale imponibile fattura</b>
	          </td>
	          <td colspan="3" align="right">
		          <b><%=formatcurrency(totale_imponibile,2)%></b>
	          </td>
          </tr>
			
			<%
			totale_imponibile=0
			trasporto=0
		end if
		
		
		
		
		
	Loop	
	set rs_ordine = Nothing
	call paginazione_end(6,true)
	
%>
        </table>
    <%else
		  end if%>
      <!-- Box CORPO FINE--> 
    </div>
<%
	txt_timer=txt_timer&"T fine:"&FormatNumber(timer() - t_inizio, 2)&" "
%>
<div id="footer"><%=txt_timer%></div>
<!--#include virtual="/pag_adm_footer_inc.asp" -->
  </div>
  
<%if oper="list" then%>


<script src="jquery/ui/i18n/jquery.ui.datepicker-it.min.js"></script>
<script>
document.title = "DDT per fatturazione a fine mese";

$(document).ready(function(){
    $("#checkAll").change(function () {
	    $("input:checkbox.checkboxenabled").prop('checked', $(this).prop("checked"));
	});
			$( "#data_fatture" ).datepicker();

	
})

function conferma1()
{
var r=confirm('Confermi la generazione delle fatture?');
posizione_errore="conferma1";

if (r==true){ 
	document.form1.operazione.value='fatture_fm';
	document.form1.submit();
  }
else
  {
  return false;
  }
}


</script>
<%end if%>
</body>
</html>

<%


sub fatture_fm()
	dim metodipagamento
	Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class

	'response.write queryeform()
	'response.end
	ddt_inclusi=request("includi")
	testo="Generazione fatture fine mese"&vbcrlf
	iduser=0
	v_tipopagamento=0
	'sql="select ordini.*, ddt.IDddt, ddt.causale, ddt.sub_idord, utenti.idintestazione, utenti.iban, utenti.banca_appoggio, utenti.pag_accordato, utenti.piva, utenti.cf, utenti.spese_0, utenti.trattamento_iva, utenti.idbanca FROM ((ordini INNER JOIN ddt ON ordini.idord = ddt.idord) LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) INNER JOIN utenti ON ordini.iduser = utenti.iduser WHERE (ordini_fatture.idord Is Null) and causale=1 "
			
	'if request("cerca_iduser")<>"" then sql=sql&" and  ordini.iduser="&request("cerca_iduser")
	'sql=sql&" ORDER BY ordini.iduser, ordini.tipopagamento"


	sql="select oddt.idord as idddt, oddt.nord as nddt, utenti.fattura_sp, oddt.iduser, oddt.data, ordini_fatture.idord, ddt.causale, ddt.da_fatturare, ddt.tipo_ddt, i.* , ordini.tipopagamento, utenti.spese_0,utenti.pag_accordato, utenti_clienti.FatturaPACodiceDestinatario, utenti.banca_appoggio, utenti.banca_appoggio, utenti.idbanca, utenti.iban , ordini.idord, ordini.nord,ordini.anno, ordini.trattamento_iva_ordine, oddt.idintestazione from (select idord,nord,data,iduser,tipo_documento, idintestazione, fattura_sp from ordini where tipo_documento='ddt') as oddt left join ordini_fatture on oddt.idord = ordini_fatture.idord inner join ddt on oddt.idord = ddt.idord inner join utenti on oddt.iduser = utenti.iduser inner join utenti_intestazioni i on utenti.idintestazione = i.id left join ordini on ddt.sub_idord = ordini.idord left join utenti_clienti on utenti.iduser = utenti_clienti.iduser  where ordini_fatture.idord is null and ddt.da_fatturare=1  ORDER BY oddt.iduser, ordini.tipopagamento"


	data=request.form("data_fatture")

	set rs_ordini=conn.execute(SQL)
	'testo=testo&"Trovati "&rs_ordini(0) & " ordini<br>"
	do until rs_ordini.eof
		'Controllo se l'ordine è spuntato
		idddt=rs_ordini("idddt")
		if instr(ddt_inclusi,"|"&idddt&"|")>0 then
			
			
			if (rs_ordini("iduser")<>iduser) or (rs_ordini("tipopagamento")<>v_tipopagamento) then	'E' cambiato l'utente- creo nuova fattura
	
				fattura_ce=0
				fattura_pa=0
	
				if rs_ordini("trattamento_iva_ordine")=10 and  DateDiff("d","31/03/2015",date())>0 then 'dopo il 31/03/2015 then
					fattura_pa=1
				elseif rs_ordini("trattamento_iva_ordine")=11 then
					fattura_ce=1
				end if	
				
				fattura_sp=rs_ordini("fattura_sp")

				Nfat=max_fatt(fattura_pa,fattura_sp,year(data))
				
				v_tipopagamento=rs_ordini("tipopagamento")
				
				'rs.close
				sql="select * from fatture"
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open sql, conn, 3, 3
				rs.addnew
				RS("Nfat")=Nfat
				
				rs("data")=data
				rs("anno")=year(data)
				rs("fine_mese")=1
				rs("tipo_fattura")="ddt"
				rs("iduser")=rs_ordini("iduser")
				rs("idintestazione")=rs_ordini("idintestazione")
				rif_idord=rs_ordini("idord")
				'if rif_idord=0 then
				'	rif_idord=rs_ordini("idord")
				'end if
				rs("idord")=rif_idord
				rs("trattamento_iva")=rs_ordini("trattamento_iva_ordine")
				rs("pa")=fattura_pa
				rs("fattura_ce")=fattura_ce

				rs("fattura_sp")=fattura_sp
				if fattura_sp=1 then
					rs("trattamento_iva")=10
				end if
				idbanca=rs_ordini("idbanca")
				'cerco banca predefinita
				if idbanca="" or idbanca=0 then
					set rsb=conn.execute("select * from banche where predefinita=1")
					idbanca=rsb("idbanca")
					rsb.close
					set rsb=nothing
				end if
	
				rs("idbanca")=idbanca
				spese_bancarie_v=0
				scadenze=0
	
				if rs_ordini("spese_0") or isnull(v_tipopagamento) then
					spese_bancarie_v=0
				else
					spese_bancarie_v=metodipagamento.spese_bancarie(v_tipopagamento)
				end if
				if isnull(v_tipopagamento) or v_tipopagamento=0 then
					v_tipopagamento=rs_ordini("pag_accordato")
				end if
				
				rs("pagamento")=v_tipopagamento
				rs("spese_bancarie")=spese_bancarie_v
				rs("causale")=1 'VENDITA
				sql="select *, utenti_clienti.FatturaPACodiceDestinatario from utenti left join utenti_clienti on utenti.iduser = utenti_clienti.iduser where utenti.iduser="&rs_ordini("iduser")
				set rs_user=conn.execute(SQL)
				bb=rs_user("banca_appoggio")
				rs("banca_appoggio")=bb
				rs("iban")=rs_user("iban")
				
				
				
				'if rs_user("Pag_accordato")>=10 and rs_user("Pag_accordato")<=12 then rs("spese_bancarie")=3.7
				
				rs("creata_da")=session("nominativo")
				rs("eliminato")=0
				rs("fine_mese")=1
				rs("aliquota_iva")=iva(date())
				'codicedestinatario=rs_ordini("codicedestinatario")
				'FatturaPACodiceDestinatario=r
				
				codiceunivoco=rs_user("FatturaPACodiceDestinatario")
				
				if inStr(codiceunivoco,":")>0 then
					
					codiceunivoco=left(codiceunivoco,inStr(codiceunivoco,":")-1)
					end if
				
				rs("codicedestinatario")=codiceunivoco
				rs("sconto_ordine")=0
				rs("totale_fattura")=0
				rs("bk_totale_fattura")=0
				rs("totale_merce")=0
				rs("imposta")=0
			
				rs.update
				IDfat=get_last_id("fatture")
				pagamento=rs("pagamento")
				rs_user.close
				rs.close
				
			end if
			'if rs_ordini("tipo_ddt")="1" then
				rif_idord=rs_ordini("idddt")
			'else
			'	rif_idord=rs_ordini("sub_idord")
			'end if
			
			'collego dettaglio ordine
			sql="select * from ordini_fatture"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql, conn, 3, 3
			rs.addnew
			rs("idord")=rif_idord
			rs("idfat")=IDfat
			rs.update
			rs.close
			iduser=rs_ordini("iduser")
			v_tipopagamento=rs_ordini("tipopagamento")
			ddt_escluso=false
		else
			ddt_escluso=true
		end if
		if ddt_escluso then 
			testo=testo&"DDT [ddt="&idddt&"]"&rs_ordini("nddt")&"[/ddt] saltato perch&egrave; non selezionato<br>"
		end if
		rs_ordini.movenext
		if not rs_ordini.eof then
			if ((rs_ordini("iduser")<>iduser) or (rs_ordini("tipopagamento")<>v_tipopagamento)) and iduser>0 then
				testo=testo&"Creata [fattura="&IDfat&"]"&nfat&"[/fattura] di"
				
				testo_esito=testo_esito&"Creata fattura <a href=""pag_adm_fatture.asp?odfat="&IDfat&""">"&nfat& get_fattura_pa(fattura_pa,fattura_ce)&"</a> di importo "&importo&"<br>"
				importo=calcola_totale_fattura(IDfat)
				testo=testo& " importo "&importo
				
				testo=testo& ", create "&n&" scadenze"
				n_incassi=collega_incassi (0,IDfat)
				testo=testo& ", collegati "&n_incassi&" incassi<br>"
				iduser=0
				testo_esito=testo_esito&"Creata fattura <a href=""pag_adm_fatture.asp?odfat="&IDfat&""">"&nfat& get_fattura_pa(fattura_pa,fattura_ce)&"</a> di importo "&importo&"<br>"
			end if
		else
			if iduser>0 then
				testo=testo&"Creata [fattura="&IDfat&"]"&nfat&"[/fattura] di"
				importo=calcola_totale_fattura(IDfat)
				testo=testo& " importo "&importo
				
				testo=testo& ", create "&n&" scadenze"
				n_incassi=collega_incassi (0,IDfat)
				testo=testo& ", collegati "&n_incassi&" incassi<br>"
				testo_esito=testo_esito&"Creata fattura <a href=""pag_adm_fatture.asp?odfat="&IDfat&""">"&nfat& get_fattura_pa(fattura_pa,fattura_ce)&"</a> di importo "&importo&"<br>"
			end if
		end if
	loop
	set rs = nothing
	rs_ordini.close
	set rs_ordini=nothing
	testo=testo&"Fine"
	add2log testo,2
	if utente_andrea then 
		response.write testo
	end if
	set metodipagamento = nothing
end sub

call CheckConnChiusa()
%>