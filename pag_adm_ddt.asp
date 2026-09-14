<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/paginazione.asp" -->
<!--#include virtual="/classeMagazzino.asp" -->
<!--#include virtual="/classeOrdine.asp" -->

<%
dim oper
dim main_page
if not utente_admin then call login()
idddt=request("idddt")
oper=lcase(request("oper"))
'add2log queryeform(),0
call via_senza_permesso("B2")


dim ccausale_ddt
if oper="update" then
	call ddt_update()
elseif oper="aggiorna-reso" then
	call aggiorna_reso()
end if
if request("tab")<>"" then
	main_page=false
	call carica_tab(request.querystring("tab"),idddt)
	call connclose()
	set ccausale_ddt = Nothing
	response.end
end if
'response.write request.form
if request.querystring("elimina_ddt")<>"" then
	idddt=request.querystring("elimina_ddt")
	Set rs_ordine = Server.CreateObject("ADODB.Recordset")
	rs_ordine.Open "select * FROM ordini where idord="& idddt , conn, 3,2
	data=rs_ordine("data")
	nddt=rs_ordine("nord")
	Set rs_ddt = Server.CreateObject("ADODB.Recordset")
	rs_ddt.Open "select * FROM ddt where idord="& rs_ordine("idord") , conn, 3,2
	log_txt="tipo_documento:"&rs_ordine("tipo_documento")&" idord:"&rs_ordine("idord")
	sub_idord=rs_ddt("sub_idord")
	tipo_ddt=rs_ddt("tipo_ddt")
	log_txt=log_txt&" idddt:"&rs_ddt("idddt")&" sub_idord:"&sub_idord&" tipo_ddt:"&tipo_ddt
	'Azzero campo consegnato
	if tipo_ddt="completo" then
		'Sconbsegno tutto
		'call consegna_tutto(sub_idord,false)
		log_txt=log_txt&" consegna_tutto(false) "
	end if
	if tipo_ddt="completo" or tipo_ddt="parziale" then
		call ripristina_rollback(rs_ddt("idrollback"))
	end if


	'Tolgo verde dall'ordine
	call imposta_verde(sub_idord,false)
	'Elimino righe da ordini dett, servirebbe solo per tipo 1 o 2
	conn.execute "delete from ordini_dett where idord="&rs_ordine("idord"),num

	log_txt=log_txt&" eliminati "&num&" ordini_dett"
	rs_ddt.delete
	rs_ddt.Close
	set rs_ddt = Nothing
	rs_ordine.delete
	rs_ordine.Close
	set rs_ordine = Nothing
	if tipo_ddt="completo" or tipo_ddt="parziale" then
		nord=conn.execute("select nord from ordini where idord="&sub_idord)(0)
		txt_ordine="per [ordine="&sub_idord&"]"&nord&"[/ordine]"


	end if
	add2log "Eliminato [ddt="&idddt&"]"&pre_fattura(data)&nddt&"/"&year(data)&"[/ddt] "&txt_ordine&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&log_txt,3
	idddt=""
	oper="list"
	call connclose()
	response.redirect questofile
end if
if request.form("operazione")="crea_fattura" then
	idord=request.form("idord")
	idfat=crea_fattura("da_ddt")

	if idfat>0 then
		response.redirect "pag_adm_fatture.asp?idfat="&idfat
	else
		script_errore ="alert (""Esiste già una fattura per questo ordine"");"
		oper="edit"
	end if
end if


if oper="" then
	if request("idord")<>"" then
			 oper="view"
			 idord=request("idord")
	elseif request.form("view")<>"" then
		oper="view"
		idddt=request.form("view")
	elseif request.querystring("idddt")<>"" then
		oper="view"
		idddt=request("idddt")
	elseif isnumeric(request.form("cerca")) and request.form("cerca")<>"" and request.form("cercain")="" then
		oper="view"
		nddt=request.form("cerca")
		idddt=0
	else


		oper="list"
	end if

end if




display_none="display:none;"
%>
<!--#include virtual="/regioni_inc.asp" -->
<!--#include virtual="/sub_head_adm.asp" -->
<script type="text/javascript" language="javascript">
<%=script_errore%>
function showhidearticoli()
{
  var elem = document.getElementById('articoli_originali')

  if (elem.style.display == 'none'){
  document.getElementById('pulsante_articoli').value = 'Nascondi articoli originali';
    elem.style.display = 'block';}
  else{
  document.getElementById('pulsante_articoli').value = 'Mostra articoli originali';
    elem.style.display = 'none';}
}
<%if oper="list" then
convalida="return Valida_ricerca()"%>
function Valida_ricerca()
{
	if ((document.form1.cercain.value == 'idord' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il numero ordine nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	if ((document.form1.cercain.value == 'iduser' )&&(document.form1.cerca.value == '')) {
		alert("Inserire id cliente nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	if ((document.form1.cercain.value == 'idddt' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il numero D.d.T. nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	if ((document.form1.cercain.value == 'azienda' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il nome azienda nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	document.form1.submit();
}
<%end if%>
</script>
<style type="text/css">
<!--
.piccolo {
	font-size: 9px;
}
.pulsanti {
	padding: 3px;
	overflow: visible;
}
-->
</style>
           <style>
	.ui-autocomplete-loading {
		background: white url('images/ui-anim_basic_16x16.gif') right center no-repeat;
	}
	.ui-autocomplete {
		max-height: 300px;
		overflow-y: auto;
		/* prevent horizontal scrollbar */
		overflow-x: hidden;
	}
	/* IE 6 doesn't support max-height
	 * we use height instead, but this forces the menu to always be this tall
	 */
	* html .ui-autocomplete {
		height: 300px;
	}
  </style>

<script>
	$(function() {
		$( "#cerca" ).autocomplete({
			source: "ajax_function.asp?autocomplete=utenti",
			cache: false ,
			minLength: 2,
			select: function( event, ui ) {
				$('#cerca_hidden').val(ui.item.id);
				$('#form1').submit();
			}
		});
	});
</script>
</head>
<body>
<div id="wrap">
<div id="header">
<%=titolo_top%>
		<%barra=5%>	<div id="barra_fissa">  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">DDT</a></div></div>

          <!-- Colonna CORPO INIZIO-->
          <%
if oper="list" then
	sql="select o.*, i.nome, i.cognome, i.azienda, i.citta, i.provincia, u.tipologia, d.causale, d.da_fatturare, d.sub_idord, d.tipo_ddt,  d.reso_tutto, d.incaricato, d.vettore, sommadiimporto, fatture.idfat, fatture.nfat, fatture.pa,fatture.fattura_ce, o2.idord as idordine, o2.nord as nordine , o2.anno as annoordine from (select ordini.* from ordini where tipo_documento='ddt') o left join utenti_intestazioni i on o.idintestazione = i.id inner join utenti u on o.iduser = u.iduser inner join ddt d on o.idord = d.idord LEFT JOIN (select incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idord ) sommadiincassi ON o.idord = sommadiincassi.idord  left join ordini_fatture on d.idord = ordini_fatture.idord left join fatture on ordini_fatture.idfat = fatture.idfat left join ordini o2 on d.sub_idord = o2.idord"

		cercain=request("cercain")
		cerca=request("cerca")
		if cerca<>"" and isnumeric(cerca) and cercain="" then cercain="idddt"
		iPageSize=Application("grec4page")
		if request("anno")="" then
			anno=""
		else
			anno=request("anno")
		end if
		if anno="" then
			where_anno  = ""
		else
			where_anno  = "  year(o.data)="&anno&" "
		end if
		'response.write where_anno
		ordine=" order by anno desc, nord desc"
		'-----------sezione ricerca inizio
		if request.form("cerca_hidden")<>"" then cercain="hidden"
		if cercain<>"" then
			select case cercain
			case "idord"
				where=" ddt.idord="&cerca
			case "iduser"
				where=" o.iduser="&cerca
			case "nddt"
				where=" nddt="&cerca
			case "dafatt"
				where=" nfat is null"
			case "azienda"
				where=" InStr([ordini.azienda],'" & cerca& "')>0"
			case "novendita"
				where=" ddt.causale<>1"
			case "hidden"
				where=" o.iduser="&request.form("cerca_hidden")
				cercain="iduser"
				cerca=request.form("cerca_hidden")
			end select

		end if

		if where_anno<>"" and where<>""  then where_anno=" and "&where_anno
		where=where&where_anno
		if where<>"" then where=" where "&where
	 	strSql = sql &where&ordine
		PaginazioneString="&cercain="&cercain&"&cerca="&cerca&"&anno="&anno&"&cerca_hidden="&request("cerca_hidden")
		%>

 <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" onSubmit="<%=convalida%>">
          <table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1">
            <tr>
              <td colspan="7"> Cerca
                <input type="text" name="cerca" id="cerca" class="casella_ricerca" value="<%=cerca%>"><input type="hidden" id="cerca_hidden" name="cerca_hidden">
                in
                <select name="cercain" style="FONT: 12px;" onChange="Valida_ricerca()">
                  <option value="" <%if cercain="" then response.write "selected"%>>Visualizza tutti</option>
                  <option value="nddt" <%if cercain="nddt" then response.write selected%>>Numero D.d.T.</option>
                  <option value="idord" <%if cercain="idord" then response.write selected%>>N°ordine</option>
                  <option value="azienda" <%if cercain="azienda" then response.write selected%>>Azienda</option>
                  <option value="dafatt" <%if cercain="dafatt" then response.write selected%>>Da fatturare</option>
                  <option value="iduser" <%if cercain="iduser" then response.write selected%>>ID cliente</option>
                <option value="novendita" <%if cercain="novendita" then response.write selected%>>NO Vendita</option>
                </select>
                Anno:
                <input name="anno" type="text" style="FONT: 12px;" value="<%=anno%>" size="5" maxlength="5">
<input type="Submit" name="Submit" value="Cerca" style="FONT: 12px;" >

<span style="float: right;">
              <%if ha_il_permesso("B1") then %>
              <input type="button" name="ffm" value="Cerca per fatture fine mese" style="FONT: 12px;" onclick="location.href='pag_adm_fine_mese.asp'">
              <%end if %>
<input type="button" value="Nuovo" class="nuovodocumento" id="nuovo_ddt"></span>

</td>
            </tr>
            <tr>
              <td align="center" bgcolor="#E5E5E5"><p> Vedi</p></td>
              <td align="center" bgcolor="#E5E5E5">N&deg; D.d.T./Data</td>
              <td bgcolor="#E5E5E5">Causale/Incaricato</td>
              <td align="left" bgcolor="#E5E5E5">Azienda/Indirizzo/Citt&agrave;/Regione</td>
              <td bgcolor="#E5E5E5">Nominativo</td>
              <td bgcolor="#E5E5E5">Ordine</td>
              <td bgcolor="#E5E5E5">Fattura</td>
            </tr>
       <%
	Set ccausale_ddt = new cl_causale_ddt
call paginazione_start(strsql,ipagesize,"access")
Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
	'verde=ccausale_ddt.verde(objPagingRS("causale"))
	'sommadiimporto=objPagingRS("sommadiimporto")
	'if not isnull(sommadiimporto) then
	'	sommadiimporto=cdbl(sommadiimporto)
	'end if

	'totale=objPagingRS("totale")
	'if not isnull(totale) then totale=cdbl(totale)
	if objPagingRS("da_fatturare")=1 and objPagingRS("reso_tutto") then
		colore="bgcolor='#ffe680'"	'Arancione chiaro reso tutto
	elseif objPagingRS("da_fatturare")=0 then
		colore="bgcolor='#ffe680'"	'Arancione chiaro non fatturare
	elseif not isnull(objPagingRS("idfat")) then
		colore="bgcolor='#99FF99'" 'Verdino saldata
	'elseif totale=0 and objPagingRS("stato")>5 then
	'	colore="bgcolor='#99FF99'" 'Verdino saldata
	'elseif totale<0 then
	'	colore="bgcolor='#99FF99'"
	'elseif isnull(sommadiimporto) then
	'	colore=""
	'elseif totale=sommadiimporto then
	'	colore="bgcolor='#99FF99'" 'Verdino saldata
	'elseif totale>sommadiimporto then
	'	colore="bgcolor='#FFCCCC'" 'Rosino da saldare
	'	testo_differenza="("&formatcurrency(sommadiimporto-totale,2)&")"
	else
		colore=""
	end if
%>
            <tr class="coprobox" <%=colore%>>
              <td align="center" style="border-bottom:1px solid;"><input type="radio" name="view" value="<%=objPagingRS("idord")%>" onClick="this.form.submit()"></td>
              <td align="center" style="border-bottom:1px solid;"><%=pre_fattura(objPagingRS("data"))%><b><%=objPagingRS("nord")%></b><br>
                <%=formatDateTime(objPagingRS("data"), vbShortDate)%>

                </td>
              <td align="center" style="border-bottom:1px solid;"><b><%=ccausale_ddt.elemento(objPagingRS("causale"))%></b><br>
                <%if objPagingRS("incaricato")=3 then%>
                <%=incaricato_ddt(objPagingRS("vettore"))%>
                <%else%>
                <%=incaricato_ddt(objPagingRS("incaricato"))%>
                <%end if%>
                </td>
              <td style="border-bottom:1px solid;">
	              <b><%=denominazione(objPagingRS("nome"),objPagingRS("cognome"),objPagingRS("azienda"))%></b>

                </td>
              <td style="border-bottom:1px solid;"><b><%=objPagingRS("cognome")%>&nbsp;<%=objPagingRS("nome")%></b>
                </td>
              <td style="border-bottom:1px solid;"><%if objPagingRS("tipo_ddt")<>"no_ordine" then %> <a href="pag_adm_ordini.asp?idord=<%=objPagingRS("idordine")%>"><%=objPagingRS("nordine")%><%if objPagingRS("annoordine")<>year(date) then %>/<%=objPagingRS("annoordine")%><%end if%></a><br><%end if %><%=get_tipo_ddt( objPagingRS("tipo_ddt"))%></td>
              <td align="center" style="border-bottom:1px solid;"><a href="pag_adm_fatture.asp?idfat=<%=objPagingRS("idfat")%>"><%=objPagingRS("nfat")&get_fattura_pa(objPagingRS("pa"),objPagingRS("fattura_ce")) %></a></td>
            </tr>
            <%
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	call paginazione_end(7,true)
	set ccausale_ddt = Nothing
%>
          </table>
          </form>
          <script>
			var ddt_o_fattura="D";

	      </script>



          <%

	          else
	'***********************************************************************************************************************************************
	'****************************SCHEDA DDT
	'***********************************************************************************************************************************************

	sql1="select ddt.idddt,ddt.vettore, ordini.idord, ordini.nord, ordini.iduser, ordini.data, fatture.IDfat as fatture_idfat, fatture.Nfat as fatture_nfat, fatture.data as fatture_data FROM ((ddt LEFT JOIN ordini ON ddt.idord = ordini.idord) LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDfat  "
	  if idddt>0 then
		sql1=sql1&" WHERE ordini.idord= " & idddt &";"
	else
		sql1=sql1&" WHERE nddt= " & nddt &" and year(ddt.data)="&year(now())&";"
	end if
	set rs1=conn.execute(sql1)

idord=0
iduser=rs1("iduser")
nddt=rs1("nord")
nord=0
idfat=rs1("fatture_idfat")
nfat=rs1("fatture_nfat")
data_fattura=rs1("fatture_data")
data_ddt=rs1("data")
if oper="view" then %>

	<div id="multitabs" style="display:none;">
        <ul>
          <li><a href="#tabs-1" id="first_tab">DDT <%=nddt%></a></li>
		  <li><a href="pag_adm_user.asp?tab=1&iduser=<%=iduser%>">Dati cliente</a></li>
		  <%if idfat<>"" then %>
		  <li><a href="pag_adm_fatture.asp?tab=1&idfat=<%=idfat%>"><strong>Fattura <%=nfat&post_fattura%></strong> del <%=formatdatetime(data_fattura,2)%></a></li>
		  <%end if%>
		   <li id="tab_altro"><a href="#">Altro...</a></li>
        </ul>
        <div id="tabs-1" class="tab-panel">


          <%
	          end if
	        main_page=true
			call carica_tab(1,idddt)
			call connclose()
			if oper="view" then
%>
          </div>
	</div>

          <%
		  end if
end if
%>
<!-- Colonna CORPO FINE-->
<div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" --></div>
<%if oper="view" then%>
<script type="text/javascript" src="pag_adm_ordini.js?<%=ver_js_css %>"></script>

<script>
document.title = "DDT <%=pre_fattura(data_ddt)&nddt%>";
//var idord=<%=idord%>;
//var nord=<%=nord%>;
var tab_iduser=<%=iduser%>;
var tab_escludi="";
</script>
<%end if%>
 <%if oper="list" then%>
<script>
document.title = "DDT";
</script>
<%end if%>
</body>
</html>
<%
function zero (numero)
	if isnull(numero) then
		zero=0
	else
		zero=numero
	end if
end function
sub carica_tab(tab,idddt)
	select case tab
	case 1
		if oper="" then
			oper="view"
		end if
		prenotazione= prenota_record("ddt",idddt,"",false)

		set ordine= (new ClasseOrdine)(array("apri","ddt",idddt))
		ordine.calcolo_totale_merce_ddt(idddt)

			sql1="select o.*,d.*,d.parziale, u.tipologia, u.email, u.telefono, u.fax, u.cellulare, ordini_fatture.idfat, subo.nord as subnord, subo.anno as annoordine from (select ordini.* from ordini where tipo_documento='ddt') o  inner join ddt d on o.idord = d.idord inner join utenti u on o.iduser = u.iduser LEFT JOIN ordini_fatture ON o.idord = ordini_fatture.idord left join ordini subo on d.sub_idord = subo.idord"

		  if idddt>0 then
			sql1=sql1&" WHERE o.idord= " & idddt &";"
		else
			sql1=sql1&" WHERE nddt= " & nddt &" and year(ddt.data)="&year(now())&";"
		end if
		set rs1=conn.execute(sql1)
		nddt=RS1("nord")
		idord=RS1("idord")
		nord=rs1("nord")
		idddt=rs1("idord")
		iduser=rs1("iduser")
		azienda=rs1("azienda")
		trasporto=rs1("trasporto")



		if isnull(trasporto) then
			trasporto=0
		end if
		Set ccausale_ddt = new cl_causale_ddt
%>
          <script>
           function bollettino(){
			<%if cdbl(trasporto)=0 then%>
			var r=confirm('SEI SICURO CHE NON CI SONO SPESE TRASPORTO ?');
			if (r==false){
				return;
			  }
			<%end if%>
			  top.location.href="bollettino.asp?idord=<%=rs1("idord")%>";
		  }
          </script>
		<form id="form-ddt" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" onSubmit="<%=convalida%>">
          <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
            <tr>
              <td colspan="3" class="ui-widget-header">DDT
	              <span style="float:right;">
		<INPUT type="button" value="PDF" onClick='top.location.href="pdf_ddt.asp?idddt=<%=idddt%>"' class="ui-button ui-widget ui-state-default ui-corner-all puls_">


			<%if isnull(rs1("idfat")) and rs1("da_fatturare")=1 and ha_il_permesso("B1") then%>
			<input type="button"  id="pulsante_crea_ddt_fattura" value="Crea fattura" class="ui-button ui-widget ui-state-default ui-corner-all puls_">
			<%end if%>
			<%if rs1("porto")<>2 then%>
			<INPUT type="button" value="Bollettino" onClick='bollettino();' class="ui-button ui-widget ui-state-default ui-corner-all puls_">
			<%end if%>
			<INPUT type="button" value="Modifica"  data-id="<%=idord%>" data-documento="ddt" class="pulsante_modifica_intestazione ui-button ui-widget ui-state-default ui-corner-all puls_">
			  <button class="puls_giu" style="display:none;">Altro...</button>
              </span>
              <ul style="display:none;">
                <li ><a href="pag_adm_user.asp?iduser=<%=rs1("iduser")%>">Vedi cliente</a></li>


                <%
				sql="select Max(nord) FROM ordini where tipo_documento='ddt' and anno="&year(date())
				max_ddt=conn.execute(sql)(0)
				if max_ddt=nddt  and main_page then
				  %>

				<li><a href="#" id="cambia_intestazione">Cambia intestazione...</a></li>
				<li><a href="#" onClick="DialogYesNo ('Confermi l\'eliminazione del DDT?',f_elimina_ddt)">Elimina</a>
                <%end if%>
              </ul>
              </td>
            </tr>
            <tr>
              <td colspan="2" class="ui-widget-header">Dati cliente</td>
            </tr>
            <tr>
              <td align="center" bgcolor="#E5E5E5">N&deg; D.d.T./Causale/Data/Ordine</td>
              <td bgcolor="#E5E5E5">Nominativo/Email/Telefoni</td>
            </tr>
            <tr >
			<%
			if rs1("da_fatturare")=1 and rs1("reso_tutto") then
				colore="bgcolor='#ffe680'"	'Arancione chiaro reso tutto
			elseif rs1("da_fatturare")=0 then
				colore="bgcolor='#ffe680'"	'Arancione chiaro non fatturare
			elseif not isnull(rs1("idfat")) then
				colore="bgcolor='#99FF99'" 'Verdino saldata
			else
				colore=""
			end if



			%>



              <td width="" align="center" style="border-bottom:1px solid;" <%=colore%>">
  	            <%=pre_fattura(rs1("data"))%><b><%=rs1("nord")%></b><%if utente_andrea then response.write "/idord:"&rs1("idord")&"/ idddt:"&rs1("idddt") end if %> del <b><%=formatDateTime(RS1("data"), vbShortDate)%></b><br>
		        Causale<b> <%=ccausale_ddt.elemento(rs1("causale"))%></b>
                <%if rs1("sub_idord")>0 then %>
                <br>
                <a href="pag_adm_ordini.asp?idord=<%=RS1("sub_idord")%>">Ordine <%=RS1("subnord")%> <%=anno_se_diverso(rs1("annoordine")) %></a>
                <%end if %>


                </td>
            <td style="border-bottom:1px solid;"><b><%=ordine.campo("cognome")%>&nbsp;<%=ordine.campo("nome")%></b><%=vedi_tipologia(ordine.campo("tipologia"))%><br>
              <a href='mailto:<%=ordine.campo("email")%>'><%=ordine.campo("email")%></a><br>
              Telefono: <%=ordine.campo("telefono")%><br>
              Fax: <%=ordine.campo("fax")%><br>
              Cellulare: <%=ordine.campo("cellulare")%></td>
          </tr>
          </table>


        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
          <tr>
            <td class="ui-widget-header" style="border-right: none;">Dati cliente</td>
            <td align="right" class="ui-widget-header"  style="border-left: none;"></td>
          </tr>

		  <%
			call ordine.dati_fatturazione_consegna()
			call ordine.dati_mepa()
			  %>

        </table>
           <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr>
              <td colspan="2" valign="top" class="ui-widget-header">Dati D.d.T.
              </td>
            </tr>

            <%if utente_andrea then %>

            <tr>
              <td valign="top">Tipo</td>
              <td><%=rs1("tipo_ddt")%>: <%=get_tipo_ddt(rs1("tipo_ddt"))%>
               </td>
            </tr>
            <tr>
              <td valign="top">Parziale</td>
              <td><%=rs1("parziale")%>
               </td>
            </tr>
            <tr>
              <td valign="top">Sub_idord</td>
              <td><%=rs1("Sub_idord")%>
               </td>
            </tr>
            <%end if %>
            <tr>
              <td valign="top">Causale</td>
              <td>
                <%
                response.write ccausale_ddt.elemento(rs1("causale"))
	            set ccausale_ddt = Nothing
              %></td>
            </tr>
            <tr>
              <td valign="top">Da fatturare</td>
              <td>
	              <input type="checkbox" name="da_fatturare" id="da_fatturare" value="1" <%if rs1("da_fatturare")=1 then response.write " checked"%> disabled>

                </td>
            </tr>



            <tr>
              <td valign="top">Porto</td>
              <td>
                <%=porto_ddt(rs1("porto"))%>
              </td>
            </tr>
            <tr>
              <td valign="top">Imballo</td>
              <td>
                <%=imballo_ddt(rs1("imballo"))%>
              </td>
            </tr>
            <tr>
              <td valign="top">Numero colli</td>
              <td>
                <%=rs1("colli")%>
              </td>
            </tr>
            <tr>
              <td valign="top">Peso</td>
              <td>
				<%if oper<>"edit" then%>
                <%=rs1("peso")%>
                <%else%>
                <input name="peso" type="text" value="<%=rs1("peso")%>">
              <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Dimensione</td>
              <td>
                <%=rs1("dimensione")%>
              </td>
            </tr>
            <tr>
              <td valign="top">Annotazioni</td>
              <td>
                <%=rs1("annotazioni")%>
              </td>
            </tr>
            <tr>
              <td valign="top">Incaricato del trasporto</td>
              <td>
	              <%=incaricato_ddt(rs1("incaricato"))%>
				</td>
            </tr>
            <tr>
              <td valign="top">Vettore</td>
              <td>
                <%=rs1("vettore")%>
              </td>
            </tr>
            <tr>
              <td valign="top">Data inizio trasporto</td>
              <td><%=imballo_ddt(rs1("data_inizio"))%></td>
            </tr>
          </table>
          <%
	        tipo_ddt=rs1("tipo_ddt")
	        sub_idord=rs1("sub_idord")
	        'response.write "tipo:"&tipo_ddt &" sub_idord:"&rs1("sub_idord")
			set rs1=nothing
			'tipo=0 completo
			'tipo=1 parziale
			'tipo=2 noordine


			if tipo_ddt="completo" or tipo_ddt="no_ordine" then
				ordine.idord=sub_idord
			end if
			if tipo_ddt="parziale" then
				ordine.idord=idddt
				ordine.tipo_ddt=tipo_ddt
			end if
			call ordine.aprirecord("completo")
			if tipo_ddt="no_ordine" then
				ordine.modifica_articoli=true
			else
				ordine.modifica_articoli=false
			end if
			call ordine.elenco_articoli_head()
			ordine.mostra_giacenze=false
			ordine.mostra_spettanze=0
			ordine.sposta_articoli=false

			call ordine.elenco_articoli()

			call ordine.elenco_totali()
			if tipo_ddt="no_ordine" then
			call ordine.riga_pulsanti(true,true,true)
			end if



			%>
            <tr>
              <td colspan="6" align="right" >

				<input type="submit" name="aggiorna_reso" value="Aggiorna reso"  style="font-size:8pt" class="modifica-da-tab" data-id="<%=idddt%>" data-url="<%=questofile%>?tab=1" data-nomeid="idddt" data-oper="aggiorna-reso" data-form="form-ddt">
                <input type="hidden" id="elimina_ddt" name="elimina_ddt" value="">
                <input type="hidden" name="operazione" >
                <input type="hidden"  name="idddt" value="<%=idddt%>">
                <input type="hidden" name="idord" value="<%=idord%>"></td>
            </tr>
          </table>
                    </form>

<script type="text/javascript" src="Jquery/js/jquery.qrcode.min.js"></script>
<script type="text/javascript" src="Jquery/js/idle.js"></script>
<script type="text/javascript" src="Jquery/js/autogrow.min.js"></script>
                    <script>


			var documento="ddt";
			var idDocumento=<%=idddt%>;
			var nDocumento=<%=nddt%>;


var ordine=-1;

var inattivo=false;
var idadmin=<%=session("iduser")%>;
var iduser=<%=iduser%>;
var t_awayTimeout=<%=Application("rl_t_idle")*60%>000;
var rl_t_refresh=<%=Application("rl_t_refresh")%>000;
var prenotazione="<%=prenotazione%>";
var questofile="<%=questofile%>";
var tabella="ordini";
var cosa="ordine";
var aggiorna_ordine;
var iframe;
var tab_iduser=<%=iduser%>;
var tab_escludi="";
var tipo_documento="ddt";
//var tipo_allegato=<%=tipo_allegato%>;
var spettanze=false;
var ddt_o_fattura="F";
		$(function() {
			puls_giu();
			    $("#cambia_intestazione").on("click", function(e) {
                    e.preventDefault();

                    $("#dialog").html("<center>Attendi...</center>");
                    $.ajax({
                        url     : "pag_admin_cambia_intestazione.asp?id_documento=" + idDocumento+"&documento="+documento,
                        type    : "post",
                        cache: false,
                        //dataType: 'json',
                        //data	: dati,
                        success: function(data){
                            $("#dialog").html(data);
                        }
                        ,error: function(xhr, textStatus, error){
                            toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});

                            var txt="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
                            invia_errore("Errore in pulsante_modifica_intestazione()",stringa_dati+txt );
                        }
                    });
                    $("#dialog").dialog({
                        autoOpen: true,
                        modal: true,
                        resizable: true,
                        width: 600,
                        height: 200,
                        title: "Crea nuovo "+cosa
                    });
                });
			});


</script>
<%
	end select

end sub
sub ddt_update()
	add2log queryeform(),0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql= "select * from ddt where idddt=" & idddt
	rs.Open sql, conn, 3, 3
	if oper="update" then
		'memorizzo situazione iniziale
		n=rs.fields.count-1
		Set Flds = rs.Fields
		dim nome_campo(50)
		dim valore_campo(50)
		'memorizzo i nomi dei campi e i valori attuali
		for i=0 to n
			nome_campo(i)=RS.Fields(i).Name
			valore_campo(i)=rs(i)
		next
	end if
	rs("annotazioni")=trim(request.form("annotazioni_ddt"))
	rs("peso")=trim(request.form("peso"))
	val=trim(request.form("colli"))
	if val="" then rs("colli")=null else rs("colli")=val
	rs("dimensione")=trim(request.form("dimensione"))
	rs("causale")=trim(request.form("causale"))
	rs("imballo")=trim(request.form("imballo"))
	rs("porto")=trim(request.form("porto"))
	rs("vettore")=trim(request.form("vettore"))
	rs("incaricato")=request.form("incaricato")
	rs.update
	if oper="update" then
		for i=0 to n
			if valore_campo(i)<>rs(i) then
				modifiche_rs=modifiche_rs&"modificato campo "&nome_campo(i)&"<br>"
			end if
		next
		if modifiche_rs="" then modifiche_rs="Nessun campo modificato"
	end if
	Set ccausale_ddt = new cl_causale_ddt
	verde=ccausale_ddt.verde(rs("causale"))
	reso_tutto=rs("reso_tutto")
	if verde=false then
		verde=reso_tutto
	end if
	call imposta_verde(rs("idord"),verde)
	add2log "Modificati dati [ddt="&idddt&"]"&pre_fattura(rs("data"))&rs("nddt")&"/"&year(rs("data"))&"[/ddt] eseguita da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&modifiche_rs&vbcrlf&"Reso tutto:"&reso_tutto&vbcrlf&"Ordine verde:"&verde,2
	rs.close
	set rs=Nothing
	oper="view"
end sub
sub aggiorna_reso()


	idord=request("idord")
	'call movimento_magazzino(0,0,0,0,4,idord,0) 'idpro,quantita,idvara,idvarb,causale,idord,idmag
	stringa=""
	sql = "select ordini_dett.* FROM ordini_dett WHERE idord=" & idord
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	reso_tutto=true
	txt_log=""
	WHILE NOT RS.EOF
		true_false=lcase(trim(request.form("reso"&rs("iddett"))))
		txt_log=txt_log&"Analizzo iddett:"&rs("iddett")&" reso_form:"&true_false
		if true_false="si" then
			true_false=1
			txt_log=txt_log&", reso_db:"&rs("reso")
			if RS("reso")=0 then
				if stringa<>"" then stringa = stringa &", "
				stringa = stringa & rs("iddett")
				txt_log=txt_log&", aggiungo a stringa:"&stringa
			end if


		else
			reso_tutto=false
			true_false=0
		end if
		txt_log=txt_log&", imposto reso:"&true_false
		RS("reso") = true_false
		rs.update
		RS.MoveNext
		txt_log=txt_log&"<br>"
	WEND
	RS.Close

	set magazzino = new classeMagazzino
	txt_log=txt_log&"chiamo classemagazzino stringa:"&stringa&" idord:"&idord&"<br>"
	call magazzino.reso_da_ddt(stringa,idord)


	sql = "select * FROM ddt WHERE idddt=" & idddt
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	rs("reso_tutto")=converti_bool(reso_tutto)
	rs.update
	Set ccausale_ddt = new cl_causale_ddt
	verde=ccausale_ddt.verde(rs("causale"))
	set ccausale_ddt = Nothing
	if verde=false then
		verde=reso_tutto
	end if
	call imposta_verde(request("idord"),verde)



	add2log "Aggiornato reso [ddt="&idddt&"]"&pre_fattura(rs("data"))&rs("nddt")&"/"&year(rs("data"))&"[/ddt] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&"Reso tutto:"&reso_tutto&vbcrlf&"Ordine verde:"&verde&"<br>"&txt_log,2
	rs.close
	set rs=Nothing
	oper="view"
end sub
call CheckConnChiusa()
%>
