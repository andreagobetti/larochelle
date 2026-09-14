<%t_inizio=timer()%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/chk_piva_cf.asp" -->
<!--#include file="ClasseFattura.asp"-->

<!--#include virtual="/paginazione.asp" -->

<%
dim oper
dim main_page
main_page=false
if not utente_admin then call login()
idfat=request("idfat")
idord=request.form("idord")
nfat=request("nfat")
oper=lcase(request("oper"))
cercain=request("cercain")

call via_senza_permesso("B2")
if oper="ingranaggio" then
	call ingranaggio()
end if
if oper="setingranaggio" then
	elencopredefinito=request.form("elencopredefinito")
	Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
	Response.Cookies(questofile)("elencopredefinito")=elencopredefinito
	call add2log("Impostato elencopredefinito:"&elencopredefinito,0)
end if

if oper="update" then
	call fatt_update()
end if
Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class

if request("tab")<>"" then
	call carica_tab(request.querystring("tab"),idfat)
	call connclose()
	response.end
end if

if request.form("idscadenza")<>"" then
	tipo_pagamento=split(Application("metodi_incasso"),vbcrlf)
	Set rs_incassi = Server.CreateObject("ADODB.Recordset")
	sql="select scadenze.* FROM scadenze WHERE scadenze.idscadenza="&request.form("idscadenza")&";"
	rs_incassi.Open sql, conn, 3, 3
	causale=rs_incassi("causale")
	idord=rs_incassi("idord")
	if causale=0 then cosa="scadenza" else cosa="incasso"
	txt="["&cosa&"="&rs_incassi("idscadenza")&"]"&rs_incassi("data")& " "&formatcurrency(rs_incassi("importo"))&" "&causale_pagamento(rs_incassi("causale"))&" "&rs_incassi("tipo_pagamento")&"[/"&cosa&"]"
	rs_incassi("pagato")=1
	rs_incassi.update
	rs_incassi.close
	set rs_incassi=nothing
	add2log "Eliminato "&txt&" da [fattura="&idfat&"]"&nfat&"[/fattura] eseguito da [admin="&session("iduser")&"]"&session("nominativo")&"[/admin]" ,2
	idord=request.form("idord")
	nord=request.form("nord")
	oper="edit"
end if
stato_form=request.form("stato")
if stato_form<>"" then
	set rs=conn.execute("select fatture.* from fatture where idfat="&idfat)
	if stato_form>-1 and stato_form<>cint(rs("stato_fattura")) and stato_form<>"" then
		'ordine.campo("stato")=stato_form

		sql="update fatture set stato_fattura="&stato_form&" where idfat="&idfat
		conn.execute(sql)
		select case stato_form
		case 1

		case 6
			'n=metodipagamento.genera_scadenze(0,0,idord,rs("data").value,cdbl(rs("totale")),rs("tipopagamento").value)
		end select
	end if
	set rs = Nothing
end if


'response.write request.form
if request.querystring("elimina_fattura")<>"" then
	idfat=request.querystring("elimina_fattura")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "select * FROM fatture where idfat="& idfat , conn, 3,2
	pa=rs("pa")
	testo="Eliminato [fattura="&idfat&"]"&pre_fattura(rs("data"))&rs("nfat")
	if pa then testo=testo&"PA"
	testo=testo&"/"&year(rs("data"))&"[/fattura] "
	if pa then
		sql="DELETE from fatturepa WHERE idfatt=" & idfat&";"
		conn.execute sql,n
	end if
	sql="DELETE from ordini_fatture WHERE (((ordini_fatture.idfat)=" & idfat&"));"
	conn.execute sql,n
	testo=testo&" legata a "&n&" ordini,"
	sql="DELETE from incassi WHERE (incassi.causale=0 and ((incassi.idfat)=" & idfat&"));"
	conn.execute sql,n
	sql="DELETE from scadenze WHERE (((scadenze.idfat)=" & idfat&"));"
	conn.execute sql,n
	testo=testo&" legata a "&n&" scadenze incasso,  eseguito da [admin="&session("iduser")&"]"&session("nominativo")&"[/admin]"
	rs.delete
	rs.close
	set rs = nothing
	add2log testo,3
	call connclose()
	response.redirect questofile
end if
if idfat<>"" then oper="view"
select case oper
	case "annulla"
		oper="list"
	case "new"
		oper="new"
	case "modifica"
		oper="update"
	case "view"
		'Lascia così
	case else
		oper="list"
		if request.form("modifica")<>"" then
			oper="view"
			idfat=request.form("modifica")
		end if
		if request.form("edit")<>"" then oper="edit"
		if request.querystring("idfat")<>"" then
			oper="view"
			idfat=request("idfat")
		end if
		if isnumeric(request.form("cerca")) and request.form("cerca")<>"" and request.form("cercain")="" then
			oper="view"
			nfat=request.form("cerca")
			idfat=0
		end if
end select
if oper="edit" then
	call via_senza_permesso("B1")
end if



'response.write oper
%>
<!--#include virtual="/sub_head_adm.asp" -->
<!--#include virtual="/regioni_inc.asp" -->

<script src="Jquery/js/jquery.ui-contextmenu.min.js" type="text/javascript"></script>

<script type="text/javascript" language="javascript">
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
	if ((document.form1.cercain.value == 'idfat' )&&(document.form1.cerca.value == '')) {
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
function Valida_saldaordine()
{

	if (document.form1.tipo_pagamento.value == '' ) {
		alert("Selezionare il tipo pagamento");
		document.form1.tipo_pagamento.focus();
		return false;
	}
	document.form1.operazione.value='saldafattura';
	document.form1.submit();
}

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

<script src="jquery/ui/i18n/jquery.ui.datepicker-it.min.js"></script>
<style>
.ui-icon { display:inline-block; }​

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
		$( "#data_inizio" ).datepicker({ minDate: "<%=formatdatetime(date())%>" });
	});



</script>


</head>



<%


if oper="list" then
sql="select fatture.data, fatture.anno, fatture.idfat, fatture.fattura_sp, fatture.nfat, i.nome, i.cognome, i.azienda, i.indirizzo, i.citta, i.provincia,  cast(fatture.totale_fattura as decimal(10,2)) as totale_fattura ,fatture.pa,fatture.fattura_ce, fatture.invio_mail, fatture.download, fatture.pagamento,  fatture.idord, ordini.nord, sommadiincassi.SommaDiimporto, fatturepa.fatturapa FROM fatture left join utenti_intestazioni i on fatture.idintestazione = i.id LEFT JOIN (select incassi.idfat, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idfat ) AS sommadiincassi ON fatture.IDFat = sommadiincassi.idfat LEFT JOIN (select DISTINCT fatturepa.idfatt as fatturapa FROM fatturepa ) AS fatturepa ON fatture.IDFat = fatturepa.fatturapa left join ordini on fatture.idord = ordini.idord"

	elencopredefinito=Request.Cookies(questofile)("elencopredefinito")

	cerca=request("cerca")
		if cerca<>"" and isnumeric(cerca) and cercain="" then cercain="nfat"
		'response.write cerca<>""&isnumeric(cerca)&cercain=""
		'response.write "Prova"
		where_fisso = " "
		if request("anno")="" then
			anno=""
			where_anno=""
		else
			anno=request("anno")
			where_anno  = "  year(fatture.data)="&anno&" "
		end if
		iPageSize=Application("grec4page")
		ordine=" order by   anno desc, data desc, nfat desc"
		'-----------sezione ricerca inizio
		if request("cerca_hidden")<>"" then cercain="hidden"
		session("cercain")=cercain
		'Se non ho impostato niente carico filtro predefinito in preset
		if cercain="" then
			select case elencopredefinito
			case "ppa"
				cercain="pa"
			case "pnormali"
				cercain="nopa"
			case else
				cercain=elencopredefinito
			end select
		end if

		if cercain<>"" then
			select case cercain
				case "idord"
					where=" fatture.idord="&cerca
					where_anno=""
				case "iduser"
					where=" fatture.iduser="&cerca
					where_anno=""
				case "nfat"
					where=" nfat="&cerca
					where_anno=""
				case "idfat"
					where=" idfat="&cerca
					where_anno=""
				case "azienda"
					where=" InStr([ordini.azienda],'" & cerca& "')>0"
				case "hidden"
					where=" fatture.iduser="&request("cerca_hidden")
					where_anno=""
					cerca=""
				case "saldare"
					where=" totale_fattura>0 and (totale_fattura>sommadiimporto or sommadiimporto is null)"
					where_anno=""
					cerca=""
				case "saldateko"
					where=" totale_fattura<sommadiimporto"
					where_anno=""
					cerca=""
				case "pa"
					'where="  ((fatturepa.fatturapa) Is Not Null)"
					where="  pa=1"
					if elencopredefinito="nopa" or elencopredefinito="" then
						Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
						Response.Cookies(questofile)("elencopredefinito")=cercain
					end if
				case "nopa"
					'where="  ((fatturepa.fatturapa) Is Not Null)"
					where="  pa=0"
					cerca=""
					if elencopredefinito="pa" or elencopredefinito="" then
						Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
						Response.Cookies(questofile)("elencopredefinito")=cercain
					end if
									case "ce"
					'where="  ((fatturepa.fatturapa) Is Not Null)"
					where="  fattura_ce=1"
					if elencopredefinito="ce" or elencopredefinito="" then
						Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
						Response.Cookies(questofile)("elencopredefinito")=cercain
					end if


				case else

			end select
		end if
		if where_anno<>"" and where<>""  then where_anno=" and "&where_anno
		'where=where_fisso&" and "&where&where_anno
		where=where&where_anno
		if where<>"" then where=" where "&where
	 	strSql = sql &where&ordine
		PaginazioneString="&cercain="&cercain&"&cerca="&cerca&"&anno="&anno&"&cerca_hidden="&request("cerca_hidden")




end if


%>
<body>
<div id="wrap">
<div id="header">
<%=titolo_top%>
        <!-- Box CORPO INIZIO-->
		<%barra=6%>	<div id="barra_fissa">  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <%
	    if cercain="pa" then
			grigio="style='color:#999;'"
			grigioce="style='color:#999;'"
			document_title="Fatture PA"
	    elseif cercain="ce" then
			grigiopa="style='color:#999;'"
			grigio="style='color:#999;'"

			document_title="Fatture CE"

		else
			grigiopa="style='color:#999;'"
						grigioce="style='color:#999;'"

			document_title="Fatture"
		end if
		%>

        <div class="ui-widget-header ui-corner-top titolo_admin">
			<a href="<%=questofile%>?cercain=nopa" <%=grigio%>>Fatture</a> / <a href="<%=questofile%>?cercain=pa" <%=grigiopa%>>Fatture PA</a> / <a href="<%=questofile%>?cercain=ce" <%=grigioce%>>Fatture CE</a>
			<span style="float: right;"><a href="#" id="ingranaggio"><span class="ui-icon ui-icon-gear"></span></a></span>
	    </div></div>

          <%
if oper="list" then
		%>
          <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" onSubmit="<%=convalida%>">
          <table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1">
            <tr>
              <td colspan="6"> Cerca
                <input type="text" name="cerca" class="casella_ricerca" id="cerca" value="<%=cerca%>"><input type="hidden" id="cerca_hidden" name="cerca_hidden">
                in
                <select name="cercain" style="FONT: 12px;" onChange="Valida_ricerca()">
					<option value="" <%if cercain="" then response.write "selected"%>>Visualizza tutti</option>
					<option value="nfat" <%if cercain="nfat" then response.write selected%>>Numero fattura</option>
					<option value="idord" <%if cercain="idord" then response.write selected%>>N°ordine</option>
					<option value="azienda" <%if cercain="azienda" then response.write selected%>>Azienda</option>
					<option value="iduser" <%if cercain="iduser" then response.write selected%>>ID cliente</option>
	                <option value="saldare" <%if cercain="saldare" then response.write selected%>>Da saldare</option>
	                <option value="saldateko" <%if cercain="saldateko" then response.write selected%>>incassi incongruenti</option>
                </select>
                Anno:
                <input name="anno" type="text" style="FONT: 12px;" value="<%=anno%>" size="5" maxlength="5">
                <input type="submit" name="Submit" value="Cerca" style="FONT: 12px;" >
              <%if ha_il_permesso("B1") then %>
              <span class="pull-right">
              <input type="button" name="ffm" value="Cerca per fatture fine mese" style="FONT: 12px;" onclick="location.href='pag_adm_fine_mese.asp'">
              <input type="button" value="Nuova" class="nuovodocumento" id="nuovo_fattura">
              </span>


              </span>
              <%end if %>

                </td>
            </tr>
            <tr>
              <td align="center" bgcolor="#E5E5E5"><p> Vedi</p></td>
              <td align="center" bgcolor="#E5E5E5">N&deg; fattura/Data</td>
              <td bgcolor="#E5E5E5">Pagamento/Importo</td>
              <td align="left" bgcolor="#E5E5E5">Azienda/Indirizzo/Citt&agrave;/Regione</td>
              <td bgcolor="#E5E5E5">Nominativo</td>
              <td bgcolor="#E5E5E5">Ordine</td>
            </tr>
            <%
call paginazione_start(strsql,ipagesize,"access")

Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
testo_differenza=""
sommadiimporto=objPagingRS("sommadiimporto")
totale_fattura=cdbl(objPagingRS("totale_fattura"))
visualizza_scadenze=false
if not isnull(sommadiimporto) then sommadiimporto=cdbl(sommadiimporto)
if totale_fattura<0 then
	colore="bgcolor='#99FF99'"
elseif isnull(objPagingRS("sommadiimporto")) then
	colore=""
	visualizza_scadenze=true
elseif totale_fattura=sommadiimporto then
	colore="bgcolor='#99FF99'" 'Verdino saldata
elseif totale_fattura>sommadiimporto then
	colore="bgcolor='#FFCCCC'" 'Rosino da saldare
	testo_differenza="("&formatcurrency(sommadiimporto-totale_fattura,2)&")"
	visualizza_scadenze=true
else
	colore="bgcolor='#CC00CC'"
	testo_differenza="(+"&formatcurrency(sommadiimporto-totale_fattura,2)&")"
end if
%>            <tr class="coprobox" <%=colore%> style="border-top:1px solid;">
              <td align="center" ><input type="radio" name="modifica" value="<%=objPagingRS("idfat")%>" onClick="this.form.submit()"></td>
              <td align="center" >
	              <%=pre_fattura(objPagingRS("data"))%><b><%=objPagingRS("Nfat")&get_fattura_pa(objPagingRS("pa"),objPagingRS("fattura_ce"))&get_fattura_sp(objPagingRS("fattura_sp")) %></b><%if utente_andrea then response.write "/"&objPagingRS("idfat")%><br>
                <%=formatDateTime(objPagingRS("data"), vbShortDate)%><%if objPagingRS("invio_mail")<>"" then%><span class="ui-icon ui-icon-mail-closed"></span><%end if%><%if objPagingRS("download")<>"" then %><span class="ui-icon ui-icon-arrowthickstop-1-s"></span><% end if %><%if objPagingRS("fatturapa")<>"" then %><span class="ui-icon ui-icon-signal-diag"></span><% end if %>
                </td>
                            <td align="center" ><%=metodipagamento.descrizione(objPagingRS("pagamento"))%><br /><strong><%=formatcurrency(totale_fattura,2)%>
              </strong><%=testo_differenza%>
              </td>
              <td >
	              <b><%=denominazione(objPagingRS("nome"),objPagingRS("cognome"),objPagingRS("azienda"))%></b>
              </td>
              <td ><b><%=objPagingRS("cognome")%>&nbsp;<%=objPagingRS("nome")%></b></td>
              <td ><a href="pag_adm_ordini.asp?idord=<%=objPagingRS("idord")%>"><%=objPagingRS("nord")%></a></td>
            </tr>
            <%
	        if visualizza_scadenze then %>
	        <tr>
		        <td colspan="6" align="right">
			        <%
				        set rs_scadenze = conn.execute("Select * from scadenze where idfat="&objPagingRS("idfat"))
				        if rs_scadenze.EOF then
					      	response.write "Nessuna scadenz trovata"
					    else
					    txt_scadenze=""

				        do while not rs_scadenze.EOF
				        if rs_scadenze("data")<date() then
					        call concatena_stringa(txt_scadenze,", ","<span class=""rosso"">"& simbolo_valuta& rs_scadenze("importo")&" al "&rs_scadenze("data")&"</span>")
				        else
					        call concatena_stringa(txt_scadenze,", ", simbolo_valuta& rs_scadenze("importo")&" al "&rs_scadenze("data"))
				        end if


				        	rs_scadenze.MoveNext
				        Loop
				        Response.write "Scadenze: " &txt_scadenze
				        end if


				        %>

		        </td>
	        </tr>



	        <%end if





		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
    call paginazione_end(6,true)

%>
          </table>
          </form>
          <%else


          	sql="select fatture.iduser, fatture.nfat, fatture.idfat, fatture.data, fatture.pa,fatture.fattura_ce from fatture "
			if idfat>0 then
				sql=sql&" WHERE idfat= " & idfat &";"
			else
				sql=sql&" WHERE nfat= " & nfat &" and year(fatture.data)="&year(now())&";"
			end if
			'set rs1=conn.execute(sql1)
			set rs=conn.Execute(sql)
			idfat=rs("idfat")
			nfat=rs("Nfat")
			iduser=rs("iduser")
			data=rs("data")
			pa=rs("pa")
			ce=rs("fattura_ce")
			set rs = nothing
			totord=0
			if oper="view" then
	       %>
	       <div id="multitabs" style="display:none;">
        <ul>
          <li><a href="#tabs-1" id="first_tab">Fattura <%=nfat&get_fattura_pa(pa,ce)%></a></li>
		  <li><a href="pag_adm_user.asp?tab=1&iduser=<%=iduser%>">Dati cliente</a></li>
          <%
	          'Cerco ddt
	          sql="select ordini.idord, ordini.nord, ordini.data FROM ddt inner join ordini on ddt.idord=ordini.idord INNER JOIN ordini_fatture ON ddt.idord = ordini_fatture.idord where idfat="&idfat
	          set rs_ddt=conn.execute(sql)
	          do while not rs_ddt.eof

	          %>
		  <li><a href="pag_adm_ddt.asp?tab=1&idddt=<%=rs_ddt("idord")%>">DDT <%=rs_ddt("nord")%></a></li>
		  <%
			  rs_ddt.MoveNext
			  Loop
			  set rs_ddt = Nothing
			  %>
			<li id="tab_altro"><a href="#">Altro...</a></li>
        </ul>
        <div id="tabs-1" class="tab-panel">
	        <%
			end if
			main_page=true

			call   carica_tab(1,idfat)
			call connclose()

end if
%>
          <!-- Colonna CORPO FINE-->
</div>
<%
	txt_timer=txt_timer&"T fine:"&FormatNumber(timer() - t_inizio, 2)&" "
%>
<div id="footer"><%=txt_timer%></div>
<!--#include virtual="/pag_adm_footer_inc.asp" -->

</div>
  <%if oper="view" then%>
<script>
document.title = "Fattura <%=pre_fattura(data)&nfat%>";
var tab_iduser=<%=iduser%>;
var tab_escludi="";
var idfat=<%=idfat%>;
$(function() {



	var iframe = $('<iframe frameborder="0" marginwidth="0" marginheight="0" allowfullscreen></iframe>');
    var dialog = $("#dialog").append(iframe).dialog({
        autoOpen: false,
        modal: true,
        resizable: false,
        width: "auto",
        height: "auto",
        close: function () {
            iframe.attr("src", "");
        }
    });





	$("#invia_fattura").click(function(e) {
		toastr.options = {"positionClass": "toast-bottom-right"};
		toastr.info("Invio email in corso");

        $.ajaxSetup({ cache: false });

		$.get("invia_fattura.asp?idfat=<%=idfat%>", function(result){
				if (result.success==true)
				{
					//toastr.clear();
					toastr.options = {"positionClass": "toast-bottom-right"};
					toastr.success(result.Message);
				}
				else
				{
					toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
					toastr.error(result.Message);
				}
			});
	});
	$("#allega_files").on("click", function (e) {
		aggiorna_ordine=true;
        e.preventDefault();
        var src = "pag_upload_allegati_new.asp?tipo_allegato=6&id_tipo_allegato=<%=idfat%>";
        var title = "Aggiungi files a fattura <%=nfat%>";
        var width = 600;
        var height = 450;
        iframe.attr({
            width: +width,
            height: +height,
            src: src
        });
        dialog.dialog("option", "title", title).dialog("open");
    });



//	$(document).contextmenu({
//		delegate: ".left_menu_inc",
//		preventSelect: true,
//		taphold: true,
//		apertura: "click",
//
//		menu: [
//			{title: "Modifica incasso", cmd: "modifica", uiIcon: "ui-icon-pencil" },
//			{title: "----"},
//			{title: "Elimina incasso", cmd: "elimina", uiIcon: "ui-icon-trash" }
//			],
//		select: function(event, ui) {
//			var $target = ui.target;
//			var idincasso=$target.attr("id").replace("incasso_", "");
//			console.log("idincasso:"+idincasso)
//			switch(ui.cmd){
//			case "modifica":
//				var dialog=$("#dialog").dialog({
//					autoOpen: true,
//					modal: true,
//					resizable: false,
//					width: 600,
//					height: "auto",
//					title: "Aggiungi incasso "
//				});
//				dialog.html("Attendi...");
//
//				$.ajaxSetup({ cache: false });
//				$.ajax({
//					url     : "pag_adm_dialog_incasso.asp?ricarica=fattura&idincasso="+idincasso,
//					type    : "post",
//					//dataType: 'json',
//					//data	: dati,
//					success: function(data){
//						dialog.html(data);
//					}
//					,error: function(xhr, textStatus, error){
//						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
//						toastr.error('Errore nel caricamento della pagina');
//						var txt="";
//						txt+="Errore in "+location.href+"<br>";
//						txt+="<br>querystring: ?"+querystring[1];
//						txt+="<br>"+xhr.statusText;
//						txt+="<br>"+xhr.responseText;
//						txt+="<br>textStatus:"+textStatus;
//						txt+="<br>error:"+error;
//
//						$.ajax({
//							url     : "searcher.asp",
//							type    : "post",
//							data	: "txt_errore="+encodeURIComponent(txt)
//						});
//					}
//				});
//
//				break
//			case "elimina":
//				//location.href="pag_adm_ordini.asp?elimina_idincasso="+idincasso
//				$.ajax({
//					url     : "ajax_function.asp?oper=elimina_incasso&idincasso="+idincasso+"&idfat=<%=idfat%>",
//					type    : "post",
//					dataType: 'html',
//					cache: false,
//					//data	: dati,
//					success: function(data){
//
//						//$("#div_tabella_incassi").html(data);
//						 $(".div_tabella_incassi").each(function () {
//							var self = $(this);
//							self.html(data);
//						  });
//
//					}
//					,error: function(xhr, textStatus, error){
//						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
//						toastr.error('Errore nel caricamento della pagina');
//						console.log("xhr.statusText:"+xhr.statusText);
//						console.log("xhr.responseText:"+xhr.responseText);
//						console.log("textStatus:"+textStatus);
//						console.log("error:"+error);
//					}
//				});
//				break
//			}
//		},
//		// Implement the beforeOpen callback to dynamically change the entries
//		beforeOpen: function(event, ui) {
//			var $menu = ui.menu,
//				$target = ui.target;
//		//				.contextmenu("replaceMenu", [{title: "aaa"}, {title: "bbb"}])
//		//				.contextmenu("replaceMenu", "#options2")
//		//				.contextmenu("setEntry", "cut", {title: "Cuty", uiIcon: "ui-icon-heart", disabled: true})
//		// Optionally return false, to prevent opening the menu now
//		}
//	});





});

</script>
<%end if%>
 <%if oper="list" then%>
<script>
document.title = "<%=document_title%>";
</script>
<%end if%>
</body>
</html>
<%
rsClose
%>
<%
function zero (numero)
	if isnull(numero) then
		zero=0
	else
		zero=numero
	end if
end function
sub carica_tab(tab,idfat)

	select case tab
	case 1
		if oper="" then oper="view"
			'response.write idfat


			sql1="select fatture.*,  utenti.cellulare, utenti.fax, ordini.trasporto, ordini.tipo_documento, utenti_clienti.FatturaPACodiceDestinatario, utenti_clienti.stato_estero, ordini.idord, ordini.nord, ordini.anno as annoordine, utenti.email, utenti.telefono,  ordini.note, ordini.tipo_trasporto, ordini.noteacq FROM ((fatture left JOIN ordini ON fatture.idord = ordini.idord) INNER JOIN utenti ON fatture.iduser = utenti.iduser) LEFT JOIN utenti_clienti ON utenti.iduser = utenti_clienti.iduser"


			if idfat>0 then
				sql1=sql1&" WHERE idfat= " & idfat &";"
			else
				sql1=sql1&" WHERE nfat= " & nfat &" and year(fatture.data)="&year(now())&";"
			end if
			'set rs1=conn.execute(sql1)
			set rs1=conn.execute( sql1)
			idfat=rs1("idfat")
			nfat=rs1("Nfat")
			idord=rs1("idord")
			n=verifica_scadenze(idfat)
			if oper="view" then disabled=" disabled"

			'set fattura_pa=conn.Execute("select fatturepa.nomefile AS UltimoDinomefile, fatturepa.idfatt FROM fatturepa GROUP BY fatturepa.idfatt HAVING (((fatturepa.idfatt)="&idfat&")) order by id desc LIMIT 1;")
			set fattura_pa=conn.Execute("select fatturepa.nomefile AS UltimoDinomefile, fatturepa.idfatt FROM fatturepa where (((fatturepa.idfatt)= "&idfat&")) order by id desc LIMIT 1")

			file_fattura_pa=""
			if not fattura_pa.eof then
				file_fattura_pa=fattura_pa("UltimoDinomefile")
			end if

			if DateDiff("d","01/01/2015",rs1("data"))<0 then 'prima del 01/01/2015
				vecchio_conteggio=true
			else
				vecchio_conteggio=false
			end if
			'Conversione decimal in dbl
			totale_fattura=cdbl(rs1("totale_fattura"))
			spese_bancariev=cdbl(rs1("spese_bancarie"))
			if not isnull(idord) then
				trasporto=cdbl(rs1("trasporto"))
				sconto_ordine=rs1("sconto_ordine")
				if not isnull(sconto_ordine) then
					sconto_ordine=cdbl(sconto_ordine)
				else
					sconto_ordine=0
				end if
			else
				trasporto=0
				sconto_ordine=0
			end if



			imposta=cdbl(rs1("imposta"))
			totale_merce=cdbl(rs1("totale_merce"))
			FatturaPACodiceDestinatario=rs1("FatturaPACodiceDestinatario")
			'FatturaPACodiceDestinatario=rs1("CodiceDestinatario")

%>
        <script>
		$(function() {
			puls_giu();
			});

          function bollettino(){
				<%if trasporto=0 then%>
				var r=confirm('SEI SICURO CHE NON CI SONO SPESE TRASPORTO ?');
				if (r==false){
					return;
				  }
				<% end if %>

				  top.location.href="bollettino.asp?idord=<%=rs1("idord")%>";
		  }
		  function generafatturapa(){
				<%if file_fattura_pa<>"" then%>
				var r=confirm('UNA FATTURA ELETTRONICA ESISTE GIA\',\nVUOI GENERARE UNA NUOVA FATTURA CON NUOVO CODICE DI INVIO ?');
				if (r==false){
					return;
				  }
				<% end if %>
				top.location.href="fatturapa_v1.2.asp?idfat=<%=idfat%>"
		  }

	  		  function scaricafattura(){
		  		  console.log('scarica');
					top.location.href="fatturapa_scarica.asp?xml=<%=file_fattura_pa%>"
		  }



          </script>
        <form id="form-fatt" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" onSubmit="<%=convalida%>">
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
            <tr>
              <td colspan="2" class="ui-widget-header">Fattura <%if rs1("fine_mese")=0 then %>accompagnatoria<%end if %>

              <span style="float:right;">
         <%if rs1("fine_mese")=0 then %>
          <INPUT type="button" value="PDF fattura accompagnatoria" onClick='top.location.href="pdf_fattura.asp?idfat=<%=idfat%>"' class="ui-button ui-widget ui-state-default ui-corner-all">
          <%else %>
          <INPUT type="button" value="PDF fattura" onClick='top.location.href="pdf_fattura.asp?idfat=<%=idfat%>"' class="ui-button ui-widget ui-state-default ui-corner-all">
          <%end if%>

            <INPUT type="button" value="Modifica"  data-id="<%=idfat%>" data-documento="fattura" class="pulsante_modifica_intestazione ui-button ui-widget ui-state-default ui-corner-all puls_">

              <button class="puls_giu" style="display:none;">Altro...</button>
              </span>
              <ul style="display:none;">
                <li ><a href="pag_adm_user.asp?iduser=<%=rs1("iduser")%>">Vedi cliente</a></li>
                <li><a href="#" id="invia_fattura">Invia fattura via email</a></li>
                <%if rs1("invio_mail")<>"" then%><li> Ultimo invio <%=rs1("invio_mail")%></li><%end if%>
			  <%if FatturaPACodiceDestinatario<>"" and not isnull (FatturaPACodiceDestinatario) then%>
	          <li ><a href="#" onClick='generafatturapa();'>Genera fattura elettronica</a></li>
	          <%end if%>
			<%if file_fattura_pa<>"" then %>

				<li><a href="fatturapa_scarica.asp?xml=<%=file_fattura_pa%>">Scarica fattura elettronica</a></li>

				<li><a href="#" onClick='top.location.href="/public/fatture_pa/<%=file_fattura_pa%>"'>Vedi fattura elettronica</a></li>
			<%end if%>
			<li><a href="#" id="allega_files">Allega files</a></li>
          <%if rs1("porto")<>2 then%>
	          <li><a href="#" onClick='bollettino();'>Bollettino</a></li>
          <%end if%>
				<li><a href="#" id="cambia_intestazione">Cambia intestazione...</a></li>
				<li><a href="#" id="elimina">Elimina </a>

              </ul>
              </td>
            </tr>
            <tr>
              <td colspan="2" class="ui-widget-header">Dati cliente </td>
            </tr>
            <tr>
              <td width="50%" align="center" bgcolor="#E5E5E5">N&deg; fattura/Causale/Data/Ordine</td>
              <td width="50%" bgcolor="#E5E5E5">Nominativo/Email/Telefoni</td>
            </tr>
            <%
			Set ccausale_ddt = new cl_causale_ddt
			CAUSALE=ccausale_ddt.elemento(rs1("causale"))
			set ccausale_ddt = Nothing
			%>
            <tr >
              <td width="33%" align="center" style="border-bottom:1px solid;">
	            <%=pre_fattura(rs1("data"))%><b><%=rs1("Nfat")&get_fattura_pa(rs1("pa"),rs1("fattura_ce"))&get_fattura_sp(rs1("fattura_sp")) %></b> del <b><%=formatDateTime(RS1("data"), vbShortDate)%></b><br>
                Causale<b> <%=CAUSALE%></b><br>
                <a href="pag_adm_ordini.asp?idord=<%=RS1("idord")%>">Ordine <%=RS1("nord")%> <%=anno_se_diverso(rs1("annoordine")) %></a>
                <%if utente_andrea then%>IDFAT:<%=idfat%><%end if%>
                </td>

              <td width="33%" style="border-bottom:1px solid;">
                <a href='mailto:<%=rs1("email")%>'><%=rs1("email")%></a><br>
                Telefono: <%=rs1("telefono")%><br>
                Fax: <%=rs1("fax")%><br>
                Cellulare: <%=rs1("cellulare")%></td>
            </tr>
          </table>
          <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
	<%
	call dati_fatturazione_consegna(rs1)


	%>
		</table>
          <!-- ALLEGATI -->
<%tipo_allegato=6
id_tipo_allegato=idfat%>
<!--#include virtual="/sub_allegati.asp" -->
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr>
              <td colspan="2" valign="top" class="ui-widget-header">Dati fattura
              <%if oper="view" and ha_il_permesso("B1") then%><%if false then %><input name="edit" type="submit"  style="FONT: 12px;" value="Modifica" class="modifica-da-tab" data-id="<%=idfat%>" data-url="<%=questofile%>?tab=1" data-nomeid="idfat" data-oper="edit" data-form="form-fatt">
                <%end if
	                end if
                %>

              </td>
            </tr>
            <%if oper="view" and session("iduser")=1 then%>
            <tr>
              <td valign="top">IDfat</td>
              <td>
                <%=rs1("idfat")%>

              </td>
            </tr>
            <tr>
              <td valign="top">tipo_fattura</td>
              <td>
                <%=rs1("tipo_fattura")%>

              </td>
            </tr>

            <%end if%>
            <%if rs1("fine_mese")=0 then%>
            <tr>
              <td valign="top">Porto</td>
              <td><%if oper="view" then%>
                <%=porto_ddt(rs1("porto"))%>
                <%else%>
                <select name="porto">
                  <%
for n=1 to porto_ddt(0)
	response.write "<option value='"&n&"'"
	if rs1("porto")=n then response.write " selected style='color:red;'"
	response.write ">" & porto_ddt(n)&"</option>"
next
%>
                </select>
              <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Imballo</td>
              <td><%if oper="view" then%>
                <%=imballo_ddt(rs1("imballo"))%>
                <%else%>
                <select name="imballo">
                  <%
for n=1 to imballo_ddt(0)
	response.write "<option value='"&n&"'"
	if rs1("imballo")=n then response.write " selected style='color:red;'"
	response.write ">" & imballo_ddt(n)&"</option>"
next
%>
                </select>
              <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Numero colli</td>
              <td><%if oper="view" then%>
                <%=rs1("colli")%>
                <%else%>
                <input name="colli" type="text" value="<%=rs1("colli")%>">
              <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Peso</td>
              <td><%if oper="view" then%>
                <%=rs1("peso")%>
                <%else%>
                <input name="peso" type="text" value="<%=rs1("peso")%>">
              <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Dimensione</td>
              <td>
                <%if oper="view" then%>
                <%=rs1("dimensione")%>
                <%else%>
                <input name="dimensione" type="text" value="<%=rs1("dimensione")%>">
              <%end if%></td>
            </tr>

            <tr>
              <td valign="top">Incaricato del trasporto</td>
              <td>
                <%if oper="view" then%>


              <%=incaricato_ddt(rs1("incaricato"))%>
                <%if rs1("incaricato")=3 then%>
                <br>
                <%=incaricato_ddt(rs1("vettore"))%>
              <%end if

	              else
		              val=rs1("incaricato")
              %>
              <select name="incaricato" class="richiesto">
                  <%
					for n=1 to incaricato_ddt(0)
					%>
					<option value='<%=n%>' <%if n=val then%>selected<%end if%>> <%= incaricato_ddt(n)%></option>
					<%
					next
					%>
                </select>
                <%end if%>

              </td>
            </tr>
			<%
			val=""
			if rs1("data_inizio")<>"" then val=formatdatetime(rs1("data_inizio"),2)
			%>
            <tr>
              <td valign="top">Data inizio trasporto</td>
              <td>      <%if oper="view" then%>

                <%=val%>
                <%else%>
                <input name="data_inizio" type="text" id="data_inizio" value="<%=val%>">
              <%end if%></td>
            </tr>
            <%end if%>
             <tr>
              <td valign="top">Annotazioni</td>
              <td><%if oper="view" then%><%=rs1("annotazioni")%><%else%>
                <input name="annotazioni" type="text" value="<%=rs1("annotazioni")%>" style="width:90%;">
                <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Spese bancarie</td>
              <td><%
					if spese_bancariev>0 then
					  val=formatnumber(rs1("spese_bancarie"),2)
				  else
					  val=formatnumber(0,2)
				  end if



			  if oper="view" then
					  response.Write(val)
			    else


				%>
                <input name="spese_bancarie" type="text" value="<%=val%>">
                <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Trattamento IVA</td>
              <td><%




	              val=rs1("trattamento_iva")

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
            <%if rs1("pa") then %>
            <tr>
              <td valign="top">Codice univoco</td>
              <td>
	              <%if oper="view" then
		              if not isnull(rs1("CodiceDestinatario")) or rs1("CodiceDestinatario")<>"" then
			              codicedestinatario=rs1("CodiceDestinatario")
			          else
				          codicedestinatario="<span class=""rosso"">CODICE DESTINATARIO MANCANTE</span>"

		              end if %>


              <%=codicedestinatario%>

	         <%else
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
								if rs1("CodiceDestinatario")=codice then response.write " selected style='color:red;'"
								response.write ">" & riga&"</option>"
							next
						%>
		                </select>
		                <%
		            end if
	            end if%>


	            </td>
            </tr>

            <%end if%>
            <tr>
              <td valign="top">Banca appoggio cliente</td>
              <td><%if oper="view" then%><%=rs1("banca_appoggio")%><%else%>
                <input name="banca_appoggio" type="text" value="<%=rs1("banca_appoggio")%>" style="width:90%;">
                <%end if%></td>
            </tr>
            <tr>
              <td valign="top">IBAN cliente</td>
              <td>
              <%if oper="view" then %><%=rs1("iban") %><%else %>
                <input name="iban" type="text" value="<%=rs1("iban")%>" style="width:90%;">
                <%end if%></td>
            </tr>

            <%if oper="view" or utente_andrea then %>
            <tr>
              <td valign="top">fine mese</td>
              <td> <%val=""

	if rs1("fine_mese")=1 then val=" checked"
%>
          <input type="checkbox" name="fine_mese" value="si"  <%response.write val%> <%=disabled%>></td>
            </tr>
            <%end if %>

            <tr>
              <td valign="top">Nostra banca</td>
              <td><%val=rs1("idbanca")%><select name="idbanca"  <%=disabled%>>
              <option value="1" <%if val="" then response.write " selected"%>>Non impostata</option>

                  <%
set rsm=conn.execute("select * from banche order by nome_breve;")
do while not rsm.eof
%>
<option value="<%=rsm("idbanca")%>" <%if val=rsm("idbanca") then response.write " selected"%>><%=rsm("nome_breve")%></option>
<%
rsm.movenext
loop
rsm.close
set rsm=Nothing
%>
                </select></td>
            </tr>
            <tr>
              <td valign="top">Creata da</td>
              <td><%=(rs1("creata_da"))%></td>
            </tr>

			<%if oper="edit" then%>
            <tr>
              <td colspan="2" align="center">
	            <%
				if oper="view" then
					txt=puls_new
				elseif oper="new" or error<>"" then
					txt="Aggiungi"
				elseif oper="edit" then
					txt="Salva modifiche"
				end if
				%>
                <input type="submit" value="<%=txt%>" name="aggiorna" class="modifica-da-tab" data-id="<%=idfat%>" data-url="<%=questofile%>?tab=1" data-nomeid="idfat" data-oper="update" data-form="form-fatt">
                <input type="submit" name="oper" value="Annulla" onClick="annulla=true;">
              </td>
            </tr>
			<script>
			$(function() {
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
			});
			</script>

            <%end if 'oper="edit"%>
          </table>
          <%if utente_andrea then

			sql="select ordini_fatture.*, ordini.tipo_documento, ddt.tipo_ddt, ordini.trasporto FROM ordini_fatture inner join ordini on ordini_fatture.idord = ordini.idord left join ddt on ordini_fatture.idord = ddt.idord where idfat="&idfat
			set rs=conn.execute (sql)
			if not rs.eof then
				call query_to_table(rs)
			End If
          end if %>
          <%if rs1("tipo_fattura")="fattura" then
				set fattura= (new ClasseFattura)(idfat)
				call fattura.ControlloOperazioni()

				fattura.elenco_articoli_head()
				fattura.elenco_articoli()
				fattura.elenco_totali()
				call fattura.riga_pulsanti (false)
				call fattura.tabella_stato()
				%>


				<%


          elseif rs1("tipo_fattura")<>"fattura" then %>
          <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
            <tr>
              <td colspan="8" class="ui-widget-header">Articoli:</td>
            </tr>
            <tr>
              <td style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Codice</b></td>
              <td style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Articolo</b></td>
              <td align="center" class=testotabella style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" ><b>U.M.</b></td>
              <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Quantit&agrave;</b></td>
              <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Prezzo</b></td>
              <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella>Sconto</td>
              <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Totale</b></td>
            </tr>
            <%
			sconto_totale=0
			totord=0

			sql_ordine=get_sql_ordine_fattura(rs1("tipo_fattura"),idfat)

			if utente_andrea then
				response.write "<b>sql_ordine</b>"&sql_ordine
			end if
			on error resume next
			set rs_ordini=conn.execute(sql_ordine)
			if err.number<>0 then
				response.write sql_ordine
				response.end
			end if
			on error goto 0

			idord_tmp=0
			idtrasportoordine=0
			do while not rs_ordini.eof
			trasporto_ordine=0
			if idfat=718 or idfat=1003 then
			else

				if rs_ordini("tipo_ddt")="parziale" then
					'trasporto ddt parziale
					if idord_tmp<>rs_ordini("idord") then
						if idtrasportoordine<>rs_ordini("idord") then
						trasporto_ordine=cdbl(rs_ordini("trasporto"))
						idtrasportoordine=rs_ordini("idord")
						end if
					end if
					if cdbl(rs_ordini("trasporto_ddt"))>0 then
						trasporto_ordine=trasporto_ordine+cdbl(rs_ordini("trasporto_ddt"))
					end if
				else
					trasporto_ordine=cdbl(rs_ordini("trasporto"))
				end if

			end if
			if rs_ordini("tipo_ddt")<>"no_ordine" then%>
			%>
			<tr valign="middle">
				<td colspan="7" style="border-top: 1px solid black;"><%=qsql%><strong>Riferimento ordine: <a href="pag_adm_ordini.asp?idord=<%=rs_ordini("idord")%>"><%=rs_ordini("nord")%></a> del <%=formatdatetime(rs_ordini("dataordine"),2)%></strong></td>
            </tr>
            <%
	        end if
	        if rs_ordini("tipo_documento")="ddt" then
		        tipo_ddt=rs_ordini("tipo_ddt")


	        %>
            <tr valign="middle">
				<td colspan="7" style="border-bottom: 1px solid gray;"><strong>Riferimento ddt: <a href="pag_adm_ddt.asp?idddt=<%=rs_ordini("idddt")%>"><%=rs_ordini("nddt")%></a> del <%=formatdatetime(rs_ordini("dataddt"))%></strong></td>
            </tr>
            <%end if%>
			<%

			if rs_ordini("tipo_documento")="ddt" then
				if tipo_ddt="parziale" then
					sql_dettaglio="select ordini_dett.*,ddt_dett_ordini.quantita, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara inner join ddt_dett_ordini on ordini_dett.iddett = ddt_dett_ordini.iddett  where ddt_dett_ordini.idddt="&rs_ordini("idddt")&" order by ordine,iddett;"
				else
					sql_dettaglio="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara  where idord="&rs_ordini("sub_idord")&" order by ordine,iddett;"
				end if
			else
					sql_dettaglio="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara  where idord="&rs_ordini("idord")&" order by ordine,iddett;"

			end if
			if utente_andrea then
				response.write "<b>sql_dettaglio:</b>"&sql_dettaglio
			end if

			set rs=conn.execute(sql_dettaglio)
			do while not rs.eof
			%>

            <tr valign="middle">
				<td align="center" style="border-bottom: 1px solid gray;">
					<b><a href="product.asp?idpro=<%=RS("idpro")%>" title="Clicca per la scheda prodotto" ><%=codice_articolo_e_variante(  rs("codice_ordine"),rs("codicevara"))%></a></b>
				</td>
				<td style="border-bottom: 1px solid gray;"><%=rs("articolo_ordine")%>
				<%
				if rs("varianti_ordine")<>"" then response.write "<br>"&rs("varianti_ordine")
				if cint(rs("numeri_di_serie"))>0 then response.write "<br>"&elenca_seriali(rs("iddett"),0)
				%>
				</td>
				<td align="center" style="border-bottom: 1px solid gray;"><%=rs("um")%></td>
				<td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%=rs("quantita")%></td>
				<td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%=formatcurrency(rs("prezzo"),2)%></td>
				<td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%=FormatNumber(rs("sconto_prodotto"),2)%></td>
				<td align="right" valign="middle" style="border-bottom: 1px solid gray;"><%=formatcurrency(totale_riga(rs("prezzo"),rs("sconto_prodotto"),rs("quantita")),2)%></td>
            </tr>
            <%
	            if vecchio_conteggio then
		            prezzo=cdbl(rs("prezzo"))
					totord=totord+round((prezzo-(cdbl(rs("sconto_prodotto"))/100)*prezzo)*rs("quantita"),2)
		        else
					totord=totord+totale_riga(rs("prezzo"),rs("sconto_prodotto"),rs("quantita"))
				end if
	            rs.movenext
			loop
			set rs=Nothing
			if trasporto_ordine>0 then
			%>
            <tr valign="middle">
              <td align="right" style="border-bottom: 1px solid gray;">&nbsp;</td>
              <td colspan="5" style="border-bottom: 1px solid gray;">Trasporto</td>
              <td align="right" style="border-bottom: 1px solid gray;"><%=formatcurrency(trasporto_ordine,2)%></td>
            </tr>
            <%

	            totord=totord+trasporto_ordine
	            end if
	            %>
	            <%if sconto_ordine>0 then %>

            <tr valign="middle">
              <%
						  sconto_totale=sconto_totale+sconto_ordine
						%>
              <td colspan="6" style="border-bottom: 1px solid gray;">Sconto ordine</td>
              <td align="right" style="border-bottom: 1px solid gray;"><%=formatcurrency(sconto_ordine)%></td>
            </tr>
            <%end if %>


			<tr valign="middle">
              <%
			trasporto_ordine=0
			idord_tmp=rs_ordini("idord")

			rs_ordini.movenext
			loop
			set rs_ordini=Nothing
			if vecchio_conteggio then
				totord=round(totord,2)
			else
				totord=roundup(totord,2)
			end if

			%>
			<tr valign="middle">
				<td colspan="7" align="right" style="border-top: 1px solid black;">Totale IVA esclusa: <%=formatcurrency(totord,2)%></td>
			</tr>
            <%if spese_bancariev>0 then%>
            <tr valign="middle">
                            <td colspan="7" align="right" style="border-bottom: 1px solid gray;">Spese bancarie: <%=formatcurrency(rs1("spese_bancarie"))%></td>
            </tr>
            <%
			totord=totord+spese_bancariev
			end if

            testo_esenzione=" "&trattamento_iva(rs1("trattamento_iva"))
			if esenzione_iva(rs1("trattamento_iva")) then
				valore_iva=0
				if rs1("trattamento_iva")="10" then
					valore_iva=roundup(totord*rs1("aliquota_iva")/100,2)
					testo_esenzione=""
				end if
			else
				valore_iva=roundup(totord*rs1("aliquota_iva")/100,2)
				totord=totord+valore_iva
			end if
			%>
			<tr valign="middle" align="right" >
				<td colspan="7" style="border-bottom: 1px solid gray; ">IVA <%=rs1("aliquota_iva")%>%<%=testo_esenzione%>: &#8364; <span id="valore_iva"><%=FormatNumber(valore_iva)%></span></td>
			</tr>
			<%if rs1("trattamento_iva")="10" then%>
				<tr valign="middle" align="right" >
					<td colspan="7" style="border-bottom: 1px solid gray; "><%=trattamento_iva(rs1("trattamento_iva"))%>: &#8364; <span id="valore_iva">-<%=FormatNumber(valore_iva)%></span></td>
				</tr>
			<%end if%>
            <tr valign="middle">
              <td colspan="7" align="right" style="border-bottom: 1px solid gray;">Sconti ordini: <%=formatcurrency(sconto_totale)%></td>
            </tr>
            <%
			totord=totord-sconto_totale
			%>
            <tr valign="middle">
              <td colspan="7" align="right" style="border-bottom: 1px solid gray;"><b>Totale: <%=formatcurrency(totord,2)%></b></td>
            </tr>
			<%
			totord=roundup(totord,2)
			if totord<>totale_fattura then
				response.write "totale fattura differente"
				call add2log("Totale [fattura="&rs1("idfat")&"]"&rs1("nfat")&"[/fattura] differente, conteggio:"&totord&" db:"&rs1("totale_fattura"),3)

				n=calcola_totale_fattura(idfat)
				'response.redirect questofile&"?idfat="&idfat
			end if
			%>
            <tr>
              <td colspan="8" align="right" >
                <input type="hidden"  name="idfat" id="idfat" value="<%=idfat%>">
				<input type="hidden" name="idord" value="<%=rs1("idord")%>">
                <input type="hidden" name="nord" value="<%=rs1("nord")%>">
                <input type="hidden" name="nfat" id="nfat" value="<%=nfat%>">
			  </td>
            </tr>
          </table>


          <%

	          end if
	          set rs = Nothing
	          set rs1=Nothing
	           if oper ="view" then

				call tabella_scadenze(idfat,0)

				call tabella_incassi(0,idfat,totord,true,false)

		  %></div>
			</div>
			<script>
			var documento="fattura";
			var idDocumento=<%=idfat%>;
			var nDocumento=<%=nfat%>;
			//var inattivo=false;
			var iduser=<%=session("iduser")%>;
			var t_awayTimeout=<%=Application("rl_t_idle")*60%>000;
			var rl_t_refresh=<%=Application("rl_t_refresh")%>000;
			var prenotazione="<%=prenotazione%>";
			var questofile="<%=questofile%>";
			var aggiorna_ordine;
			var iframe;
			var tab_iduser=<%=iduser%>;
			var tab_escludi="";
			var tipo_allegato=<%=tipo_allegato%>;
			var spettanze=0;


			</script>


		<script type="text/javascript" src="Jquery/js/idle.js"></script>
		<script type="text/javascript" src="Jquery/js/autogrow.min.js"></script>

			<script type="text/javascript" src="pag_adm_ordini.js?<%=ver_js_css %>"></script>

			<script>
			$(function() {
				$( "#elimina" ).click(function(){
					var idfat=<%=idfat%>;
					$.ajax({
						url     : "ajax_function.asp?oper=verifica_elimina_fattura",
						type    : "post",
						//dataType: 'html',
						dataType: 'json',
						cache: false,
						data	: {idfat: idfat},
						success: function(data){

							$("#dialog").html(data.message);
							var dialog=$("#dialog").dialog({
								autoOpen: true,
								modal: true,
								resizable: false,
								width: "auto",
								height: "auto",
								title: "Elimina fattura",
								buttons: {

									"No": function () {
									$(this).dialog('close');
									//callback(false);
								},
									"Si": function () {
										$(this).dialog('close');
										f_elimina_fattura();
									}
								}
							});

								//alert(data.message);


						}
						,error: function(xhr, textStatus, error){
							toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
							toastr.error('Errore nel caricamento della pagina');
							console.log("xhr.statusText:"+xhr.statusText);
							console.log("xhr.responseText:"+xhr.responseText);
							console.log("textStatus:"+textStatus);
							console.log("error:"+error);
						}
					});	//$ajax

				});
			});
			</script>


			<%



			end if 'if oper="view"
			%>
			          </form>
<%

	          end Select
end sub
sub fatt_update()
	Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class


	Set rs = Server.CreateObject("ADODB.Recordset")
	sql= "select * from fatture where idfat=" & idfat
	rs.Open sql, conn, 1, 3
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
	'response.write "CAUSALE:"&trim(request.form("causale"))
	'if request.form("data")<>"" then

		'call add2log(,0)
	'	if cdate(request.form("data"))>=cdate(request.form("data_minima")) and cdate(request.form("data"))<=rs("data") then
			rs("data")=request.form("data")
	'	end if
	'end if
	if request.form("new_nfat")<>"" then
		rs("nfat")=request.form("new_nfat")

	end if
	rs("causale")=cint(trim(request.form("causale")))

	if rs("fine_mese")=false then
		rs("peso")=trim(request.form("peso"))
		rs("dimensione")=trim(request.form("dimensione"))
		rs("imballo")=trim(request.form("imballo"))
		rs("porto")=trim(request.form("porto"))
	end if
	rs("annotazioni")=trim(request.form("annotazioni"))
	rs("pagamento")=trim(request.form("pagamento"))
	rs("banca_appoggio")=trim(request.form("banca_appoggio"))
	rs("iban")=trim(request.form("iban"))
	rs("incaricato")=request.form("incaricato")
	rs("spese_bancarie")=aggiusta_decimale(trim(request.form("spese_bancarie")),"asp")
	if utente_andrea then
		if request.form("fine_mese")="si" then
			rs("fine_mese")=1
		else
			rs("fine_mese")=0
		end if
	end if
	if request.form("iva_0")="si" then
		rs("iva_0")=true
		else
		rs("iva_0")=false
	end if
	if request.form("trattamento_iva")<>"" then
		rs("trattamento_iva")=request.form("trattamento_iva")
	end if
	rs("idbanca")=trim(request.form("idbanca"))
	if request.form("CodiceDestinatario")="" then
		rs("CodiceDestinatario")=NULL
	else
		rs("CodiceDestinatario")=request.form("CodiceDestinatario")
	end if
	rs.update
	totale_fattura=calcola_totale_fattura(idfat)

	if oper="update" then
		on error resume next
		for i=0 to n
			if valore_campo(i)<>rs(i) then
				modifiche_rs=modifiche_rs&"modificato campo "&nome_campo(i)&"<br>"
			end if
		next
		on error goto 0
		if modifiche_rs="" then modifiche_rs="Nessun campo modificato"
	end if
	n_incassi=collega_incassi (idord,IDfat)

	set rs_incassi=conn.execute ("select Sum(incassi.importo) AS SommaDiimporto FROM incassi WHERE (((incassi.idfat)="&IDfat&"));")
	if not IsNull(rs_incassi("SommaDiimporto")) then
		SommaDiimporto=cdbl(rs_incassi("SommaDiimporto"))
	else
		SommaDiimporto=0
	end if

	add2log "Modificati dati [fattura="&idfat&"]"&rs("nfat")&"/"&year(rs("data"))&"[/fattura] eseguita da [admin="&session("iduser")&"]"&session("nominativo")&"[/admin],importo "&totale_fattura&", creato "&n&" scadenze riba, associato "&n_incassi&" incassi"&vbcrlf&modifiche_rs,2
	rs.close
	set rs = nothing
	oper="view"
end sub
sub ingranaggio()
elencopredefinito=Request.Cookies(questofile)("elencopredefinito")
 %>

<form action="<%=questofile%>" method="post">

<input type="hidden" name="oper" value="setingranaggio">
<p>Elenco fatture predefinito:
 <input type="radio" value="pnormali" name="elencopredefinito" <%if elencopredefinito="pnormali" then response.write "checked" end if%> >Fatture normali
 <input type="radio" value="ppa" name="elencopredefinito" <%if elencopredefinito="ppa" then response.write "checked" end if%>>Fatture PA
 <input type="radio" value="" name="elencopredefinito" <%if elencopredefinito="pa" or elencopredefinito="nopa" then response.write "checked" end if%>>Ricorda ultima selezione
 </p>

	<div class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix" >
		<div class="ui-dialog-buttonset">
			<button type="submit" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only" >
				<span class="ui-button-text">Salva</span>
			</button>
		</div>
	</div>
</form>


<%
call connclose()
response.end

end sub

call CheckConnChiusa()
%>
