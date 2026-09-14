<%
'Verifica chiusure 04_12_2015
%>
<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/chk_piva_cf.asp" -->
<!--#include virtual="/pag_adm_fatture_for_inc.asp" -->
<!--#include virtual="/paginazione.asp" -->
<%
if session("idadmin") = "" then call login()
idfat=request.form("idfat")
oper=lcase(request("oper"))
cosa=request("cosa")
'if request.querystring("cosa")<>"" then oper=request.querystring("cosa")
call via_senza_permesso("G1")
if request("anno")="" then
	anno=year(now())
else
	anno=request("anno")
end if
%>
<%
if request("elimina_allegato")<>"" then
	call elimina_allegato(request("elimina_allegato"))
end if
if request.form("elimina")<>"" then
	sql="DELETE from scadenze_for WHERE idfat=" & idfat&";"
	conn.execute sql,n
	testo=testo&" legata a "&n&" scadenze,"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "select * FROM files where cosa=1 and idcosa="& idfat , conn, 3,2
	testo=testo&" legata a "&rs.recordcount&" allegati,"
	do while not rs.eof
		call elimina_allegato (rs("idfiles"))
		rs.movenext
	loop
	rs.close
	testo=testo&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"
	rs.Open "select * FROM fatture_for where idfat="& idfat , conn, 3,2
	testo="Eliminato [fatturafornitore="&idfat&"]"&rs("nfat")&"/"&year(rs("data"))&"[/fatturafornitore] "&testo
	rs.delete
	rs.close
	set rs=Nothing
	add2log testo,3
	idfat=""
	Application("cache_cerca_senza_scadenze")=""
	call connclose()
	response.redirect questofile
end if
if request.form("elimina_scadenza")<>"" then
	sql="DELETE from scadenze_for WHERE id=" & request.form("elimina_scadenza")&";"
	conn.execute sql,n
	call tabella_scadenze_for()
	call connclose()
	response.end
	oper="view"
end if
if request.form("paga_scadenza")<>"" then
	sql="select * from scadenze_for   WHERE id=" & request.form("paga_scadenza")&";"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	idfat=rs("idfat")
	if rs("pagato")=0 then
		rs("pagato")=1
	else
		rs("pagato")=0
	end if
	rs.update
	rs.close
	set rs=nothing
	'call tabella_scadenze_for(idfat)
	call connclose()
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		js("success")=true
		js("message")="Scadenza pagata"
		js.Flush
		set js=Nothing
	response.end
end if
if request.form("paga_scadenza_stop")<>"" then
	sql="select * from scadenze_for   WHERE id=" & request.form("paga_scadenza_stop")&";"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	idfat=rs("idfat")
	if rs("pagato")=0 then
		rs("pagato")=1
	else
		rs("pagato")=0
	end if
	rs.update
	rs.close
	set rs=nothing
	call connclose()
	Response.ContentType = "application/json; charset=utf-8"
	Response.CodePage = 65001
	Set Js = jsObject()
	js("success")=true
	js("message")="Scadenza pagata"
	js.Flush
	set js=Nothing
	response.end
end if



if request.form("pagato")<>"" then
	sql="select * from scadenze_for   WHERE id=" & request.form("pagato")&";"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	if rs("pagato")=0 then
		rs("pagato")=1
	else
		rs("pagato")=0
	end if
	rs.update
	rs.close
	set rs=nothing
	call connclose()
	response.redirect questofile&"?cosa=scadenze"
	oper="list"
end if
if request.form("aggiungi_scadenze")<>"" then
'	totale_fattura=cdbl(conn.execute("select totale_fattura from fatture_for where idfat="&idfat)(0))
'	totale_fattura=request.form("totale_fattura")
	call aggiungi_scadenze(idfat)
	call tabella_scadenze_for(idfat)
	call connclose()
	response.end
	oper="view"
end if
select case oper
	case "annulla"
		oper="list"
	case "new"
		oper="new"
	case "aggiungi","aggiungi altre"
		oper="add"
	case "update"
		oper="update"
	case "view"
		oper="view"
	case "modifica"
		oper="update"
	case else
		oper="list"
		if cosa="" then cosa="fatture"
		if request.form("modifica")<>"" then
			oper="view"
			idfat=request.form("modifica")
		end if
		if request.form("aggiorna")<>"" then oper="update"
		if request.form("edit")<>"" then oper="edit"
		if request.querystring("idfat")<>"" then
			oper="view"
			idfat=request.querystring("idfat")
		end if

end select
'response.write request.form("edit")
if (oper="add" or oper="update") and error="" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql= "select * from fatture_for where idfat=" & idfat
	rs.Open sql, conn, 3, 3
	if oper="add" then
		rs("idfor")=request.form("fornitore_hidden")
	end if
	nfat=trim(request.form("nfat"))
	rs("nfat")=nfat
	rs("data")=trim(request.form("data"))
	rs("conto")=trim(request.form("conto"))
	if request.form("srl")="1" then
    		rs("srl")=1
    	else
    		rs("srl")=0
    	end if

	totale_fattura=aggiusta_decimale (request.form("totale_fattura"),"asp")

	'Non posso ricalcolare le scadenze perchè non ho la data inizio e il numero scadenze
	'if oper="update" then
	'	if cdbl(rs("totale_fattura"))<>cdbl(totale_fattura) then
	'		conn.execute("delete scadenze_for where ")
	'		call aggiungi_scadenze(idfat)
	'
	'
	'	end if
	'end if

	rs("totale_fattura")=totale_fattura
	rs.update

	if oper="add" then
		txt= "Aggiunto "

		'call aggiungi_scadenze(idfat)
	else
		txt= "Modificati dati "
	end if
	rs.close
	set rs=Nothing
	call add2log (txt&" [fatturafornitore="&idfat&"]"&nfat&"[/fatturafornitore] eseguita da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&queryeform(),2)
	call connclose()
	if request.form("oper")="Aggiungi" then
		Response.redirect questofile
	else
		Response.redirect questofile&"?oper=new"
	end if
end if

sub aggiungi_scadenze(idfat)
	if request.form("totale_fattura")<>"" then
		totale_fattura=aggiusta_decimale ( request.form("totale_fattura"),"asp")
	else
	totale_fattura=cdbl(conn.execute("select totale_fattura from fatture_for where idfat="&idfat)(0))
	end if
	conn.execute("delete from scadenze_for where idfat="&idfat)
	dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "select * from scadenze_for", conn, 3, 3
	n_scadenze=request.form("n_scadenze")
	if n_scadenze="altro" then
		n_scadenze=request.form("n_scadenze_altro")
	end if
	if request.form("acconto")<>"" then
		if request.form("data_acconto")="" then
			data_acconto=date()
		else
			data_acconto=request.form("data_acconto")
		end if
		acconto=aggiusta_decimale (request.form("acconto"),"asp")
		rs.addnew
		rs("idfat")=idfat
		rs("scadenza")=data_acconto
		rs("pagato")=1
		rs("importo")=acconto
		totale_fattura=totale_fattura-acconto
	end if

	n_scadenze=cint(n_scadenze)
	importo=roundup(totale_fattura/n_scadenze,2)
	scadenza=request.form("scadenza")
	if scadenza="altro" then
		scadenza=request.form("scadenza_altro")
	end if
	importo_progressivo=0
	for i= 1 to n_scadenze
		rs.addnew
		rs("idfat")=idfat
		if n_scadenze=1 then
			rs("scadenza")=scadenza
		else
			rs("scadenza")=finemese(scadenza,i)
		end if
		if i=n_scadenze then
			importo=totale_fattura-importo_progressivo
			importo=roundup(importo,2)
			rs("importo")=importo
		else
			rs("importo")=importo
		end if
		importo_progressivo=importo_progressivo+importo
		call add2log("caricato scadenza "&i&":"&importo&" importo_progressivo:"&importo_progressivo,0)

		rs.update
	next
	rs.close
	set rs=Nothing



end sub

'response.write oper
%>
<!--#include virtual="/regioni_inc.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
	<title><%=application("brwstitle")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
	<!--#include virtual="/sub_head.asp" -->
	<script type="text/javascript" language="javascript">


function Valida_ricerca()
{
	if ((document.form1.cercain.value == 'idfat' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il numero fattura nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	if ((document.form1.cercain.value == 'azienda' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il nome fornitore nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
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
.ui-autocomplete-loading {
	background: white url('images/ui-anim_basic_16x16.gif') right center no-repeat;
}
label {
	display: inline-block;
	width: 5em;
}
.box_errore{
	position: relative; left: 10px; top:-10; width:300px; z-index:100;
	display: none;
	}
</style>
	<script>
	$(function() {
		$.ajaxSetup({ cache: false });
		//$.datepicker.setDefaults($.datepicker.regional['it']);
		$( "#data2" ).datepicker({dateFormat: 'dd/mm/yy'});
		$( "#data1" ).datepicker();
		$( "#data3" ).datepicker();
		$( document ).tooltip({position: { my: 'left center',   at: 'right center'  }  });
		//$( document ).tooltip();
		$( "#fornitore" ).autocomplete({
			source: "ajax_function.asp?autocomplete=fornitori",
			minLength: 2,
			cache: false ,
			select: function( event, ui ) {
				$('#fornitore_hidden').val(ui.item.id);
			}
		});
		$( "#cerca" ).autocomplete({
			source: "ajax_function.asp?autocomplete=fornitori",
			minLength: 2,
			cache: false ,
			select: function( event, ui ) {
				$('#cerca_hidden').val(ui.item.id);
			}
		});
		$("#nfat").blur(function(){
			$.ajax({
			 url: "ajax_test.asp?test=1&nfat="+$("#nfat").val() +"&idfor="+$("#fornitore_hidden").val() ,
			 type: "post",
			 success: function(data) {
				var result = eval(data); //result sarà true se data == "true", false se data =="false"
						if(result){
						   //Controllo passato
						   $("#nfat_error").hide();
						}
						else{
						   //Controllo non passato
						   $('#nfat_error').addClass('ui-state-error ui-corner-all');
						   $("#nfat_error").text("Fattura già caricata");
						   //$( "#nfat_error" ).dialog();
						  $("#nfat_error").show();
						   			//alert("Fattura già caricata");
						}
			 }
			});
		});
	});


	</script>
	</head>
	<body>
    <div id="wrap">
      <div id="header"> <%=titolo_top%>
        <%
        barra=0
        %><div id="barra_fissa">
        <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <%
	if cosa="scadenze" then
		grigiocl="style='color:#999;'"
		document_title="Fatture fornitori"
	else
		grigiofr="style='color:#999;'"
		document_title="Scadenze fatture fornitori"
	end if
	%>
        <div class="ui-widget-header ui-corner-all titolo_admin">
      <a href="<%=questofile%>?cosa=fatture" <%=grigiocl%>>Fatture fornitori</a> / <a href="<%=questofile%>?cosa=scadenze" <%=grigiofr%>>Scadenze fatture fornitori</a>
        </div></div>
        <!-- Box CORPO INIZIO-->
          <%

		if oper="list" then


		if cosa="scadenze" then
			sql="select fatture_for.*, scadenze_for.*,utenti_intestazioni.azienda FROM (scadenze_for INNER JOIN fatture_for ON scadenze_for.idfat = fatture_for.idfat) INNER JOIN utenti ON fatture_for.idfor = utenti.iduser inner join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id "
			clausola_where=" WHERE "
			'where_anno  = " year(fatture_for.data)="&anno&" "
			ordine=" order by scadenze_for.scadenza desc"

		elseif cosa="fatture" then

			sqla="SELECT distinct idfat FROM `scadenze_for` idfat where pagato=0"
			sqlb="SELECT sum(case when pagato=0 then 1 else 0 end) as dasaldare, count(*) as conteggio  FROM `scadenze_for` group by idfat"

			sql="select fatture_for.idfat,fatture_for.idfor, fatture_for.nfat, fatture_for.data, utenti.Azienda, Count(scadenze_for.ID) AS ConteggioDiID, fatture_for.conto, fatture_for.totale_fattura, utenti.iduser FROM (fatture_for INNER JOIN utenti ON fatture_for.idfor = utenti.iduser) LEFT JOIN scadenze_for ON fatture_for.idfat = scadenze_for.idfat GROUP BY fatture_for.idfat, fatture_for.idfor,fatture_for.nfat, fatture_for.data, utenti.Azienda, fatture_for.conto, fatture_for.totale_fattura, utenti.iduser   "

			sql="select fatture_for.idfat,fatture_for.idfor, fatture_for.nfat, fatture_for.data, utenti.Azienda, fatture_for.conto, fatture_for.totale_fattura, utenti.iduser,s.dasaldare, s.conteggio FROM (fatture_for INNER JOIN utenti ON fatture_for.idfor = utenti.iduser) LEFT JOIN (SELECT idfat, sum(case when pagato=0 then 1 else 0 end) as dasaldare, count(*) as conteggio  FROM `scadenze_for` group by idfat) as s ON fatture_for.idfat = s.idfat GROUP BY fatture_for.idfat, fatture_for.idfor,fatture_for.nfat, fatture_for.data, utenti.Azienda, fatture_for.conto, fatture_for.totale_fattura, utenti.iduser "
			sql="select fatture_for.idfat,fatture_for.srl,fatture_for.idfor, fatture_for.nfat, fatture_for.data, utenti_intestazioni.Azienda, fatture_for.conto, fatture_for.totale_fattura, utenti.iduser,s.dasaldare, s.conteggio FROM (fatture_for INNER JOIN utenti ON fatture_for.idfor = utenti.iduser) LEFT JOIN (SELECT idfat, sum(case when pagato=0 then 1 else 0 end) as dasaldare, count(*) as conteggio  FROM `scadenze_for` group by idfat) as s ON fatture_for.idfat = s.idfat inner join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id  "



			clausola_where=" HAVING "
			'where_anno="(Year(data)="&anno&")"
			ordine=" order by fatture_for.data desc"
		end if
		if anno="Tutti" then where_anno=""
		cercain=request("cercain")
		cerca=request("cerca")
		if cerca<>"" and isnumeric(cerca) and cercain="" then cercain="iduser"
		iPageSize=Application("grec4page")

		if request("cerca_iduser")<>""  then
			cercain="iduser"
			cerca=request("cerca_iduser")
		end if
		if cercain<>"" then
			select case cercain
			case "iduser"
				where=" fatture_for.idfor="&cerca
				where_anno  = ""
				response.write "qui"
			case "idfat"
				where=" nfat='"&cerca&"' "


			case "saldate"
			if cosa="scadenze" then
				where=" pagato=1 "
				else
				where=" dasaldare=0 "
				end if


			case "dasaldare"


			if cosa="scadenze" then
				where=" pagato=0 "
				else
				where=" dasaldare=1 "
				end if


			case "noscadenze"
				where=" conteggio is null and totale_fattura>0"
				where_anno  = ""
			end select
		end if

		if where_anno<>"" then
			'response.write "<b>WHERE:"&len(where)&"</b>"'
			if where<>"" then str_and=" AND "

			where=where&str_and&where_anno
		end if
		if where<>"" then where=clausola_where &where
		strSql = sql &where&ordine
		paginazionestring="&cercain="&cercain&"&cerca="&cerca&"&cosa="&cosa
		if where_anno="" then anno="Tutti"
		'-----------sezione ricerca fine
		'response.write strSql
		%>
        <form id="form1"  name="form1" method="post" action="<%=questofile%>" style="margin-top:0px;">
          <table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1">
            <tr>
              <td colspan="7"><input type="hidden" name="cosa" value="<%=cosa%>" /> Cerca
                <input type="text" name="cerca" id="cerca" style="FONT: 12px;" value="<%'=cerca%>"><input type="hidden" id="cerca_hidden" name="cerca_iduser">
                in
                <select name="cercain" style="FONT: 12px;" onChange="Valida_ricerca()">
                  <option value="" <%if cercain="" then response.write "selected"%>>Visualizza tutti</option>
                  <option value="idfat" <%if cercain="idfat" then response.write selected%>>Numero fattura</option>
                  <option value="azienda" <%if cercain="iduser" then response.write selected%>>Azienda</option>
                  <option value="saldate" <%if cercain="saldate" then response.write selected%>>Saldate</option>
                  <option value="dasaldare" <%if cercain="dasaldare" then response.write selected%>>Da saldare</option>
                </select>
                Anno:
                <input name="anno" type="text" style="FONT: 12px;" value="<%=anno%>" size="5" maxlength="4">
                <input type="button" value="Cerca" style="FONT: 12px;" onClick="Valida_ricerca()"><span style="float:right;"><input type="button"  value="Nuovo" style="FONT: 12px;" onclick="window.location.href='<%=questofile%>?oper=new'"></span></td>
            </tr>
            <%
iRecordsShown = 0
if cosa="fatture" then

			response.write cerca_senza_scadenze()
	%>
            <tr>
              <td colspan="2" align="center" bgcolor="#E5E5E5"><p> Vedi</p></td>
              <td align="center" bgcolor="#E5E5E5">Data</td>
              <td bgcolor="#E5E5E5">N&deg; fattura</td>
              <td align="left" bgcolor="#E5E5E5">Fornitore</td>
              <td bgcolor="#E5E5E5">Conto</td>
              <td align="right" bgcolor="#E5E5E5">Totale fattura</td>
              <td align="right" bgcolor="#E5E5E5">Srl</td>
            </tr>
            <%
      call paginazione_start(strsql,iPageSize,"access")
		Do While iRecordsShown < iPageSize And Not objPagingRS.EOF

			if isnull(objPagingRS("conteggio")) then
			colore="bgcolor='red'"
		elseif cint(objPagingRS("dasaldare"))=0 then
			colore="bgcolor='#99FF99'"

		else
			colore=""
		end if





%>
            <tr class="coprobox" style="border-bottom:1px solid;" <%=colore%>>
              <td colspan="2" align="center" ><input type="radio" name="modifica" value="<%=objPagingRS("idfat")%>" onClick="this.form.submit()"></td>
              <td align="center" >
                <%=formatDateTime(objPagingRS("data"), vbShortDate)%></td>
              <td align="center" ><b><%=objPagingRS("Nfat")%></b></td>
              <td ><b><%=objPagingRS("azienda")%></b></td>
              <td ><%=conto(objPagingRS("conto"))%></td>
              <td align="right" ><%=formatcurrency(objPagingRS("totale_fattura"),2)%></td>
              <td align="right" ><%
              if objPagingRS("srl") then
              %>
              si
              <%
              end if
              %></td>
            </tr>
            <%
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
		Loop
	elseif cosa="scadenze" then
		response.write cerca_senza_scadenze()
			%>
            <tr>
              <td align="center" bgcolor="#E5E5E5"><p> Vedi</p></td>
              <td align="center" bgcolor="#E5E5E5">Scadenza</td>
              <td bgcolor="#E5E5E5">N&deg; fattura</td>
              <td align="left" bgcolor="#E5E5E5">Fornitore</td>
              <td bgcolor="#E5E5E5">Conto</td>
              <td align="right" bgcolor="#E5E5E5">Importo</td>
              <td align="right" bgcolor="#E5E5E5">Srl</td>
              <td align="center" bgcolor="#E5E5E5">Pagato</td>
            </tr>
            <%

      call paginazione_start(strsql,iPageSize,"access")
		Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
		if objPagingRS("pagato") then
			colore="bgcolor='#99FF99'"
		else
			colore=""
		end if
%>
            <tr class="coprobox" <%=colore%>>
              <td align="center" style="border-bottom:1px solid;"><input type="radio" name="modifica" value="<%=objPagingRS("idfat")%>" onClick="this.form.submit()"></td>
              <td align="center" style="border-bottom:1px solid;"><%=formatDateTime(objPagingRS("scadenza"), vbShortDate)%></td>
              <td align="center" style="border-bottom:1px solid;"><b><%=objPagingRS("Nfat")%></b></td>
              <td style="border-bottom:1px solid;"><b><%=objPagingRS("azienda")%></b></td>
              <td style="border-bottom:1px solid;"><%=conto(objPagingRS("conto"))%></td>
              <td align="right" style="border-bottom:1px solid;"><%=formatcurrency(objPagingRS("importo"),2)%></td>
                            <td align="right" style="border-bottom:1px solid;"><%
                            if objPagingRS("srl") then
                            %>
                            si
                            <%
                            end if
                            %></td>
              <td align="center" style="border-bottom:1px solid;"><%if colore="" then %><input type="radio" id="<%=objPagingRS("id")%>" class="paga" /><%end if %></td>
            </tr>
            <%
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
end if
	call paginazione_end(7,true)
%>
                </table>
              </td>
            </tr>
          </table>
          </form>


          <script>
	          $(function() {
	          $(".paga").click(function(e){
		          e.preventDefault();
		          var obj=$(this)
		          var id=obj.attr("id");
		          console.log("paga "+id);

 					$.ajax({
			            url: "pag_adm_fatture_for.asp",
			            method: "post",
			            cache: false,
			            dataType: 'json',
			            data	: {
				            'paga_scadenza_stop':id
			            } ,
			            success: function(data) {

				          obj.closest('tr').css("background",'#99FF99');
				          obj.remove();

			            },
			            error: function(xhr, textStatus, error) {
			                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
			                console.log("xhr.statusText:" + xhr.statusText);
			                console.log("xhr.responseText:" + xhr.responseText);
			                console.log("textStatus:" + textStatus);
			                console.log("error:" + error);
			            }
			        });








	          })
	          });


          </script>

          <%
	           else
		           %>
        <form id="form1"  name="form1" method="post" action="<%=questofile%>" style="margin-top:0px;">
		           <%
	          'oper="list"
				if oper="view" then disabled=" disabled"
				if oper="view" or oper="edit" then
					Set rs = Server.CreateObject("ADODB.Recordset")
					'response.write idfat
					sql1="select  utenti.*, fatture_for.* FROM fatture_for INNER JOIN utenti ON fatture_for.idfor = utenti.iduser  WHERE idfat= " & idfat &";"
					'set rs=conn.execute(sql1)
					rs.open sql1,conn,3,3
				end if
			if oper<>"new" then

			%>
          <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
            <tr>
              <td colspan="3" class="ui-widget-header">Dati Fornitore <span style="float:right;"><button type="button" onclick="window.location.href='pag_adm_user.asp?iduser=<%=rs("iduser")%>'" style="font-size:8pt">Vedi fornitore</button></span></td>
            </tr>
            <tr>
              <td align="center" bgcolor="#E5E5E5">N&deg; fattura/Causale/Data/Ordine</td>
              <td align="left" bgcolor="#E5E5E5">Azienda/Indirizzo/Cap Citt&agrave;/Regione</td>
              <td bgcolor="#E5E5E5">Nominativo/Email/Telefoni</td>
            </tr>
            <tr >
              <td width="33%" align="center" style="border-bottom:1px solid;"><b><%=rs("Nfat")%></b><%if utente_andrea then response.write "/"&rs("idfat")%><br>
                <%=formatDateTime(rs("data"), vbShortDate)%></td>
              <td width="33%" style="border-bottom:1px solid;">
			  <%=dati_fatturazione(rs("idintestazione"))%>

			  </td>
              <td width="33%" style="border-bottom:1px solid;">
                <a href='mailto:<%=rs("email")%>'><%=rs("email")%></a><br>
                Telefono: <%=rs("telefono")%><br>
                Fax: <%=rs("fax")%><br>
                Cellulare: <%=rs("cellulare")%></td>
            </tr>
          </table>

		<%end if
		if oper="new" then
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open "fatture_for",conn,3,3
			rs.addnew
			rs("idfor")=0
			rs.update
			idfat=Get_last_id("fatture_for")
			rs.close
			set rs = Nothing
          end if %>


          <input type="hidden" name="idfat" value="<%=idfat%>" /><%if utente_andrea then response.write "idfat:"&idfat%>

          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr>
              <td colspan="2" valign="top" class="ui-widget-header">Dati fattura
                <%if oper="view"  then%>
                <input name="edit" type="submit"  style="FONT: 12px;" value="Modifica">
                <%end if%></td>
            </tr>
            <%if oper="new" then%>
            <tr>
              <td width="20%" valign="top">Fornitore</td>
              <td width="80%"><input name="fornitore" type="text" id="fornitore" style="FONT: 12px;" class="richiesto" value="<%'=cerca%>" size="60" title="Selezionare il fornitore, campo obbligatorio.">
              <input type="hidden" id="fornitore_hidden" name="fornitore_hidden"></td>
            </tr>
            <%end if%>
            <tr>
              <td valign="top">Fattura n°</td>
              <td><%if oper="view" then%>
                <%=rs("nfat")%>
                <%else %>
                <%if oper="new" then
                            val=request.form("nfat")
                        else
                            val=rs("nfat")
                        end if%>
                <input name="nfat" type="text" class="richiesto" id="nfat" value="<%=val%>" size="40" title="Inserisci il numero della fattura, campo obbligatorio."><span style="position:absolute;"><div id="nfat_error" class="box_errore"></div></span>
              <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Data fattura</td>
              <td><%if oper="view" then%>
                <%=formatdatetime(rs("data"),2)%>
                <%else%>
                <%if oper="new" then
					val=request.form("data")
				else
					val=formatdatetime(rs("data"),2)
				end if %>
                <input name="data" type="text" id="data1" value="<%=val%>" size="40" class="richiesto" title="Inserisci la data della fattura, campo obbligatorio.">
                <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Conto</td>
              <td><%if oper="view" then%>
                <%=conto(rs("conto"))%>
                <%else%>
                <%if oper="new" then
					val=24
				else
					val=rs("conto")
				end if%>
                <select name="conto" class="richiesto" >
                  <option value="0">Nessuno</option>
                  <%
					for n=1 to conto(0)
					response.write "<option value='"&n&"'"
					if val=n then response.write " selected "
					response.write ">" & conto(n)&"</option>"
					next
					%>
                </select>
                <%end if%></td>
            </tr>
            <tr>
              <td valign="top">Totale Fattura</td>
              <td><%if oper="view" then
			   totale_fattura=cdbl(rs("totale_fattura"))
			   %>
                <%=formatcurrency(rs("totale_fattura"),2)%>
                <%else%>
                <%if oper="new" then
					val=request.form("totale_fattura")
				else
					val=rs("totale_fattura")
				end if %>
                <input name="totale_fattura"  type="text"  value="<%=val%>" size="40" class="richiesto" title="Inserisci l'importo totale della fattura, campo obbligatorio.">
                <%end if %>
                </td>
            </tr>
            <%if oper="new" then %>
            <tr>
              <td valign="top">Acconto</td>
              <td>
               Importo <input name="acconto"  type="text"  value="" size="20" class="" > Data <input name="data_acconto" id="data3" type="text"  value="" size="12" class="" >
                </td>
            </tr>
            <%end if %>
                   <tr>
                    <td>Srl:</td>
                    <td><%val=""
                        			   if oper="new" then

                        				else
                        					if rs("srl")=1 then val=" checked"
                        				end if%>
            			<input name="srl" type="checkbox" value="1"  <%response.write val%>>
            			</td>
                  </tr>


            <%
	            totale_scadenze=conn.execute("select sum(importo) from scadenze_for where idfat="&idfat)(0)
	            if isnull(totale_scadenze) then
		            totale_scadenze=0
		        else
		            totale_scadenze=cdbl(totale_scadenze)
			    end if
              if totale_fattura>totale_scadenze or oper="new" then %>
              <tr>
	              <td>N&deg; scadenze</td>
	              <td>
		            <%if oper<>"new" then
		            	da_saldare=totale_fattura-totale_scadenze
		             %>
	              <input type="hidden" name="da_saldare" value="<%=da_saldare%>">
	              	<%end if %>

		              <input type="radio" name="n_scadenze" value="1" checked >1&nbsp;&nbsp;<input type="radio" name="n_scadenze" value="2">2&nbsp;&nbsp;<input type="radio" name="n_scadenze" value="3">3&nbsp;&nbsp;<input type="radio" name="n_scadenze" value="4">4 <input type="radio" name="n_scadenze" value="altro">
	                    <input name="n_scadenze_altro" type="text"  maxlength="2" size="2"/>


                </td>
              </tr>


            <tr>
              <td align="" valign="top">Data prima scadenza</td>
              <td>
	              <input type="radio" name="scadenza" value="<%=finemese(date(),1)%>" checked ><%=finemese(date(),1)%>&nbsp;&nbsp;<input type="radio" name="scadenza" value="<%=finemese(date(),2)%>"><%=finemese(date(),2)%>&nbsp;&nbsp;<input type="radio" name="scadenza" value="<%=finemese(date(),3)%>"><%=finemese(date(),3)%>&nbsp;&nbsp;<input type="radio" name="scadenza" value="<%=finemese(date(),4)%>"><%=finemese(date(),4)%>&nbsp;&nbsp;
              <input type="radio" name="scadenza" value="altro">
              <input name="scadenza_altro" type="text" id="data2"  value="<%=finemese(date(),0)%>" maxlength="10"  <%if view="view" then response.write " disabled"%> title="Inserisci la data di scadenza del pagamento, campo obbligatorio." />
              </td>
            </tr>
            <tr>
            	<td colspan="2">
	            	<input type="button" id="aggiungi_scadenze" value="Aggiungi scadenze">
            	</td>
            </tr>


          <%end if%>



            <%

  		tipo_allegato=1
		id_tipo_allegato=idfat



	            if oper="edit" or oper="new" then%>
            <tr>
              <td colspan="2" align="center"><%
				if oper="view" then
					txt=puls_new
				elseif oper="new" or error<>"" then
					txt="Aggiungi"
				elseif oper="edit" then
					txt="Modifica"
				end if
				cancellabile=true
				%>
                <%if true then%>
                <input type="submit" name="oper" value="<%=txt%>" <%if oper="new" then %> class="disabled" disabled <%end if%>>
                <%if oper="new" then %>
                <input type="submit" name="oper" value="Aggiungi altre" class="disabled" disabled>
                <%end if %>
                <%end if%>
                <%if oper<>"view" then%>
                <input type="submit" name="annulla" value="Annulla" onClick="annulla=true;">
                <%end if%>
                <%if oper="edit" then%>
                <input type="submit" name="elimina" value="Elimina" <%if cancellabile=false then%> disabled <%end if%>>
                <%
                set rs = Nothing
                end if%></td>
            </tr>
            <%end if %>
          </table>


      			<script>
			$(function() {

				$(document).on("click",".paga_scadenza",function() {
					var id=$( this ).closest("tr").attr("id" ).replace("scadenza_","");
					var obj = $(this);
					console.log( id);
					$.ajax({
			            url: "pag_adm_fatture_for.asp",
			            type: "post",
			            cache: false,
			            //dataType: 'json',
			            data	: {paga_scadenza: id, idfat: <%=idfat%>} ,
			            success: function(data) {
				          obj.closest('tr').css("background",'#99FF99');
				          obj.remove();
			            },
			            error: function(xhr, textStatus, error) {
			                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
			                console.log("xhr.statusText:" + xhr.statusText);
			                console.log("xhr.responseText:" + xhr.responseText);
			                console.log("textStatus:" + textStatus);
			                console.log("error:" + error);
			            }
			        });

				});

				$("#aggiungi_scadenze").click(function(e){
					var dati=$("#form1").serialize();
					$.ajax({
			            url: "pag_adm_fatture_for.asp",
			            type: "post",
			            cache: false,
			            //dataType: 'json',
			            data	: dati+"&aggiungi_scadenze=si" ,
			            success: function(data) {
			                $("#div_tabella_scadenze_for").html(data);
			                $(".disabled").prop("disabled",false);

			            },
			            error: function(xhr, textStatus, error) {
			                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
			                console.log("xhr.statusText:" + xhr.statusText);
			                console.log("xhr.responseText:" + xhr.responseText);
			                console.log("textStatus:" + textStatus);
			                console.log("error:" + error);
			            }
			        });


	          });
	          });


          </script>

          <div id="div_tabella_scadenze_for">
          <%


		if oper="view" then

			call tabella_scadenze_for(idfat)
		end if

  %></div>


          <%if oper="view" then%>
          <%


if oper="view" then


  %>
          <%if oper="view" then

		%>
		<div id="sub_allegati">
            <!--#include virtual="/sub_allegati.asp" -->
		</div>
          <%
			set rs = Nothing
          end if%>
          <%
		  end if
		  end if%>
		  		</form>

<!--#include virtual="/sub_dropzone.asp" -->


          <%
          call connclose()
end if	'if oper="list"
%>
          <!-- Colonna CORPO FINE-->
        <!-- Box CORPO FINE-->
      </div>
<%
	txt_timer=txt_timer&"T fine:"&FormatNumber(timer() - t_inizio, 2)&" "
%>
	<div id="footer"><%=txt_timer%></div>
	<!--#include virtual="/pag_adm_footer_inc.asp" -->
    </div>
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
function cerca_senza_scadenze()
	if application("cache_cerca_senza_scadenze")="" then
		dim txt,n
		sql="select Count(*)  FROM fatture_for LEFT JOIN scadenze_for ON fatture_for.idfat = scadenze_for.idfat WHERE scadenze_for.ID Is Null and fatture_for.idfor>0 and totale_fattura>0"
		n=clng(conn.execute(sql)(0))
		if n>0 then
			txt="<tr class=""coprobox""><td colspan=""7"" bgcolor=""red"">Attenzione: ci sono "&n&" fatture senza scadenze, <a href="""&questofile&"?cosa=fatture&cercain=noscadenze""><strong>vedi fatture</strong></a></td></tr>"
		else
			txt=""
		end if
		Application.Lock
		Application("cache_cerca_senza_scadenze")=txt
		Application.Unlock
	end if
	cerca_senza_scadenze=application("cache_cerca_senza_scadenze")
end function

%>
