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
tabella="preventivi"
'add2log queryeform(),0
if session("idadmin") = "" then call login()
if request("idord")<>"" then idord=request("idord")
mese=request.form("mese")
cercain=request("cercain")
cerca=request("cerca")

if request.form("recupera")<>"" then
	sql="select * FROM preventivi WHERE ordini.idord="& idord&";"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	comunicazioni_precedenti= now()&" <b>"&session("nominativo")&"</b>: ripristinato da eliminazione"
	if isnull(rs("comunicazioni_precedenti")) or len(trim(rs("comunicazioni_precedenti")))=0  then
		rs("comunicazioni_precedenti")=comunicazioni_precedenti
	else
		rs("comunicazioni_precedenti")=rs("comunicazioni_precedenti")&"<hr />"& comunicazioni_precedenti
	end if
		rs("eliminato")=0
	rs.update
	rs.close
	set rs = nothing
	oper="view"
end if



if request.form("aggiorna_note")<>"" then
	sql = "select * FROM "&tabella&" WHERE idord="&idord
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	rs("note_gestore")=trim(request.form("note_gestore"))
	rs.update
	rs.close
	set rs=Nothing
	oper="view"
end if

stato_form=request.form("stato")
if stato_form="" then
	stato_form=cint(-1)
else
	stato_form=cint(stato_form)
end if
if request.form("stato")<>"" or request.form("conferma")<>"" then
	Set rs_preventivo = Server.CreateObject("ADODB.Recordset")
	sql="select * FROM preventivi WHERE preventivi.idord="& idord&";"
	rs_preventivo.Open sql, conn, 3, 3
	comunicazioni_precedenti=""
	if request.form("stato")<>"" then
		invio_mail=false
		rs_preventivo("stato")=request.form("stato")
		select case request.form("stato")
		case 1
			rs_preventivo("data_stato1")=now()
		case 2
			rs_preventivo("data_stato2")=now()
		case 3 'confermato
			rs_preventivo("data_stato3")=now()
			invio_mail=true
		case 4
			rs_preventivo("data_stato4")=now()
		case 5
			rs_preventivo("data_stato5")=now()
		case 6 'ordine
			if isnull(rs_preventivo("data_stato6")) or rs_preventivo("data_stato6")="" then
				set ordine= (new ClasseOrdine)(array("copia_da","preventivi",idord,"ordini"))
				idord=ordine.idord()
				set ordine = Nothing
			end if
		end select
		comunicazioni_precedenti= now()&" <b>"&session("nominativo")&":</b> stato preventivo: "&stato_preventivo(rs_preventivo("stato"))
		if ok=true then comunicazioni_precedenti=comunicazioni_precedenti&", creato ordine "&idord_new
	end if
	if request.form("note_gestore")<>"" then
		rs_preventivo("note_gestore")=request.form("note_gestore")
	end if
	if request.form("conferma")<>"" or  invio_mail=true then
		rs_preventivo("conferma_ordine")=request.form("conferma_ordine")
		if comunicazioni_precedenti="" then
				comunicazioni_precedenti= now()&" <b>"&session("nominativo")&":</b> stato "&stato_preventivo(rs_preventivo("stato"))&"<br>Messaggio: "&request.form("conferma_ordine")
		else
				comunicazioni_precedenti=comunicazioni_precedenti& "<br>Messaggio: "&request.form("conferma_ordine")
		end if
		invio_mail=true
		
	end if
	if isnull(rs_preventivo("comunicazioni_precedenti")) or len(trim(rs_preventivo("comunicazioni_precedenti")))=0  then
		rs_preventivo("comunicazioni_precedenti")=comunicazioni_precedenti
	else
		rs_preventivo("comunicazioni_precedenti")=rs_preventivo("comunicazioni_precedenti")&"<hr />"& comunicazioni_precedenti
	end if
	rs_preventivo.update
	rs_preventivo.close
	set rs_preventivo=Nothing
	oper="list"
	'if invio_mail=true then email_ordine idord,"preventivo",true
	'call connclose()
	'response.redirect "pag_adm_ordini.asp?idord="&idord_new
end if
if request.form("operazione")="elimina_ordine" then
	set ordine= (new ClasseOrdine)(array("elimina",tabella,idord))
	set ordine = Nothing
	oper="list"
	cercain=""
	idord=""

end if

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
tabellaid="idord"


display_none="display:none;"
if request("anno")="" then
	anno=year(date())
else
	anno=request("anno")
end if
%>
<!--#include virtual="/regioni.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
  <title><%=application("brwstitle")%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

  <!--#include virtual="/sub_head.asp" -->
  <script src="Jquery/js/jquery.ui-contextmenu.min.js" type="text/javascript"></script>
  <script type="text/javascript" language="javascript">
<%=script_errore%>
function Valida_ricerca() 
{	posizione_errore="Valida_ricerca";
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
	if ((document.form1.cercain.value == 'azienda' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il nome azienda nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	document.form1.submit();
} 
function Valida_crea_ordine() 
{	
	posizione_errore="Valida_crea_ordine";
	if (document.form1.cerca_iduser.value == '' ) {
		alert("Selezionare un cliente");
		document.form1.cerca.focus();
		return false;
	}
	document.form1.operazione.value='crea_ordine';
	document.form1.submit();
} 

</script>
  <script Language="JavaScript"> 
annulla=false;
function Validator(theForm) 
{
posizione_errore="Validator";
if (annulla==true)  { 
	return(true);    
  } 
 
  if ($("#stato").val() == 3)  { 
 if(tinyMCE.get('conferma_ordine').getContent()==""){
	 alert("Inserisci il testo del messaggio di aggiornamento ordine."); 
	tinyMCE.editors['conferma_ordine'].focus();
	return (false); 
	 }
	  } 
  return (true); 
} 
</script>
  <style type="text/css">
.piccolo {
	font-size: 9px;
}
.pulsanti {
	padding: 3px;
	overflow: visible;
	margin: 1px;
	font-size: 10px;
}
span.red {
	color: red;
}

	


  </style>
  <script src="jquery/ui/i18n/jquery.ui.datepicker-it.min.js"></script>
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
<%if oper="list" then%>
$(function() {
	$( "#cerca" ).autocomplete({
		source: "ajax_function.asp?autocomplete=utenti",
		cache: false ,
		minLength: 2,
		select: function( event, ui ) {
			$('#cerca_hidden').val(ui.item.id); 
			//$('#form1').submit();
		}
	});
	$("#puls_cerca").focus();
});

<%end if%>

</script>
  <style>
.ui-menu {
	position: absolute;
	width: 200px;
	font-size: 8px;
}
.ui-menu-item {
	font-size: 12px;
}
#apDiv1 {
	position: absolute;
	width: 200px;
	height: 115px;
	z-index: 1;
}
.div_icona {
display: inline-block; float:left;
}
.right
{
 float:right;
}


/*CSS Validator  */
#form_ddt label.error {
	float: none; color: red; 
   padding-left: .5em;
   vertical-align: top; 
   display: block;
   width:100%;
   }
 #form_ddt input.error {
	 background-color: #F30;
   }  
 #form_ddt input.valid {
/*	 background-color: #0C3;*/
   }

.mieidettagli {
	display: none;
}


  </style>
  </head>
  <body>
  <div id="wrap">
    <div id="header"> <%=titolo_top%> 
      <!-- Box CORPO INIZIO-->
      <%barra=0%>
      <div id="barra_fissa"> 
        <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Preventivi</a></div>
      </div>
      <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" onSubmit="return Validator(this)">
        <!-- Box CORPO INIZIO-->
        <input type="hidden" name="operazione" id="operazione">
        
        <%
	        
	    response.write get_tabella_ordini("preventivi")
call report_errore()
if oper="list" then



		sql="select preventivi.*, i.nome, i.cognome, i.azienda, utenti.email, utenti.voto,utenti.telefono, utenti.fax, utenti.cellulare, utenti.tipologia FROM preventivi  INNER JOIN utenti ON preventivi.iduser = utenti.iduser inner join utenti_intestazioni i on preventivi.idintestazione = i.id "

		if cerca<>"" and isnumeric(cerca) and cercain="" then cercain="idord"
		iPageSize=20
		order= " order by ordini.anno desc, nord desc"
		where_fisso = " ordini.eliminato=0 "
		
		
		if not ha_il_permesso("C5") then
			where_fisso=where_fisso&" and idagente="&session("idagente")
		end if
		
		
		
		
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
		
		
		'-----------sezione ricerca inizio
		if request.form("cerca_iduser")<>"" and request.form("elimina_ordine")="" and cercain<>"niente" then
			cercain="iduser"
			cerca=request.form("cerca_iduser")
		end if
		if request("iduser")<>"" and request.form("elimina_ordine")="" and cercain<>"niente" then
			cercain="iduser"
			cerca=request("iduser")
		end if
		if request("ffm")<>"" then
			cercain="finemese"
			fine_mese=true
		end if
		
		if cercain<>"" then
			select case cercain
			case "stato0"
				where=" stato=0"
				where_anno=""
			case "stato1"
				where=" stato=1"
				where_anno=""
			case "stato2"
				where=" stato=2"
				where_anno=""
			case "stato3"
				where=" stato=3"
				where_anno=""
			case "stato4"
				where=" stato=4"
				where_anno=""
			case "stato5"
				where=" stato=5"
				where_anno=""
			case "stato6"
				where=" stato=6"
				where_anno=""
			case "stato7"
				where=" stato=7"
				where_anno=""
			case "idord"
				where=" ordini.idord="&cerca 
				where_anno=""
			case "iduser"
				where=" ordini.iduser="&cerca
				where_anno=""
			case "da_stampare"
				where=" da_stampare=true "
				where_anno=""
			case "eliminati"
				where_fisso = " ordini.eliminato=1 "
				where_anno=""
				cerca=""	
			case "azienda"
				where=" InStr(azienda,'" & cerca& "')>0"
				where_anno=""
			case "finemese"
				
				sql=sql&" WHERE (ordini_fatture.idord Is Null) and ddt.causale=1 "
				where=""
				order=""
				PaginazioneString="&ffm=si"
				'response.write "finemese"
			case else
				where=" true "
			end select
			
		else	'nessun parametro di ricerca allora applico filtro anno
			'where=where_fisso& where_anno
		end if
		'if where<>"" then where=" where "&where
		
		

		'PaginazioneString=PaginazioneString&"&cercain="&cercain&"&cerca="&cerca
		if where<>"" then
			where=where_fisso & " and " &where
		else
			where=where_fisso
		end if
			
		if where_anno<>"" and where<>""  then where_anno=" and "&where_anno
		'where=where_fisso&" and "&where&where_anno
		where=where&where_anno
		if where<>"" then where=" having "&where
	 	strSql = sql &where&order
	 	'response.write strSql
		PaginazioneString="&cercain="&cercain&"&cerca="&cerca&"&anno="&anno&"&cerca_hidden="&request("cerca_hidden")

		'-----------sezione ricerca fine
		if cercain<>"finemese" then strSql = sql &where&order
		'add2log strsql,1
		'response.write strSql
				sql="select Count(carrello.idcar) AS ConteggioDiidcar FROM carrello INNER JOIN prodotti ON carrello.idpro = prodotti.IDpro WHERE (((prodotti.Attivo)=True) AND ((prodotti.vendita)=True)   AND ((carrello.iduser)="&session("iduser")&"))"
				set rs=conn.execute(sql)
				carrello=clng(rs("conteggiodiidcar"))
				rs.close
				set rs = nothing

		%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1">
		 <%if carrello>0 then%>  
				   
            <tr>
              <td colspan="6">Ci sono articoli nel carrello: <strong>crea ordine cliente con <%=carrello%> articoli nel carrello</strong></td>
            </tr>
           
            <%end if%>
          <tr>
            <td colspan="6"> Cerca
              <input type="text" name="cerca" class="casella_ricerca" id="cerca" value="" />
              <input type="hidden" id="cerca_hidden" name="cerca_iduser">
              in
              <select name="cercain" style="FONT: 12px;" onChange="Valida_ricerca()">
                <option value="" <%if cercain="" then response.write "selected"%>>Visualizza tutti</option>
                <option value="idord" <%if cercain="idord" then response.write selected%>>N°preventivo</option>
                <option value="azienda" <%if cercain="azienda" then response.write selected%>>Azienda</option>
                <option value="iduser" <%if cercain="iduser" then response.write selected%>>ID cliente</option>
                <option value="eliminati" <%if cercain="eliminati" then response.write selected%>>Eliminati</option>
              </select>
              <input type="submit" name="Submit" value="Cerca" style="FONT: 12px;" onClick="Valida_ricerca()" id="puls_cerca">
              &nbsp;Anno:
              <input name="anno" type="text" style="FONT: 12px;" value="<%=anno%>" size="5" maxlength="5">
              <span style="float:right;">
          
              <%if ha_il_permesso("C1") then %>
			  <input type="button" value="+Nuovo..." class="nuovodocumento" id="nuovo_preventivo">
              <%end if %>
              </span></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#E5E5E5"><p> Edit</p></td>
            <td align="center" bgcolor="#E5E5E5">N&deg; ordine/Data</td>
            <td bgcolor="#E5E5E5">Stato/Tot. ordine</td>
            <td align="left" bgcolor="#E5E5E5">Azienda/Indirizzo/Citt&agrave;</td>
            <td bgcolor="#E5E5E5">Nominativo/Email/tel</td>
            <td align="center" bgcolor="#E5E5E5">Documenti</td>
          </tr>
          <%
	strsql=replace(strsql,"ordini",tabella)
call paginazione_start(strsql,ipagesize,"access")

Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
	
	totale=cdbl(objPagingRS("totale"))

%>
          <tr class="coprobox" <%=colore%>>
            <td align="center" style="border-top:1px solid;"><input type="radio" name="modifica" value="<%=objPagingRS("idord")%>" onClick="this.form.submit()"></td>
            <td align="center" style="border-top:1px solid;"><%=objPagingRS("nord")%><br>
              <%=objPagingRS("data")%></td>
            <td align="center" style="border-top:1px solid; " class="statop_<%=objPagingRS("stato")%>"><b><%=stato_preventivo(objPagingRS("stato"))%></b><br>
              <b><%=formatcurrency(objPagingRS("totale"),2)%></b> <%=testo_differenza%></td>
            <td style="border-top:1px solid;"><%=vedi_tipologia(objPagingRS("tipologia"))%><b><%=denominazione(objPagingRS("nome"),objPagingRS("cognome"),objPagingRS("azienda"))%></b>
              </td>
            <td style="border-top:1px solid;"><b><%=objPagingRS("cognome")%>&nbsp;<%=objPagingRS("nome")%></b><br>
              <a href='mailto:<%=objPagingRS("email")%>'><%=objPagingRS("email")%></a><br>
              <%=objPagingRS("telefono")%>-<%=objPagingRS("cellulare")%></td>
            <td align="center" style="border-top:1px solid;">
            
            </td>
          </tr>

          <% 	
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop	
	call paginazione_end(6,true)
	
%>
        </table>
    <%else
'***********************************************************************************************************************************************
'****************************ORDINE SINGOLO
'***********************************************************************************************************************************************

		cancellabile=true
		cambia_stato=true
		set ordine= (new ClasseOrdine)(array("apri","preventivi",idord))
		
		
		
		dim vecchio_conteggio
		
		if not ha_il_permesso("C5") and session("idagente")<>ordine.campo("idagente") then
			call via_senza_permesso("C5")
		end if
		
		
		
		idord=ordine.campo("idord")
		stato=ordine.campo("stato")
		nord=ordine.campo("nord")
		iduser=ordine.campo("iduser")
		prenotazione= prenota_record("preventivo",idord,"",false)
		
        sql="select ordini_dett.*, magazzino.quantita_magazzino, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, varianti_a.variante_a, varianti_b.variante_b FRom ((magazzino RIGHT JoIN ordini_dett oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JoIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara where idord="&idord &" order by ordine,iddett;"
		set rs=conn.execute(sql)
		if rs.eof then fase="due"
		data_documento=ordine.campo("data")
		if DateDiff("d","01/01/2015",data_documento)<0 then 'prima del 01/01/2015
			vecchio_conteggio=true
		else
			vecchio_conteggio=false
		end if
		call ordine.avviso_eliminato()

 %>
 	<div id="multitabs" style="display:none;">
        <ul>
          <li><a href="#tabs-1" id="first_tab">Preventivo <%=nord%></a></li>
		  <li><a href="#tabs-2">Cronologia</a></li>
		  <li><a href="pag_adm_user.asp?tab=1&iduser=<%=iduser%>">Dati cliente</a></li>
		  <%if idddt<>"" then %>
		  <li><a href="pag_adm_ddt.asp?tab=1&idddt=<%=idddt%>"><strong>DDT <%=nddt%></strong> del <%=formatdatetime(data_ddt,2)%></a></li>
		  <%end if%>
		  <%if idfat<>"" then %>
		  <li><a href="pag_adm_fatture.asp?tab=1&idfat=<%=idfat%>"><strong>Fattura <%=nfat&post_fattura%></strong> del <%=formatdatetime(data_fattura,2)%></a></li>
		  <%end if%>
		   <li id="tab_altro"><a href="#">Altro...</a></li>
		</ul>
	<div id="tabs-1" class="tab-panel">
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
          <tr>
            <td colspan="2" class="ui-widget-header">Preventivo
              <input name="idord" type="hidden" value="<%=idord%>" id="idord">
              <input name="nord" type="hidden" value="<%=nord%>" id="nord">
              <span style="float:right;">
              
              <INPUT type="button" value="Stampa" onClick="window.open('pag_adm_ordini_printer.asp?cosa=ordine&idord=<%=idord%>','mywindow','width=400,height=200')" class="ui-button ui-widget ui-state-default ui-corner-all">
              <INPUT type="button" value="PDF" onClick="location.href='pdf_ordine_preventivo.asp?tabella=preventivi&idord=<%=idord%>'" class="ui-button ui-widget ui-state-default ui-corner-all">
				<INPUT type="button" value="Modifica"  data-id="<%=idord%>" data-documento="preventivo" class="pulsante_modifica_intestazione ui-button ui-widget ui-state-default ui-corner-all puls_">              <button class="puls_giu" style="display:none;">Altro...</button>
              </span>
              <ul style="display:none;">
                <%if session("iduser")=1 then%>
				<li><a href="pag_adm_critici.asp?cosa=ordine&id=<%=idord%>" id="show_registro">Cerca nel registro</a></li>
                <%end if%>
                <li><a href="#" id="allega_files">Allega files...</a></li>
                <li><a href="#" id="duplica" class="oper_articoli">Duplica</a>
                <li ><a href="pag_adm_user.asp?iduser=<%=ordine.campo("iduser")%>">Vedi scheda cliente</a></li>
              </ul></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#E5E5E5">N&deg; preventivo/Stato/Data</td>
            <td align="left" bgcolor="#E5E5E5">Nominativo/Email/Telefoni</td>
          </tr>
          <tr >
            <td align="center" style="border-bottom:1px solid;"><b><%=ordine.campo("nord")%></b><%if session("iduser")=1 then%>/<%=ordine.campo("idord")%><%end if%><br>
              <b><%=stato_preventivo(stato)%></b><br>
              <%=ordine.campo("data")%>
              <%if ordine.campo("creato_da_admin")>=0 then response.write "<br>Creato da "&nome_admin(ordine.campo("creato_da_admin"))%>
              
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
            <td align="right" class="ui-widget-header"  style="border-left: none;"><%if session("iduser")=1 then%><input type="submit" name="ricalcola" value="Ricalcola" style="font-size:8pt"><%end if%>
          </tr>
          
          <%
			call ordine.dati_fatturazione_consegna()
	        call ordine.dati_mepa()  
	          
	      %>
          
        </table>
        <%
		if ordine.campo("note")<>"" then%>
        <div class="div_titolo">Note cliente sull'ordine</div>
        <div class="div_testo"><%=ordine.campo("note")%></div>
	    <div class="div_spaziatore"></div>
        <%end if%>
	  
<!-- NOTE GESTORE -->
<%
      	val=ordine.campo("note_gestore")
		if val="" or isnull(val) then
		nascondi_nota="style=""display:none;"""
		end if
      %>
		          <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
          <tr>
            <td colspan="2" class="ui-widget-header">Note (nascosto al cliente):  <input type="button" id="pulsante_tabella_nota" onClick="showhide_tabella_note()" value="Aggiungi nota" style="font-size:8pt;" class="right"><%if nascondi_nota="" then%> <input type="button" id="pulsante_tabella_modifica_nota" onClick="showhide_tabella_modifica_note()" value="Modifica nota" style="font-size:8pt;" class="right"><%end if%>
            </td>
          </tr>
          <tr valign="middle" align="right" <%=nascondi_nota%> id="note">
            <td align="left" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;" id="cella_note_gestore"><%=val%></td>
          </tr>
          <tr valign="middle" align="right" id="tabella_note" style="display:none;">
            <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
			<!--#include virtual="/tinymce4_conf2.asp" -->
              <textarea id="note_gestore" class="mceEditor" rows="6" style="width:100%;"></textarea>
              <label>
                <input type="button"  id="salva_nota" value="Salva">
              </label></td>
          </tr>
        </table>
<!-- ALLEGATI -->
<%tipo_allegato=11
id_tipo_allegato=idord%>
<!--#include virtual="/sub_allegati.asp" -->
<!-- DETTAGLIO ORDINE -->
	<input type="hidden" name="ordinati" id="ordinati" />
	<%if ordine.campo("calcolato")=0 then response.write "<b>DA RICALCOLARE</b>"%>
	<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1" id="tab_articoli">
		<%
			da_saldare=false
			ordine.modifica_articoli=true
			if not ha_il_permesso("C1") then
				ordine.modifica_articoli=false
			end if
			
			ordine.elenco_articoli_head()
			ordine.mostra_spettanze=ordine.campo("spettanze")
			ordine.elenco_articoli() 
			ordine.elenco_totali() %>
          <%if stato<=2 and ordine.modifica_articoli and ha_il_permesso("C1") then%>
          <tr valign="middle" >
            <td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
              <input type="submit" value="Sposta o copia" id="sposta_articoli" class="oper_articoli"><%if ordine.modifica_articoli then%><span style="float: right;"><input type="button" value="Aggiungi articolo" id="puls_aggiungi_articolo"><input type="submit" name="Aggiorna_ordine" id="Aggiorna_ordine" value="Aggiorna ordine"></span><%end if%></td>
          </tr>          
          <%end if%>
        </table>
       		
        
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
          <tr>
            <td colspan="3" class="ui-widget-header">Invia messaggio:</td>
          </tr>
          <tr valign="middle" align="right" >
            <td colspan="3" align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"></td>
          </tr>
          <tr valign="middle" align="right" >
            <td colspan="3" align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><%
      	val=ordine.campo("conferma_ordine")
      %>
              <textarea name="conferma_ordine" id="conferma_ordine" class="mceEditor" rows="6" style="width:100%;" ><%'=val%>
</textarea></td>
          </tr>
          <tr valign="middle" align="right" >
            <td width="33%" align="left" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><script>
risposte = new Array(); 
<%
set rsm=conn.execute("select * from ordini_risposte_predefinite where tipo=1 order by nome_risposta ;")
do while not rsm.eof
testo_risposta=replace(rsm("testo_risposta"),"'","\'")
testo_risposta=replace(testo_risposta,"&","\&")
testo_risposta=replace(testo_risposta,vbCrLf,"<br>")
%>risposte [<%=rsm("idrisposta")%>]='<%=testo_risposta%>';
		<%
rsm.movenext
loop
rsm.close
%>
function Inseriscirisposta(){
		tinyMCE.editors['conferma_ordine'].setContent(risposte[document.form1.risposta.options[document.form1.risposta.selectedIndex].value]);
}
</script>
              <select name="risposta" id="risposta" onChange="Inseriscirisposta();">
                <option value="">Scegli una risposta</option>
                <%
				set rsm=conn.execute("select * from ordini_risposte_predefinite where tipo=1 order by nome_risposta ;")
				do while not rsm.eof
				response.write "<option value=""" & rsm("idrisposta") &""""&">" & rsm("nome_risposta")& "</option>"
				rsm.movenext
				loop
				rsm.close
				set rsm=nothing
				
				%>
              </select></td>
            <td width="33%" align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><input type="submit" name="conferma" id="button2" value="Invia messaggio"></td>
            <td width="33%" align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">&nbsp;</td>
          </tr>
        </table>
        <%if ha_il_permesso("C2") then %>
          <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
            <tr>
              <td colspan="4" class="ui-widget-header">Stato avanzamento:</td>
            </tr>
            <%
			if cambia_stato=false then disabilitato=" disabled"
			%>
            <tr>
              <td width="20%" align="center" valign="bottom" ><input type="hidden" name="stato" />
                <%=ordine.campo("data_stato1")%><br>
                <button type="submit" name="" class="pulsanti" value="<%=1%>" <%if ordine.campo("stato")=1 then%> style="background-color:green;" <%end if%> onClick="stato.value=<%=1%>; return;" <%=disabilitato%>><%=stato_preventivo(1)%></button></td>
              <td width="20%" align="center" valign="bottom" ><%=ordine.campo("data_stato2")%><br>
                <button type="submit" name="" class="pulsanti" value="<%=2%>" <%if ordine.campo("stato")=2 then%> style="background-color:green;" <%end if%> onClick="stato.value=<%=2%>; return;" <%=disabilitato%>><%=stato_preventivo(2)%></button>
              </td>


              <td width="20%" align="center" valign="bottom" ><%=ordine.campo("data_stato5")%><br>
                <button type="submit" name="" class="pulsanti" value="<%=5%>" <%if ordine.campo("stato")=5 then%> style="background-color:green;" <%end if%> onClick="stato.value=<%=5%>; return;" <%=disabilitato%>><%=stato_preventivo(5)%></button></td>
              <td width="20%"  align="center" valign="bottom" ><%=ordine.campo("data_stato6")%><br>
                <button type="submit" name="" class="pulsanti" value="<%=6%>" <%if ordine.campo("stato")=6 then%> style="background-color:green;" <%end if%> onClick="stato.value=<%=6%>; return;" ><%=stato_preventivo(6)%></button>
              </td>
            </tr>
            <tr>
              <td width="20%" align="center" >&nbsp;</td>
              <td width="20%" align="center" ><span class="piccolo"><br>
              </span></td>
              <td width="20%" align="center" >&nbsp;</td>
              <td width="20%" align="center" ><span class="piccolo">Trasforma in ordine              </span></td>
            </tr>
            <tr>
              <td colspan="6" align="right" >
	              
	               <%if cancellabile=true and ha_il_permesso("C3") then%><button type="submit" name="" value="<%=7%>" <%if ordine.campo("stato")=7 then%> style="background-color:red;" <%end if%> onClick="stato.value=<%=7%>; return;"><%=stato_preventivo(7)%></button>
              <%end if %>
			  <%if ha_il_permesso("C4") then %>
                <input  type="button" value="Elimina" onClick="DialogYesNo ('Confermi l\'eliminazione del preventivo?<%=txt_seriali%>',elimina_ordine)" style="font-size:8pt" <%if cancellabile=false then response.write " disabled "%>>
                <%end if%></td>
            </tr>
          </table>
          <%end if %>
          
          		</div>
		<div id="tabs-2" class="tab-panel">
		<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
          <tr>
            <td class="ui-widget-header">Cronologia:</td>
          </tr>
          <tr valign="middle" align="right">
            <td align="left" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><%=ordine.campo("comunicazioni_precedenti")%></td>
          </tr>
        </table>
		</div>
		</div>
      </form>

	  <%
		  spettanze=ordine.campo("spettanze")
		  set ordine = nothing
		  call connclose()
		  end if%>
      <!-- Box CORPO FINE--> 
    </div>
<%
	txt_timer=txt_timer&"T fine:"&FormatNumber(timer() - t_inizio, 2)&" "
%>
<div id="footer"><%=txt_timer%></div>
<!--#include virtual="/pag_adm_footer_inc.asp" -->
  </div>
  
  <%
  if oper="view" then
  %>
<script type="text/javascript" src="Jquery/js/idle.js"></script>
<script type="text/javascript" src="Jquery/js/autogrow.min.js"></script>
<script>
	
	
var documento="preventivo";
var idDocumento=$("#idord").val();
var nDocumento=$("#nord").val();
var cosa="preventivo";
var inattivo=false;
var idadmin=<%=session("iduser")%>;
var iduser=<%=iduser%>;
var t_awayTimeout=<%=Application("rl_t_idle")*60%>000;
var rl_t_refresh=<%=Application("rl_t_refresh")%>000;
var prenotazione="<%=prenotazione%>";
var questofile="<%=questofile%>";
var aggiorna_ordine;
var iframe;
var tab_iduser=<%=iduser%>;
var tab_escludi="";
var tipo_allegato=<%=tipo_allegato%>;
var spettanze=<%=spettanze%>;

<%if request("avviso")=1 then%>
toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
toastr.info('L\' ordine che stavi visualizzando è stato rilasciato perchè richiesto da un altro utente');
<%end if%>
</script>
<script type="text/javascript" src="pag_adm_ordini.js?<%=ver_js_css %>"></script>
<script>
function enableSaveBtn(){
	return true;
}
$(function() {
	doAutogrow();
	puls_giu();
    $("#prova").on("click", function (e) {
        e.preventDefault();
        var tabs = $( "#tabs" ).tabs();
		var ul = tabs.find( "ul" );
		$( "<li><a href='public/manuali/Manuale_norme_generali_di_sicurezza_per_tutte_le_macchine_laser_co2_153418.pdf'>New Tab</a></li>" ).appendTo( ul );
		tabs.tabs( "refresh" );
    });
	

	$("#provaio").click(function(e) {
		toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
		var note_gestore = tinyMCE.get('note_gestore').getContent({format : 'html'});
			toastr.error(note_gestore);
		 note_gestore = encodeURIComponent(tinyMCE.get('note_gestore').getContent({format : 'html'}));
			toastr.error(note_gestore);
    });
	$("#iva_esclusa").on("click", function (e) {
		$("#mlc").toggle();
    });
	$("#mlc").on("click", function (e) {
				$.ajaxSetup({ cache: false });
				$.get("pag_adm_ordini_ajax.asp?tabella="+tabella+"&oper=tbody&idord="+idord+"&nord="+nord+"&mlc=si",function(result){
					$("#tab_articoli_body").remove();
					$("#tab_articoli_totali").remove();
					$('#tab_articoli_testa').after(result);
					toastr.options = {"positionClass": "toast-bottom-right"};
					toastr.success('Articoli aggiornati');
				 });	
    });
	

});
	
function showhidearticoli()
{
posizione_errore="showhidearticoli";
  var elem = document.getElementById('articoli_originali');
  
  if (elem.style.display == 'none'){
  document.getElementById('pulsante_articoli').value = 'Nascondi articoli originali';
    elem.style.display = 'block';}
  else{
  document.getElementById('pulsante_articoli').value = 'Mostra articoli originali';
    elem.style.display = 'none';}
}

function showhide_tabella_note()
{
posizione_errore="showhide_tabella_note";
	if($('#tabella_note').css('display') == 'none'){ 
		$('#tabella_note').fadeIn(800);
		$('#pulsante_tabella_nota').prop('value', 'Nascondi');
		tinyMCE.get('note_gestore').setContent("");
		$("#salva_nota").prop('value', 'Salva');

	}
	else
	{
		$('#tabella_note').fadeOut(200);
		$('#pulsante_tabella_nota').prop('value', 'Aggiungi nota'); 
	}
}
function showhide_tabella_modifica_note()
{
posizione_errore="showhide_tabella_modifica_note";
	if($('#tabella_note').css('display') == 'none'){ 
		$('#tabella_note').fadeIn(800);
		$('#pulsante_tabella_modifica_nota').prop('value', 'Nascondi'); 
		tinyMCE.get('note_gestore').setContent($("#cella_note_gestore").html());
		$("#salva_nota").prop('value', 'Modifica');
	}
	else
	{
		$('#tabella_note').fadeOut(200);
		$('#pulsante_tabella_modifica_nota').prop('value', 'Modifica nota');
		tinyMCE.get('note_gestore').setContent("");
	}
}

document.title = "Preventivo <%=nord%>";
</script>
<%end if%>
<%if oper="list" then%>
<script>
document.title = "Preventivi";
var tabella="preventivi";
var cosa="preventivo";
</script>
<%end if%>
</body>
</html>

<%
call CheckConnChiusa()
%>