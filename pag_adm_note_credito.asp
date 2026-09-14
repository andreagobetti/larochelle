<%
'Verifica chiusure 04_12_2015
%>
<%t_inizio=timer()%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/chk_piva_cf.asp" -->
<!--#include virtual="/paginazione.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->
<!--#include virtual="/classeMagazzino.asp" -->
<!--#include virtual="/classeElencoArticoli.asp" -->
<%
tabella="notacredito"
tipo_documento="notacredito"
'add2log queryeform(),0
if session("idadmin") = "" then call login()
if request("idord")<>"" then idord=request("idord")
mese=request.form("mese")
cercain=request("cercain")
cerca=request("cerca")
operazione=request.form("operazione")
Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class

if request.form("recupera")<>"" then
	sql="select * FROM ordini WHERE ordini.idord="& idord&";"
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

stato_form=request.form("stato")
if stato_form<>"" then
	set rs=conn.execute("select ordini.* from ordini where idord="&idord)
	if stato_form>-1 and stato_form<>cint(rs("stato")) and stato_form<>"" then
		'ordine.campo("stato")=stato_form
		
		sql="update ordini set stato="&stato_form&" where idord="&idord
		conn.execute(sql)
		select case stato_form
		case 1
	
		case 6
			n=metodipagamento.genera_scadenze(0,0,idord,rs("data").value,cdbl(rs("totale")),rs("tipopagamento").value)
		end select
	end if
	set rs = Nothing
end if




if request("elimina_nota")<>"" then
	idord=request("elimina_nota")
	set ordine= (new ClasseOrdine)(array("eliminatutto",tabella,idord))
	
	set ordine = Nothing
		'sql="DELETE   FROM ordini_dett WHERE idord=" & m_idord
		'conn.execute(sql)
		'sql="DELETE   FROM incassi WHERE idord=" & m_idord
		'conn.execute(sql)
		'sql="DELETE   FROM ordini WHERE idord=" & m_idord
		'conn.execute(sql)

	
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
 <%if mod_larochelle="" then %>
  if ($("#stato").val() == 3)  { 
 if(tinyMCE.get('conferma_ordine').getContent()==""){
	 alert("Inserisci il testo del messaggio di aggiornamento ordine."); 
	tinyMCE.editors['conferma_ordine'].focus();
	return (false); 
	 }
	  } 
	  <%end if %>
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
   
   
/*Stile per label numeri seriali  */
#form_aggiungi_seriali label {
    white-space:nowrap;
    border-bottom: 2px solid green;
    
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
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Note di credito</a></div>
      </div>
      <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" onSubmit="return Validator(this)">
        <!-- Box CORPO INIZIO-->
        <input type="hidden" name="operazione" id="operazione">
        
        <%
	        
call report_errore()
if oper="list" then


	
	sql="select ordini.*,utenti_intestazioni.nome, utenti_intestazioni.cognome, utenti_intestazioni.azienda,utenti.email, utenti.cellulare, utenti.telefono,utenti.tipologia, scadenze.totscadenze, scadenze.totpagato FROM ordini LEFT JOIN (select idord, count(*) as totscadenze, sum(pagato) as totpagato from scadenze group by idord) as scadenze ON ordini.idord = scadenze.idord LEFT JOIN utenti ON ordini.iduser = utenti.iduser left join utenti_intestazioni on ordini.idintestazione = utenti_intestazioni.id "

		if cerca<>"" and isnumeric(cerca) and cercain="" then cercain="idord"
		iPageSize=20
		order= " order by ordini.anno desc, nord desc"
		
		where_fisso = " ordini.eliminato=0 and tipo_documento='notacredito'"
		
		if not ha_il_permesso("C5") then
			where_fisso=where_fisso&" and creato_da_admin="&sessioniduser
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
		if request.form("cerca_iduser")<>""  and cercain<>"niente" then
			cercain="iduser"
			cerca=request.form("cerca_iduser")
		end if
		if request("iduser")<>"" and cercain<>"niente" then
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
			case "saldare"
				where=" totale>0 and (cast(totale as decimal (10,2))>sommadiimporto or sommadiimporto is null) and stato<7"
				where_anno=""
				cerca=""
			case "saldateko"
				where=" cast(totale as decimal (10,2))<sommadiimporto"
				where_anno=""
				cerca=""	
			case "eliminati"
				where_fisso = " ordini.eliminato=1 "
				where_anno=""
				cerca=""	
			case "azienda"
				where=" InStr(azienda,'" & cerca& "')>0"
				where_anno=""
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
	 	'response.write strSql
		PaginazioneString="&cercain="&cercain&"&cerca="&cerca&"&anno="&anno&"&cerca_hidden="&request("cerca_hidden")

		'-----------sezione ricerca fine
		strSql = sql &where&order
		'if utente_andrea then 
		'	response.write strsql
		'end if
		sql="select Count(carrello.idcar) AS ConteggioDiidcar FROM carrello INNER JOIN prodotti ON carrello.idpro = prodotti.IDpro WHERE (((prodotti.Attivo)=True) AND ((prodotti.vendita)=True)   AND ((carrello.iduser)="&session("iduser")&"))"
		set rs=conn.execute(sql)
		carrello=clng(rs("conteggiodiidcar"))
		rs.close
		set rs = nothing

		%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1" id="tabella_elenco_ordini">
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
                <option value="idord" <%if cercain="idord" then response.write selected%>>N°ordine</option>
                <option value="azienda" <%if cercain="azienda" then response.write selected%>>Azienda</option>
                <option value="iduser" <%if cercain="iduser" then response.write selected%>>ID cliente</option>
                <option value="saldare" <%if cercain="saldare" then response.write selected%>>Da saldare</option>
                <option value="saldateko" <%if cercain="saldateko" then response.write selected%>>Incassi incongruenti</option>
                <option value="eliminati" <%if cercain="eliminati" then response.write selected%>>Eliminati</option>
              </select>
              <input type="submit" name="Submit" value="Cerca" style="FONT: 12px;" onClick="Valida_ricerca()" id="puls_cerca">
              &nbsp;Anno:
              <input name="anno" type="text" style="FONT: 12px;" value="<%=anno%>" size="5" maxlength="5">
              <span style="float:right;">
              <%if ha_il_permesso("B1") then %>
              <input type="button" name="ffm" value="Cerca per fatture fine mese" style="FONT: 12px;" onclick="location.href='pag_adm_fine_mese.asp'">
              <%end if %>
              
              <%if ha_il_permesso("C1") then %>
              <input type="button" value="Nuovo..." id="pulsante_nuovo_ordine" style="display: none;">
				<input type="button" value="+Nuovo..." class="nuovodocumento" id="nuovo_notacredito">
              <%end if %>
              </span></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#E5E5E5"><p> Edit</p></td>
            <td align="center" bgcolor="#E5E5E5">N&deg;/Data</td>
            <td bgcolor="#E5E5E5">Totale</td>
            <td align="left" bgcolor="#E5E5E5">Azienda/Indirizzo/Citt&agrave;</td>
            <td bgcolor="#E5E5E5">Nominativo/Email/tel</td>
            <td align="center" bgcolor="#E5E5E5">Documenti</td>
          </tr>
          <%
call paginazione_start(strsql,ipagesize,"access")

Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
'	testo_differenza=""
	
	totscadenze=objPagingRS("totscadenze")
	if  not isnull(totscadenze) then 
		totscadenze=cdbl(totscadenze)
	end if
	totpagato=objPagingRS("totpagato")
	if  not isnull(totpagato) then 
		totpagato=cdbl(totpagato)
	end if
'	
'	totale=cdbl(objPagingRS("totale"))
'	if objPagingRS("verde")=1 then
'		colore="bgcolor='#CCFF99'"	'Verdino chiaro reso tutto/verde
'	elseif totale=0 and objPagingRS("stato")>5 then
'		colore="bgcolor='#99FF99'" 'Verdino saldata per importo 0
'	elseif totale<0 then
'		colore="bgcolor='#99FF99'"
'	elseif isnull(sommadiimporto) then
'
'		colore=""
'	elseif totale=sommadiimporto then
'		colore="bgcolor='#99FF99'" 'Verdino saldata
	if totscadenze=totpagato then
		colore="bgcolor='#99FF99'" 'Verdino saldata
		
'		testo_differenza="("&formatcurrency(sommadiimporto-totale,2)&")"
	else
		colore="bgcolor='#FFCCCC'" 'Rosino da saldare
'		testo_differenza="(+"&formatcurrency(sommadiimporto-totale,2)&")"
	end if
	'colore=""

%>
          <tr class="coprobox" <%=colore%>>
            <td align="center" style="border-top:1px solid;"><input type="radio" name="modifica" value="<%=objPagingRS("idord")%>" onClick="this.form.submit()"></td>
            <td align="center" style="border-top:1px solid;"><b><%=objPagingRS("nord")%></b><br>
              <%=objPagingRS("data")%></td>
            <td align="center" style="border-top:1px solid; " class="stato_<%=objPagingRS("stato")%>">
              <b><%=formatcurrency(objPagingRS("totale"),2)%></b> <%=testo_differenza%></td>
            <td style="border-top:1px solid;"><%=vedi_tipologia(objPagingRS("tipologia"))%><b><%=denominazione(objPagingRS("nome"),objPagingRS("cognome"),objPagingRS("azienda"))%></b></td>
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
		if nord<>"" then
			sql1="select idord FROM ordini "		
			if instr(anno,">")>0 then
				where_anno  = "  year(ordini.data)>="&replace(anno,">","")&" "
			else
				where_anno  = "  year(ordini.data)="&anno&" "
			end if	
			sql1=sql1&"  where nord="&nord &"  and "&where_anno&" "
			set rs=conn.execute(sql1)
			idord=rs("idord")
			Set rs = Nothing
		end if
		
		set ordine= (new ClasseOrdine)(array("apri","ordini",idord))
		
		

		
		
		
		
		
		'Set rs1 = conn.execute(sql1)
		'rs1.open sql1,conn,3,3
		
		dim vecchio_conteggio
		idord=ordine.campo("idord")
		stato=ordine.campo("stato")
		nord=ordine.campo("nord")
		iduser=ordine.campo("iduser")
		prenotazione= prenota_record("ordine",idord,"",false)
		
        sql="select ordini_dett.*, magazzino.quantita_magazzino, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, varianti_a.variante_a, varianti_b.variante_b FRom ((magazzino RIGHT JoIN ordini_dett oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JoIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara where idord="&idord &" order by ordine,iddett;"
		set rs=conn.execute(sql)
		if rs.eof then fase="due"
		data_documento=ordine.campo("data")
		'Cerco DDT
		sql="select ordini.idord, ordini.nord, ordini.data,ddt.causale, ddt.reso_tutto FROM ddt inner join ordini on ddt.idord = ordini.idord where sub_idord="&idord
		set rs_ddt=conn.execute(sql)
		if not rs_ddt.eof then
			idddt=rs_ddt("idord")
			nddt=rs_ddt("nord")
			data_ddt=rs_ddt("data")
			esiste_ddt=true
			Set ccausale_ddt = new cl_causale_ddt
			if ccausale_ddt.verde(rs_ddt("causale")) then
				verde=ccausale_ddt.elemento(rs_ddt("causale"))
			elseif rs_ddt("reso_tutto") then
				verde="RESO TUTTO IL MATERIALE"
			end if
			
			set ccausale_ddt = Nothing
			
		else
			esiste_ddt=false
		end if
		set rs_ddt = Nothing
		sql="select * FROM fatture where  idord="&idord
		sql="select fatture.*, ordini_fatture.idord FROM fatture INNER JOIN ordini_fatture ON fatture.IDfat = ordini_fatture.idfat where  ordini_fatture.idord="&idord

		set rs_fatture=conn.execute(sql)
		if not rs_fatture.eof then
			idfat=rs_fatture("idfat")
			nfat=rs_fatture("nfat")
			data_documento=rs_fatture("data")
			data_fattura=data_documento
			if rs_fatture("pa") then post_fattura=" PA" else post_fattura=""
			esiste_fattura=true
			cancellabile=false
			cambia_stato=false
		else
			esiste_fattura=false
		end if
		set rs_fatture = Nothing
		
		
		
		call ordine.avviso_eliminato()
 %>

 
 
 	<div id="multitabs" style="display:none;">
        <ul>
          <li><a href="#tabs-1" id="first_tab">Nota di credito <%=nord%></a></li>
		  <li><a href="pag_adm_ordini_ajax.asp?tabella=ordini&oper=cronologia&idord=<%=idord%>">Cronologia</a></li>
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
            <td colspan="2" class="ui-widget-header">Nota di credito
              <input name="idord" type="hidden" value="<%=idord%>" id="idord">
              <input name="nord" type="hidden" value="<%=nord%>" id="nord">
              <span style="float:right;">
              <%if false then %>
              <input type="button" value="QR" id="pulsante_qr" class="ui-button ui-widget ui-state-default ui-corner-all">
              <%end if %>
              
              
              
              
              
              <INPUT type="button" value="PDF" onClick="location.href='pdf_nota_credito.asp?tabella=ordini&idord=<%=idord%>'" class="ui-button ui-widget ui-state-default ui-corner-all">
              <button class="puls_giu" style="display:none;">Altro...</button>
              </span>
              <ul style="display:none;">
                <li ><a href="pag_adm_user.asp?iduser=<%=ordine.campo("iduser")%>">Vedi cliente</a></li>
                <li><a href="#" id="pulsante_modifica_intestazione">Modifica </a></li>
                
                <%
				sql="select Max(nord) FROM ordini where tipo_documento='notacredito' and anno="&year(date())
				max_ddt=conn.execute(sql)(0)
				if max_ddt=nord  then
				  %>
				  
				<li><a href="#" onClick="DialogYesNo ('Confermi l\'eliminazione della nota di credito?',f_elimina_notacredito)">Elimina</a>
                <%end if%>

              </ul></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#E5E5E5">N&deg; ordine/Stato/Data</td>
            <td align="left" bgcolor="#E5E5E5">Nominativo/Email/Telefoni</td>
          </tr>
          <tr >
            <td align="center" style="border-bottom:1px solid;" class="stato_<%=ordine.campo("stato")%>"><b><%=ordine.campo("nord")%></b><%if session("iduser")=1 then%>/<%=ordine.campo("idord")%><%end if%><br>
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
<!-- ALLEGATI -->
<%tipo_allegato=3
id_tipo_allegato=idord%>
<!--#include virtual="/sub_allegati.asp" -->


<!-- DETTAGLIO ORDINE -->
	<input type="hidden" name="ordinati" id="ordinati" />
	<%if ordine.campo("calcolato")=0 then response.write "<b>DA RICALCOLARE</b>"%>
		
		<%
		if 	ordine.campo("stato")=6 then
		ordine.modifica_articoli=false
		else
		ordine.modifica_articoli=true
		end if
		call ordine.elenco_articoli_head()
		
		ordine.mostra_giacenze=false
		ordine.mostra_spettanze=0

		call ordine.elenco_articoli()
		
		
		call ordine.elenco_totali()
        call ordine.riga_pulsanti(true,false,false)  
	       %>
        </table>

        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
          <tr>
            <td colspan="2" class="ui-widget-header">Stato: </td>
          </tr>
          <tr>
            <td width="20%" align="center" valign="bottom" ><input type="hidden" name="stato" value="" id="stato"/>
              <%=ordine.campo("data_stato1")%><br>
              <button type="submit" name="" class="pulsanti ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only<%=disabilitato%>" value="<%=1%>" <%if stato=1 then%> style="background:#66c466;" <%end if%> onClick="stato.value=<%=1%>; return;" >In modifica</button></td>
            <td width="20%" align="center" valign="bottom" class=""><%=ordine.campo("data_stato2")%><br>
              <button type="submit" name="" class="pulsanti ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only<%=disabilitato%>" value="<%=6%>" <%if stato=6 then%> style="background:#66c466;" <%end if%> onClick="stato.value=<%=6%>; return;" >Chiusa</button></td>
          </tr>
        </table>
        
        <%
	        
		call tabella_scadenze(0,idord)

	        
	        
	        %>
        
        
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
  
<script type="text/javascript" src="Jquery/js/jquery.qrcode.min.js"></script>
<script type="text/javascript" src="Jquery/js/idle.js"></script>
<script type="text/javascript" src="Jquery/js/autogrow.min.js"></script>
<script type="text/javascript" src="//cdn.jsdelivr.net/jquery.dirtyforms/2.0.0/jquery.dirtyforms.min.js"></script>
<script>
	
var documento="notacredito";
var idDocumento=$("#idord").val();
var nDocumento=$("#nord").val();
	
	
	
var tipo_documento='<%=tipo_documento%>';
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
var ddt_o_fattura="DF";	                    
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
	doAutogrow();
	puls_giu();
//	$("#form1").dirrty({
//	  preventLeaving: true,
//	  leavingMessage: 'message',
//	  onDirty: function(){
//		  
//		  alert("modificato");
//		  
//	  }
//	});
//	$("#form1").dirrty().on("dirty", function(){
//
//        $("#status").html("dirty");
//		  $("#Aggiorna_ordine").css('color', 'green');
//
//	});

$('#form1').dirtyForms();
$('#form1').on('dirty.dirtyforms clean.dirtyforms', function (ev) {
        if (ev.type === 'dirty') {
           $("#Aggiorna_ordine").css('color', 'green');
           $("#puls_aggiungi_articolo").attr("disabled","disabled");
        } else {
            $("#Aggiorna_ordine").css('color', '');
        }
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

document.title = "Nota di credito <%=nord%>";
</script>
<%end if%>
<%if oper="list" then%>
<script>
document.title = "Ordini";
var tabella="ordini";
var cosa="notacredito";
</script>
<%end if%>
</body>
</html>

<%


function libera_seriali (idord)
	conn.execute "update numeri_seriali set idord= null, iddett=null where idord="&idord,n
	if n>0 then
		libera_seriali="Liberati "&n&" numeri seriali"
	else
		libera_seriali=""
	end if
		
end function
call CheckConnChiusa()
%>