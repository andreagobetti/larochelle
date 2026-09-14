<%
'Verifica chiusure 04_12_2015
%>
<%t_inizio=timer()%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/chk_piva_cf.asp" -->
<!--#include virtual="/paginazione.asp" -->
<!--#include virtual="/parser.asp" -->
<!--#include file="ClasseOrdine.asp"-->
<!--#include file="ClasseModificheRS.asp"-->
<!--#include file="ClasseUtente.asp"-->
<!--#include file="ClasseFattura.asp"-->
<!--#include file="jsonObject.class.asp"-->
<%
if session("idadmin") = "" then call login()
cercain=request("cercain")
cerca=trim(request("cerca"))
oper=lcase(request("oper"))
iduser=request.form("iduser")
call via_senza_permesso("E2")
if request("filtra")<>"" then
	filtro=true
else
	filtro=false
end if
'Misure
if request.querystring("aggiungi_misura")<>"" then
	call aggiungi_misura("add")
	call connclose()
	response.end
end if
if request.querystring("modifica_misura")<>"" then
	call aggiungi_misura("edit")
	call connclose()
	response.end
end if
if request.querystring("tabella_misure")<>"" then
	call tabella_misure(request.querystring("tabella_misure"),false)
	call connclose()
	response.end
end if
'Gradi
dialog=request.querystring("dialog")
select case dialog
	case "modifica_grado"
		call dialog_grado("edit")
		call connclose()
		response.end
	case "aggiungi_grado"
		call dialog_grado("add")
		call connclose()
		response.end
	case "aggiungi_articolo_spettanze"
		call dialog_articolo_spettanze("add")
		call connclose()
		response.end
	case "modifica_articolo_spettanze"
		call dialog_articolo_spettanze("edit")
		call connclose()
		response.end
end select
'Fine misure
if request.form("nuovo_utente")<>"" then
	set utente= (new Classeutente)("add_da_dialog")
	iduser=utente.get_iduser()

end if
if request.form("crea_ordine")<>"" then
	if iduser="" then
		iduser=request.form("iduser")
	end if
	idord=0
	if request.form("carrello_utente")="" then
		idusercarrello=session("iduser")
	else
		idusercarrello=request("iduser")
	end if






	set ordine= (new ClasseOrdine)(array("nuovo","ordini",iduser,idusercarrello))
	idord=ordine.idord()
	if request.form("idpreventivo")<>"" then
		call ordine.copia_articoli_spettanze(request.form("idpreventivo"), "preventivi")
		call ordine.CalcolaTotaleMerce(false)
	end if



	set ordine = Nothing
	call connclose()

	response.redirect "pag_adm_ordini.asp?idord="&idord
end if

if request.querystring("trasformainordine")<>"" then

	call add2log("trasforma",0)

	idfoglio=request.querystring("trasformainordine")


	set ordine= (new ClasseOrdine)(array("foglioinordine","ordini",idfoglio))
	idord=ordine.idord()




	set ordine = Nothing
	call connclose()

	response.redirect "pag_adm_ordini.asp?idord="&idord
end if





if request.form("crea_sostituzione")<>"" then
	if iduser="" then
		iduser=request.form("iduser")
	end if
	idord=0
	if request.form("carrello_utente")="" then
		idusercarrello=session("iduser")
	else
		idusercarrello=request("iduser")
	end if
	set ordine= (new ClasseOrdine)(array("nuovo","sostituzione",iduser,0))
	idord=ordine.idord()
	if request.form("seleziona_ordine")<>"" then
		conn.execute("insert into ordini_sostituzioni (idord,idsostituzione) values ("&request.form("seleziona_ordine")&","&idord&")")
	end if
	set ordine = Nothing
	call connclose()

	response.redirect "pag_adm_sostituzioni.asp?idord="&idord
end if
if request.form("crea_fattura")<>"" then
	if iduser="" then
		iduser=request.form("iduser")
	end if
	'set objdocumento= (new ClasseFattura)(array("crea","tipo_fattura:ordine,ddt,fattura","tipo:accompagnatoria","iddocumento","iduser"))
	set objdocumento= (new ClasseFattura)(array("crea","fattura","","",iduser))
	idfat=objdocumento.idfat()
	set objdocumento = nothing
	'idfat=crea_fattura_fattura(iduser)
	call connclose()
	response.redirect "pag_adm_fatture.asp?idfat="&idfat
end if



if request.querystring("crea_sostituzione2")<>"" then
	idord=request.querystring("idord")
	set ordine= (new ClasseOrdine)(array("crea_sostituzione",idord))
	idord=ordine.idord()
	set ordine = Nothing
	call connclose()

	response.redirect "pag_adm_sostituzioni.asp?idord="&idord
end if

if request.form("crea_notacredito")<>"" then
	if iduser="" then
		iduser=request.form("iduser")
	end if
	idord=0
	idusercarrello=0
	set ordine= (new ClasseOrdine)(array("nuovo","notacredito",iduser,idusercarrello))
	idord=ordine.idord()
	set ordine = Nothing
	call connclose()

	response.redirect "pag_adm_note_credito.asp?idord="&idord
end if


if request.form("crea_preventivo")<>"" then
	if iduser="" then
		iduser=request.form("iduser")
	end if
	idord=0
	if lcase(request.form("crea_preventivo"))="crea preventivo" then
		if request.form("carrello_utente")="" then
			idusercarrello=session("iduser")
		else
			idusercarrello=request("iduser")
		end if
		set ordine= (new ClasseOrdine)(array("nuovo","preventivi",iduser,idusercarrello))


	else 'Duplica
		set ordine= (new ClasseOrdine)(array("duplica","preventivi",iduser,request.form("duplicaidord")))


	end if
	idord=ordine.idord ()
	set ordine = Nothing

	call connclose()
	response.redirect "pag_adm_preventivi.asp?idord="&idord
end if
if request.form("crea_ddt")<>"" then
		set ordine= (new ClasseOrdine)(array("nuovoddt",request.form("iduser"),""))

		idddt_creato=ordine.idord()
		nddt_creato=ordine.nord()

		'idddt_creato=crea_ddt()

		if idddt_creato<>"-1" then
			if request.form("spettanze")<>"" then

				call add2log("aggiorno spettanze",0)
				m_tabella="ordini"
				set ordine= (new ClasseOrdine)(array("aggiorna_quantita_spettanze",m_tabella,idord,idddt_creato))
				call add2log("passa di qui",0)
				call ordine.calcolo_totale_merce_ddt(idddt_creato)
			end if
		end if

	response.redirect "pag_adm_ddt.asp?idddt="&idddt_creato
end if

if request.form("crea_ordine_fornitore")<>"" then
	'if request.form("cerca_iduser")<>"" then
	if request.form("crea_ordine_fornitore")="Duplica ordine fornitore" then
		set ordine= (new ClasseOrdine)(array("duplica","ordini_fornitori",iduser,request.form("duplicaidord")))
	elseif true then
		add2log "Crea ordine a fornitore singolo",1

			iduser=request.form("iduser")
			idusercarrello=session("iduser")
			set ordine= (new ClasseOrdine)(array("nuovo","ordini_fornitori",iduser,idusercarrello))

		'idord=crea_ordine_preventivo("ordine_fornitore",0,session("idadmin"),request.form("cerca_iduser"))
		oper="edit"
	else
		add2log "Crea ordine a fornitore da carrello",1
		sql_for="SELECT carrello.iduser, prodotti.idfor FROM carrello INNER JOIN prodotti ON carrello.idpro = prodotti.IDpro WHERE ((visibilita=0) AND ((prodotti.vendita)=1)) and prodotti.idfor>0 and carrello.iduser="&session("iduser")&" GROUP BY carrello.iduser, prodotti.idfor; "
		set rs_fornitori=conn.execute(sql_for)

		do until rs_fornitori.eof
			iduser=rs_fornitori("idfor")
			idusercarrello=session("idadmin")
			'idord=crea_ordine_preventivo("ordine_fornitore",0,session("idadmin"),rs_fornitori("idfor"))
			set ordine = (new ClasseOrdine)(array("nuovo","ordini_fornitori",iduser,idusercarrello))
			set ordine = nothing

			rs_fornitori.movenext
		loop
		rs_fornitori.close
		set rs_fornitori=Nothing
		sql="DELETE carrello.iduser FROM carrello WHERE carrello.iduser=" & session("idadmin")
		conn.execute (sql)
		session("nelcarrello")=""
		oper="view"
	end if
	idord=ordine.idord()
	set ordine = Nothing
	call connclose()
	response.redirect "pag_adm_ordini_fornitori.asp?idord="&idord



end if

if request.form("iduser_primario")<>"" then
	iduser_primario=request.form("iduser_primario")
	idintestazione=conn.execute("select idintestazione from utenti where iduser="&iduser_primario)(0)



	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "select * from utenti where iduser="&iduser,conn,3,3
	rs("secondario")=iduser_primario
	rs("idintestazione")=idintestazione
	rs.update
	rs.Close
	set rs = Nothing
	oper="view"
	call add2log("reso secondario",0)





end if







if oper="update_accessori" then

	iduser=request.querystring("iduser")
	set JSON_accessori = New JSONobject
	JSON_accessori.Add "velcrocamicia", zerosenull(request.form("velcrocamicia"))
	JSON_accessori.Add "velcrogiacca", zerosenull(request.form("velcrogiacca"))
	JSON_accessori.Add "velcrogiubbino", zerosenull(request.form("velcrogiubbino"))
	if request.form("cintura")=1 then
		cintura=1
		elseif request.form("cintura")=0 then
		cintura=0
		else
			cintura=" "
			end if

	JSON_accessori.Add "cintura", cintura
	JSON_accessori.Add "camiciam", zerosenull(request.form("camiciam"))
	JSON_accessori.Add "camiciac", zerosenull(request.form("camiciaC"))
	JSON_accessori.Add "camiciottob", zerosenull(request.form("camiciottob"))
	JSON_accessori.Add "camiciottop", zerosenull(request.form("camiciottop"))
	val=request.form("interasse")
	if val="" then val=null
	JSON_accessori.Add "interasse", val
	val=request.form("note_accessori")
	if val="" then val=null
	JSON_accessori.Add "note", val

	Set rs_cliente = Server.CreateObject("ADODB.Recordset")
	rs_cliente.Open "select utenti_clienti.* from utenti_clienti where iduser="&iduser, conn, 1, 3
	if rs_cliente.eof then
		rs_cliente.addnew
		rs_cliente("iduser")=iduser
	end if
	rs_cliente("mod_dip")=JSON_accessori.serialize()
	rs_cliente.update
	rs_cliente.Close
	set rs_cliente= Nothing
	set JSON_accessori = nothing
	dati_salvati=true

end if
dim in_tab
Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class


if request.querystring("tab")<>"" then
	in_tab=true
	call carica_tab(request.querystring("tab"),request.querystring("iduser"))
	response.end
end if
if request.form("modifica")<>"" then
	iduser=request.form("modifica")
end if
if request.form("elimina_utente")<>"" then

	set utente= new Classeutente
	if utente.elimina_utente(iduser) ="" then
		oper="list"
	else
		oper="view"
	end if
	set utente = Nothing
end if
if request("reset_pw")<>"" then
		iduser=request("reset_pw")
		set utente= (new Classeutente)(array("new_password",iduser,"123456"))
		add2log "Resettata password [utente="&iduser&"]"&get_denominazione(iduser)&"[/utente] eseguita da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]",2

		call connclose()
		response.end
end if
if request.form("aggiungi_nota")<>"" then
	Set rs = Server.CreateObject("ADODB.Recordset")
'aggiorno tabella ticket_messaggi
	sql="select * from note_su_utenti where idnota=0 "
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	rs.addnew
	rs("data_nota")=now()
	rs("iduser")=iduser
	rs("iduser2")=session("iduser")
	rs("nota")=request.form("nota")
	if request.form("tipo_contatto")<>"" then rs("tipo_contatto")=cint(request.form("tipo_contatto"))
	rs("idagente")=request.form("idagente")
	rs.update
	rs.movefirst
	rs.close
	Set rs = Nothing
	select case oper
	case "add_new"
		'call invio_mail ("add_new",id_messaggio_ticket)
	case "add_reply"
		'call invio_mail ("add_reply",id_messaggio_ticket)
	end select
	oper="view"
end if
'response.write "step1"&oper
if request.form("annulla")<>"" then oper="list"
select case oper
	case "list"
		'oper="list"
	case "new"
		oper="new"
	case "add"
		oper="add"
	case "update"
		oper="update"
	case "view"
		oper="view"
	case else
		oper="list"
		if request.form("modifica")<>"" then oper="view"
		if request.form("edit")<>"" then oper="edit"
		if request.querystring("iduser")<>"" then
			oper="view"
			iduser=request("iduser")
		end if
		if isnumeric(request.form("cerca")) and request.form("cerca")<>"" and request.form("cercain")="" then
			oper="view"
			iduser=request.form("cerca")
		end if
		if request.form("cerca_hidden")<>"" then
			oper="view"
			iduser=request.form("cerca_hidden")
		end if

end select


if oper="add" or oper="update" then
	'controlli
	'CONTROLLO CF, PIVA, EMAIL
	if request.form("piva")<>"" or request.form("cf")<>"" or request.form("email")<>"" then
			fine_controllo=""
			if oper="update" then fine_controllo=" and utenti_intestazioni.iduser<>"& iduser
			if request.form("cf")<>"" and request.form("pivaduplicata")="" then
				controllo_cf=" cf='"& request.form("cf")&"'"
				sql_controllo="select * from utenti_intestazioni inner join utenti on utenti_intestazioni.id = utenti.idintestazione where "&controllo_cf&fine_controllo
				set rs=conn.execute(sql_controllo)
				if not rs.eof then error="CF già utilizzato da utente ID:"&rs("iduser") &" "&rs("nome")&" "&rs("cognome")
				rs.close
			end if
			if request.form("piva")<>"" and request.form("pivaduplicata")="" then
				controllo_piva=" piva='"& request.form("piva")&"'"
				sql_controllo="select * from utenti_intestazioni inner join utenti on utenti_intestazioni.id = utenti.idintestazione where "&controllo_piva&fine_controllo
				set rs=conn.execute(sql_controllo)
				if not rs.eof then
					if error<>"" then error=error&"\n"
					error=error&"PIVA già utilizzata da utente ID:"&rs("iduser") &" "&rs("nome")&" "&rs("cognome")
				end if
				rs.close
			end if
			if request.form("email")<>"" then
				controllo_email=" email='"& request.form("email")&"'"
				fine_controllo=replace(fine_controllo,"utenti_intestazioni.","")
				sql_controllo="select * from  utenti where "&controllo_email&fine_controllo
				'call add2log(sql_controllo,0)
				set rs=conn.execute(sql_controllo)
				if not rs.eof then
					if error<>"" then error=error&"\n"
					error=error&"Email già utilizzata da utente ID:"&rs("iduser") &" "&rs("nome")&" "&rs("cognome") &" "&rs("azienda")
				end if
				rs.close
			end if
			Set rs = Nothing
			if error<>"" then
				if oper="add" then oper="new"
				if oper="update" then oper="edit"
			end if
	end if	'FINE CONTROLLI
end if
tabella="utenti"
tabellaid="iduser"
'response.Write(oper)



if (oper="add" ) and error="" then
	set utente= (new Classeutente)("add")
	oper="list"
end if

if (oper="update" ) and error="" then
	set utente= (new Classeutente)(array("update",iduser))
	oper="list"
end if



%>
<!--#include virtual="/regioni_inc.asp" -->

<!--#include virtual="/sub_head_adm.asp" -->
<script type="text/javascript" src="chk_piva_cf.js"></script>
<script Language="JavaScript">
		function nome_cognome(){
		return true;
		}
function Valida_ricerca()
{
<%if false then %>

	if ((document.form1.cercain.value == 'nominativo' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il nominativo nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	if ((document.form1.cercain.value == 'id' )&&(document.form1.cerca.value == '')) {
		alert("Inserire id cliente nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	if ((document.form1.cercain.value == 'citta' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il nome citta nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	if ((document.form1.cercain.value == 'azienda' )&&(document.form1.cerca.value == '')) {
		alert("Inserire il nome azienda nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	if ((document.form1.cercain.value == 'email' )&&(document.form1.cerca.value == '')) {
		alert("Inserire l\'indirizzo email nella casella cerca.");
		document.form1.cerca.focus();
		return false;
	}
	<%end if %>
<%if oper="edit" then%>
//document.form1.oper.value='update';
<%else%>
//document.form1.oper.value='add';
<%end if%>
document.form1.submit();
}
</script>
  <style>
label {
	float: left;
	width: 150px;
	margin-bottom: 5px;
	font-weight: bold;

}
input[type="text"] {
	/*width: 500px;*/
	margin-bottom: 5px;
}
input[type="radio"] {
	width: 30px;
	margin-bottom: 5px;
}
textarea {
	width: 550px;
	height: 50px;
	margin-bottom: 5px;
}
.boxes {
	width: 1em;
}
.b_salva {
	margin-left: 120px;
	margin-top: 5px;
}
br {
	clear: left;
}
.sezione {
	background: #E5E5E5;
	width: 100%;
	margin-bottom: 5px;
	padding: 3px 0px 3px 3px;
}
</style>
<style>
.small-button {
   font-size: .9em !important;
}
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
.cl_selected{
	border-color: green;
	color: green;
	font-weight: bold;
}
</style>
<script>

$(function() {
<%if oper="view" then%>
$("#new_pw").click(function() {
	toastr.options = {"positionClass": "toast-bottom-right"};
	toastr.info("Invio email in corso");
	$.ajax({
		 url: "ajax_function.asp?oper=invia_pw&iduser=<%=iduser%>"  ,
		 type: "post",
		 success: function(data) {



		if (data.success==true)
		{
			//toastr.clear();
			toastr.options = {"positionClass": "toast-bottom-right"};
			toastr.success(data.Message);
		}
		else
		{
			toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
			toastr.error(data.Message);
		}
		menu.toggle();
		}
	});
});
$("#new_dip").click(function() {
	var tabs = $( "#multitabs" ).tabs();
	var ul = tabs.find( "ul" );
 		$( "<li><a href='<%=questofile&"?tab=16&oper=new&iduser="&iduser%>'>Dipendenti</a></li>" ).appendTo( ul );


		tabs.tabs( "refresh" );
		tabs.tabs('option', 'active', -1);
});
<%end if%>
<%if oper="list" then%>
		$( ".btn" ).button();

	$( "#cerca" ).autocomplete({
	<%if cercain="fornitori" then%>
		source: "ajax_function.asp?autocomplete=fornitori",
		<%else%>
		source: "ajax_function.asp?autocomplete=utenti",
		<%end if%>
		cache: false ,
		minLength: 2,
		select: function( event, ui ) {
			$('#cerca_hidden').val(ui.item.id);
			Valida_ricerca();
		}
	});
  <%end if%>
<%if oper="edit" or oper="new" then%>
		$( "#citta" ).autocomplete({
		source: "searcher_comuni.asp",
		cache: false ,
		minLength: 3,
		select: function( event, ui ) {
			$("#provincia").val(ui.item.targa);
			$("#regione").val(ui.item.regione);
			$("#cap").val(ui.item.cap);
		}
	});
	  <%end if%>
});
</script>
<script>
function Errore(){
<%if error<>"" then%>
	alert('<%=error%>');
<%end if%>
	}

</script>
<style>
.ui-menu {
	position: absolute;
	width: 200px;
	font-size:8px;
}
.box_errore{
	position: relative; left: 10px; top:-10; width:300px; z-index:100;
	display: none;
	}
	<%if cercain="fornitori" then
	grigiocl="style='color:#999;'"
	document_title="fornitori"
	else
	grigiofr="style='color:#999;'"
	document_title="Clienti"
	end if%>

	<%if oper="view" then %>
	/*notification-bubble  */
.notification-bubble {
	float:none;
	display:inline-block;
}

	<%end if %>


</style>
<link rel="stylesheet" type="text/css" href="Jquery/css/select2.css"/>
</head>
<body onLoad="Errore()">
<div id="wrap">
  <div id="header"> <%=titolo_top%>

    <!-- Box CORPO INIZIO-->
    <%barra=7%>
    <div id="barra_fissa"> <!--#include virtual="/sub_barra_adminsf2.asp" -->
      <div class="ui-widget-header ui-corner-all titolo_admin" id="barra_titolo"><a href="<%=questofile%>" <%=grigiocl%>>Clienti</a><%if ha_il_permesso("A5") then %> / <a href="<%=questofile%>?cercain=fornitori" <%=grigiofr%>>Fornitori</a><%end if %></div>
    </div>
    <form name="form1" id="form_user" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" >
    <!-- Colonna CORPO INIZIO-->
    <%
if oper="list" then
		sql="select utenti.*, utenti.nome, utenti.cognome, i.azienda, i.cf, i.piva, i.indirizzo, i.cap, i.citta, i.provincia, utenti_clienti.FatturaPACodiceDestinatario,email_inoltro_fattura,utenti_clienti.tipo, utenti_clienti.stato_estero FROM utenti LEFT JOIN utenti_clienti ON utenti.iduser = utenti_clienti.iduser left join utenti_intestazioni i on utenti.idintestazione = i.id "
		if request.querystring("iduser")<>"" then
			sql1=" utenti.iduser=" & request.querystring("iduser")&""
		end if
		if cerca<>"" and isnumeric(cerca) and cercain="" then cercain="id"
		iPageSize=Application("grec4page")
		ordine=" order by iduser desc"
		titolo_pagina="Clienti"
		ordinaper=request.cookies("user")("ordineclienti")
		ordinaper_form=request.form("ordina")
		if ordinaper_form<>"" then
			response.cookies("user").expires = date()+1000
			response.cookies("user")("ordineclienti")=ordinaper_form
			ordinaper=ordinaper_form
		end if
		if ordinaper="r" then

		elseif ordinaper="a" then
				ordine=" order by i.azienda, i.cognome, i.nome"
		end if
		if mod_larochelle then
			if cercain="" then
				cercain="clienti"
			end if


		end if
		'-----------sezione ricerca inizio
		if cercain<>"" or request.form("cerca_hidden")<>"" then
					cerca=replace(cerca,"'","''")
			select case cercain
			case "novisite"
				sql1=" numvisit=0"
			case "id"
				sql1=" utenti.iduser=" &cerca
			case "intestazione"
				sql1=" InStr(nome,'" & cerca& "')>0 or InStr(cognome,'" & cerca& "')>0 or InStr(azienda,'" & cerca& "')>0"
			case "nominativo2"
				sql1=" nome&' '&cognome='"&cerca&"'"
			case "admin"
				sql1=" admin=true"
			case "piva"
				sql1=" piva like '%"&cerca&"%'"
			case "fornitori"
				call via_senza_permesso("A5")
				sql1=" fornitore=1 "
				ordine=" order by azienda"
			case "clienti"
				 sql1=" fornitore=0 and (modalita_registrazione=1 or modalita_registrazione=4) "
			case "web"
				 sql1=" fornitore=0 and (modalita_registrazione=2 or modalita_registrazione=3) "
			case "carrello"
				 sql=sql&" inner join (select distinctrow iduser from carrello) c on utenti.iduser = c.iduser "
			case "preferiti"
				 sql=sql&" inner join (select distinctrow iduser from preferiti) c on utenti.iduser = c.iduser "
			case "dipendenti"
				 sql=sql&" inner join (select distinctrow iduser from utenti_dipendenti) d on utenti.iduser = d.iduser "
			case "spettanze"
				 sql=sql&" inner join (select distinctrow iduser from utenti_spettanze) d on utenti.iduser = d.iduser "
			case "secondario"
				 sql1=" secondario>0 "




			case else
				if cercain<>"" then	sql1=" InStr(" & cercain &",'" & cerca& "')>0"
				if request.form("cerca_hidden")<>"" then sql1=" iduser="& request.form("cerca_hidden")
			end select

			if session("idagente")<>"" and not ha_il_permesso("Z1") then
				'Determino filtro agente
				filtro_agente=split(session("permessi_agente"),"|")
				where_agente=""
				if filtro_agente(0)<>"" then
					call concatena_stringa( where_agente," and "," provincia='"&filtro_agente(0)&"'")
				end if
				if filtro_agente(1)<>0 then
					if filtro_agente(1)=1 then
						call concatena_stringa( where_agente," and "," trattamento_iva=10")
					elseif filtro_agente(1)=2 then
						call concatena_stringa( where_agente," and "," trattamento_iva<>10")
					end if
				end if
				if where_agente<>"" then where_agente=where_agente&" and "
				if sql1<>"" then
					sql1="idagente=0 or idagente="&session("idagente")&" or ("&sql1&")"
				else
					sql1="idagente=0 or idagente="&session("idagente")
				end if



				if sql1<>"" then
					sql1=where_agente&"  ("&sql1&")"
				else
					sql1=where_agente
				end if
			end if

			if len(sql1)>0 then
				sql1=" where " & sql1
			end if
			PaginazioneString="&cercain="&cercain&"&cerca="&cerca&"&ordinaper="&ordinaper
		end if
		if filtro then
			sql= "select DISTINCT utenti.iduser, utenti.Cognome, utenti.Nome, i.azienda, i.Citta, i.Indirizzo, i.Provincia, utenti.Telefono, i.regione, utenti.numvisit, utenti.mailing, utenti.email, utenti.Lastvisit, utenti.sito_suggerito, utenti.password_inviata, utenti_clienti.tipo, utenti.tipologia FROM (ordini RIGHT JOIN (utenti LEFT JOIN fatture ON utenti.iduser = fatture.iduser) ON ordini.iduser = utenti.iduser) LEFT JOIN utenti_clienti ON utenti.iduser = utenti_clienti.iduser left join utenti_intestazioni i on utenti.idintestazione = i.id "
			sql1=""
			ordine=""
			where=""
			if request("iva")="1" then
				where=" (utenti.trattamento_iva=2 or utenti.trattamento_iva=10)"
			elseif request("iva")="2" then
				where=" utenti.trattamento_iva<>2 and utenti.trattamento_iva<>10"
			end if
			if request("cli_for")<>"" then
				if where<>"" then where=where &" and "
				if request("cli_for")="C" then
					where=where&" utenti.fornitore=false"
				else
					where=where&" utenti.fornitore=true"
				end if
			end if
			if request("agente")<>"" then
				if request("agente")="-1" then
					if where<>"" then where=where &" and "
					where=where&" (utenti.idagente=0 or utenti.idagente is null)"


				else
					if where<>"" then where=where &" and "
					where=where&" utenti.idagente="&request("agente")
				end if
			end if
			if request("tipo")<>"" then
				if where<>"" then where=where &" and "
				if request("tipo")="N" then
					where=where&" utenti_clienti.tipo='' or utenti_clienti.tipo is null"
				else
					where=where&" utenti_clienti.tipo='"&request("tipo")&"'"
				end if
			end if
			if request("tipologia")<>"" then
				if where<>"" then where=where &" and "
				where=where&" instr(tipologia,'"&request("tipologia")&"')>0"
			end if
			if request("provincia")<>"" then
				if where<>"" then where=where &" and "
				where=where&" i.provincia='"&request("provincia")&"'"
			end if
			if request("regione")<>"" then
				if where<>"" then where=where &" and "
				where=where&" i.regione="&request("regione")
			end if
			if request("ordini")="1" then
				if where<>"" then where=where &" and "
				where=where&" (ordini.idord) Is Not Null"
			elseif request("ordini")="2" then
				if where<>"" then where=where &" and "
				where=where&" (ordini.idord) Is Null"
			end if
			if request("fatture")="1" then
				if where<>"" then where=where &" and "
				where=where&" (fatture.idfat) Is Not Null"
			elseif request("fatture")="2" then
				where=where&" (fatture.idfat) Is Null"
			end if
			if request("anno")<>"" and request("fatture")<>"" then
				if where<>"" then where=where &" and "
				where=where&" year(fatture.data) ="&request("anno")
			end if
			if request("anno")<>"" and request("ordini")<>"" then
				if where<>"" then where=where &" and "
				where=where&" year(ordini.data) ="&request("anno")
			end if


			if where<>"" then sql=sql&" WHERE "&where
				sql=sql&" order by azienda, cognome, nome"

			PaginazioneString=replace("&"&request.querystring&request.form,"page="&request("page"),"")

		end if

		'-----------sezione ricerca fine
		'response.write sql
		strSql = sql&sql1 &ordine
		'if utente_andrea then
			'response.write strSql
		'end if
		%>
    <%=response.write (errore)%>
    <table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1">
      <tr>
        <td colspan="5">Cerca
          <input type="text" name="cerca" id="cerca" class="casella_ricerca" value="<%=cerca%>" >
          <input type="hidden" id="cerca_hidden" name="cerca_hidden">
          in
          <select name="cercain" id="cercain" style="FONT: 12px;" onChange="Valida_ricerca()">
            <option value="" <%if cercain="" then response.write "selected"%>>Visualizza tutti</option>
            <%if ha_il_permesso("A5") then %>
            <option value="fornitori" <%if cercain="fornitori" then response.write selected%>>fornitori</option>
            <%end if %>
            <option value="intestazione" <%if cercain="intestazione" then response.write selected%>>Nominativo o azienda</option>
            <option value="nominativo2" <%if cercain="nominativo2" then response.write selected%>>Nominativo esatto</option>
            <option value="piva" <%if cercain="piva" then response.write selected%>>Partita Iva</option>
            <option value="id" <%if cercain="id" then response.write selected%>>ID cliente</option>
            <option value="azienda" <%if cercain="azienda" then response.write selected%>>Azienda</option>
            <option value="citta" <%if cercain="citta" then response.write selected%>>Città</option>
            <option value="email" <%if cercain="email" then response.write selected%>>Email</option>
            <option value="novisite" <%if cercain="novisite" then response.write selected%>>Nessuna visita</option>
            <option value="web" <%if cercain="web" then response.write selected%>>Utenti web</option>
            <option value="carrello" <%if cercain="carrello" then response.write selected%>>Con articoli nel carrello</option>
            <option value="preferiti" <%if cercain="preferiti" then response.write selected%>>Con articoli preferiti</option>
            <option value="dipendenti" <%if cercain="dipendenti" then response.write selected%>>Con dipendenti</option>
            <option value="spettanze" <%if cercain="spettanze" then response.write selected%>>Con foglio spettanze</option>
            <option value="secondario" <%if cercain="secondario" then response.write selected%>>Account secondario</option>
          </select>
		  <select name="ordina" onChange="Valida_ricerca()">
			<option value="r" <%if ordinaper="r" then response.write "selected" %>>Ordina per: Registrazione</option>
			<option value="a" <%if ordinaper="a" then response.write "selected" %>>Ordina per: Alfabeticamente</option>
		</select>
          <input type="button" name="Submit" value="Cerca" style="FONT: 12px;" onClick="Valida_ricerca()" class="">
          <input type="button" id="filtri" value="+filtri" style="FONT: 12px;">
          <span style="float:right;">
          <%if ha_il_permesso("E1") then %>
          <input type="button"  value="Nuovo" style="FONT: 12px;" onclick="window.location.href='pag_adm_user.asp?oper=new'">
          <%end if %>
          </span></td>
      </tr>
	  <%if filtro then
			display=""
		else
			display="display: none;"
	  end if

	  %>
	  <tr style="border-bottom:1px solid; <%=display%>" id="tr_filtri">
	  <td  colspan="5">
	<%
	val=request("cli_for")
	%>
			<select name="cli_for" id="cli_for" <%=class_selected(val)%>>
				<option value="" <%if val="" then response.write " selected"%>>Clienti e fornitori: Tutti</option>
				<option value="C" <%if val="C" then response.write " selected"%>>Solo clienti</option>
				<option value="F" <%if val="F" then response.write " selected"%>>Solo fornitori</option>
			</select>
	<%
	val=request("tipo")
	%>
			<select name="tipo" id="tipo"  <%=class_selected(val)%>>
				<option value="" <%if val="" then response.write " selected"%>>Tipo: Tutti</option>
				<option value="S" <%if val="S" then response.write " selected"%>>Società</option>
				<option value="D" <%if val="D" then response.write " selected"%>>Ditta individuale</option>
				<option value="P" <%if val="P" then response.write " selected"%>>Privato</option>
				<option value="A" <%if val="A" then response.write " selected"%>>Associazione</option>
				<option value="N" <%if val="N" then response.write " selected"%>>Non specificato</option>
			</select>
	<%
	val=request("iva")
	%>
              <select name="iva" id="iva"  <%=class_selected(val)%>>
                              <option value="">IVA: Tutti</option>
                              <option value="1" <%if request("iva")="1" then%> selected <%end if%>>Iva Vs. carico ex art.17</option>
                              <option value="2" <%if request("iva")="2" then%> selected <%end if%>>Escluso Iva Vs. carico ex art.17</option>
              </select>
<%
			val=request("provincia")
%>
              <select name="provincia" id="provincia" <%=class_selected(val)%>>
                <option value="">Province: Tutti</option>
                <%
				set rs_comuni=conn.execute("select * from elenco_province order by denominazione_provincia")
				do while not rs_comuni.eof
					response.write "<option value='"&rs_comuni("sigla_automobilistica")&"'"
					if rs_comuni("sigla_automobilistica")=val then response.write " selected "
					response.write ">" & rs_comuni("denominazione_provincia")&"</option>"
					rs_comuni.movenext
				loop
				rs_comuni.close
				set rs_comuni=nothing

				%>
              </select>
			  <%
			  val=request("regione")
			  %>
			<select name="regione" id="regione" <%=class_selected(val)%>>
                <option value="">Regioni: Tutte</option>
                <%
				set rs_regioni=conn.execute ("select DISTINCT elenco_province.Codice_regione, elenco_province.Denominazione_regione FROM elenco_province order by Denominazione_regione")
				do while not rs_regioni.eof
					response.write "<option value='"&rs_regioni("Codice_regione")&"'"
					if rs_regioni("Codice_regione")=val then response.write " selected "
					response.write ">" & rs_regioni("Denominazione_regione")&"</option>"
					rs_regioni.movenext
				loop
				rs_regioni.close
				set rs_regioni=nothing
			%>
			</select>
              <%
				val=request("tipologia")
			  %>         <select name="tipologia"  id="tipologia" <%=class_selected(val)%>>
                <option value="">Tipologia: Tutte</option>
                <%
				set rs=conn.execute("select * from utenti_tipologia order by tipologia")
				do while not rs.eof
					response.write "<option value='"&rs("id")&"'"
					if cstr(rs("id"))=val then response.write " selected "
					response.write ">" & rs("tipologia")&"</option>"
					rs.movenext
				loop
				rs.close
%>
              </select>
			  <%

				val=request("agente")
			  %>         <select name="agente" id="agente" <%=class_selected(val)%>>
                <option value="">Agenti: Tutti</option>
                <option value="-1" <%if cstr(-1)=val then response.write " selected "%>>Senza agente</option>
                <%
				set rs=conn.execute("select * from agenti order by nominativo")
				do while not rs.eof
				response.write "<option value='"&rs("idagente")&"'"
				if cstr(rs("idagente"))=val then response.write " selected "
				response.write ">" & rs("nominativo")&"</option>"
				rs.movenext
				loop
				rs.close
				set rs = nothing
				%>
              </select>
			<%
			val=request("ordini")
			%>
			<select name="ordini" id="ordini" <%=class_selected(val)%>>
				<option value="">Ordini: Tutti</option>
				<option value="1" <%if val="1" then response.write " selected"%>>Con ordini</option>
				<option value="2" <%if val="2" then response.write " selected"%>>Senza ordini</option>
			</select>
			<%
			val=request("fatture")
			%>
			<select name="fatture" id="fatture" <%=class_selected(val)%>>
				<option value="">Fatture: Tutti</option>
				<option value="1" <%if val="1" then response.write " selected"%>>Con fatture</option>
				<option value="2" <%if val="2" then response.write " selected"%>>Senza fatture</option>
			</select>
              Anno (ordini o fatture)
              <input name="anno" id="anno" type="text" value="" size="4">


              <input type="submit" name="filtra" id="filtra" value="Applica filtro">

              </td>
          </tr>
      <tr>
        <td align="center" bgcolor="#E5E5E5"><p> Edit</p></td>
        <td bgcolor="#E5E5E5">Denominazione/Email</td>
        <td bgcolor="#E5E5E5">&nbsp;&nbsp;Nominativo/Indirizzo</td>
        <td align="center" bgcolor="#E5E5E5">Citt&agrave;/Regione&nbsp;</td>
        <td align="center" bgcolor="#E5E5E5">Visite/Ultima</td>
      </tr>
      <%
call paginazione_start(strsql,ipagesize,"access")
Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
%>
      <tr class="coprobox" id="iduser_<%=objPagingRS("iduser")%>">
        <td align="center" style="border-bottom:1px solid;">
        <a href="<%=questofile%>?iduser=<%=objPagingRS("iduser")%>" class="btn small-button">Modifica</a>
                <%if utente_andrea then response.write "<br>"&objPagingRS("iduser")%>

        </td>
        <td style="border-bottom:1px solid;">
          <b><%=denominazione(objPagingRS("nome"),objPagingRS("cognome"),objPagingRS("azienda"))%></b><br>
          <a href='mailto:<%=objPagingRS("email")%>'><%=objPagingRS("email")%></a>
          <%response.write suggerisci_sito(objPagingRS("sito_suggerito"))%>
          <br>
          <%if isnull(objPagingRS("password_inviata")) and objPagingRS("email")<>"" then%>
		  <span class="invia_pw_data">
          <a href="#" title="Clicca per inviare la password" style="color: red;" class="invia_pw">Password non inviata</a></span>
          <%else
			  response.Write(objPagingRS("password_inviata"))
              end if

			  %></td>
        <td style="border-bottom:1px solid;"><%=vedi_tipologia(objPagingRS("tipologia"))%><b><%=objPagingRS("cognome")&" "&objPagingRS("nome")%></b><br>
          <%=objPagingRS("Indirizzo")%></td>
        <td align="center" style="border-bottom:1px solid;"><%=objPagingRS("citta")%><%if objPagingRS("provincia")<>"" then response.write  " ("&objPagingRS("provincia")&")"%></td>
        <td align="center" class=corpobox style="border-bottom:1px solid;"><%=objPagingRS("numvisit") & "<br>" & objPagingRS("lastvisit")%></td>
      </tr>
      <%
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	call paginazione_end(5,true)
	%>
    </table>
    <%
	else
	'***********************************************************************************************************************************************
	'****************************SCHEDA CLIENTE
	'***********************************************************************************************************************************************
		  %>
    <%if oper="view" then disabled=" disabled"%>
    <%if oper="view" or oper="edit" then
				  if request.querystring("iduser")<>"" then iduser=request.querystring("iduser")
					sql="select utenti.* FROM utenti  where utenti.iduser=" & iduser
					'sql="select * from utenti where iduser=" & iduser
					set rs=conn.execute(sql)
					nominativov=get_denominazione(iduser)
					cancellabile=true
					Set rs1 = Server.CreateObject("ADODB.Recordset")
					sql="select count(iduser) FROM ordini where tipo_documento='ordine' and eliminato=0 and iduser="& iduser
					set rs1= conn.execute (sql)
					n_ordini=clng(rs1(0))
					if n_ordini>0 then cancellabile=false


					sql="select Count(fatture.IDfat) AS ConteggioDiIDfat, Count(scadenze.idScadenza) AS ConteggioDiidScadenza FROM fatture LEFT JOIN scadenze ON fatture.IDfat = scadenze.idfat WHERE iduser="& iduser

					'sql="select iduser FROM fatture where iduser="& iduser
					set rs1= conn.execute (sql)
					fatture=clng(rs1("ConteggioDiIDfat"))
					scadenze=clng(rs1("ConteggioDiidScadenza"))
					rs1.close
					if fatture>0 then cancellabile=false
					sql="select count(iduser) FROM ordini where tipo_documento='ddt' and ordini.iduser="& iduser
					set rs1= conn.execute (sql)
					ddt=clng(rs1(0))
					if ddt>0 then cancellabile=false

					sql="select count(iduser) FROM preventivi where eliminato=0 and iduser="& iduser
					set rs1= conn.execute (sql)
					n_preventivi=clng(rs1(0))
					if n_preventivi>0 then cancellabile=false
					rs1.close
					sql="select count(iduser) FROM carrello where iduser="& iduser
					set rs1= conn.execute (sql)
					n_carrello=clng(rs1(0))
					rs1.close
					sql="select count(*) from preferiti where iduser="& iduser
					n_preferiti=clng(conn.execute(sql)(0))

					sql="select count(*) from utenti_foglio_spettanze where iduser="& iduser
					fogli_spettanze=clng(conn.execute(sql)(0))

					if rs("idloged")<>0 then
						sql="select count(iduse) FROM log_action where iduse="& rs("iduser")
						set rs1= conn.execute (sql)
						n_log=clng(rs1(0))
					else
						n_log=0
					end if
					if rs("fornitore") then
						Set rs1 = Server.CreateObject("ADODB.Recordset")
						sql="select count(iduser) FROM ordini_fornitori where  iduser="& iduser
						set rs1= conn.execute (sql)
						n_ordini_for=clng(rs1(0))
						if n_ordini_for>0 then cancellabile=false
						rs1.close

						sql="select count(idpro) FROM prodotti where idfor="& iduser
						set rs1= conn.execute (sql)
						articoli_forniti=clng(rs1(0))
						if articoli_forniti>0 then cancellabile=false
						rs1.close

						sql="select count(*) FROM fatture_for where idfor="& iduser
						set rs1= conn.execute (sql)
						fatture_fornitore=clng(rs1(0))
						if fatture_fornitore>0 then cancellabile=false
						rs1.close

						sql="select count(scadenze_for.scadenza) FROM fatture_for INNER JOIN scadenze_for ON fatture_for.IDFat = scadenze_for.idfat WHERE ((scadenze_for.Pagato)=False) AND (fatture_for.idfor)="& iduser
						set rs1= conn.execute (sql)
						scadenze_fornitore=clng(rs1(0))
					end if
					fornitore=rs("fornitore")
					secondari=clng(conn.execute("select count(*) from utenti where secondario="&iduser)(0))



					n_dipendenti=0
					if mod_dip then
						if rs("secondario")>0 then
							iduserdipendenti=rs("secondario")
							else
								iduserdipendenti=iduser

						end if
						n_dipendenti=clng(conn.execute("select count(*) from utenti_dipendenti where iduser="&iduserdipendenti)(0))
					end if
					set rs1=Nothing
					set rs=Nothing
			  end if
			  %>

	<%if oper="view" then %>
	<div class="ui-widget-header ui-corner-all titolo_admin" style="font-size: larger;">
		<%=nominativov%>
	</div>

		<%
		if fornitore=false then
			titolo_tab="Dati cliente"
		else
			titolo_tab="Dati fornitore"
		end if
	%>
	<div id="multitabs" style="display:none;">
        <ul>
          <li><a href="#tabs-1" id="first_tab"><%=titolo_tab%></a></li>
          <%if n_ordini>0 then %>
          <li><a href="<%=questofile%>?tab=2&iduser=<%=iduser%>">Ordini (<%=n_ordini%>)</a></li>
          <li><a href="<%=questofile%>?tab=7&iduser=<%=iduser%>">Articoli ordinati</a></li>
          <% end if%>
          <%if n_ordini_for>0 and ha_il_permesso("A5") then %>
          <li><a href="<%=questofile%>?tab=15&iduser=<%=iduser%>">Ordini fornitore(<%=n_ordini_for%>)</a></li>
          <% end if%>
          <%if n_preventivi>0 then %>
          <li><a href="<%=questofile%>?tab=5&iduser=<%=iduser%>">Preventivi (<%=n_preventivi%>)</a></li>
          <% end if%>
          <%if fatture>0 and ha_il_permesso("B2") then %>
          <li><a href="<%=questofile%>?tab=3&iduser=<%=iduser%>">Fatture (<%=fatture%>)</a></li>
          <%end if %>
          <%if fatture_fornitore>0 and ha_il_permesso("A5") then %>
          <li><a href="<%=questofile%>?tab=12&iduser=<%=iduser%>">Fatture fornitore (<%=fatture_fornitore%>)</a></li>
          <%end if %>
          <%if scadenze>0 and ha_il_permesso("B2") then %>
          <li><a href="<%=questofile%>?tab=9&iduser=<%=iduser%>">Scadenze (<%=scadenze%>)</a></li>
          <% end if%>
          <%if scadenze_fornitore>0 and ha_il_permesso("A5") then %>
          <li><a href="<%=questofile%>?tab=13&iduser=<%=iduser%>">Scadenze fornitore (<%=scadenze_fornitore%>)</a></li>
          <% end if%>
          <%if ddt>0 and ha_il_permesso("B2") then %>
          <li><a href="<%=questofile%>?tab=4&iduser=<%=iduser%>">Ddt (<%=ddt%>)</a></li>
          <% end if%>
          <%if articoli_forniti>0 and ha_il_permesso("A5") then %>
          <li><a href="<%=questofile%>?tab=14&iduser=<%=iduser%>">Articoli forniti (<%=articoli_forniti%>)</a></li>
          <% end if%>
          <%if n_carrello>0 then
          carrellosi="&carrello=si"
          %>
          <li><a href="<%=questofile%>?tab=6&iduser=<%=iduser%>">Carrello (<%=n_carrello%>)</a></li>
          <% end if%>
          <%if n_preferiti>0 then
          carrellosi="&carrello=si"
          %>
          <li><a href="<%=questofile%>?tab=22&iduser=<%=iduser%>">Preferiti (<%=n_preferiti%>)</a></li>
          <% end if%>

          <li><a href="#tabs-10">Note</a></li>
          <%if n_log>0 then %>
          <li><a href="<%=questofile%>?tab=8&iduser=<%=iduser%>">Attività sul sito (<%=n_log%>)</a></li>
          <% end if%>
          <%if mod_dip then %>
	          <li><a href="<%=questofile%>?tab=16&oper=list&iduser=<%=iduser%>">Dipendenti (<%=n_dipendenti%>)</a></li>
	          <li><a href="<%=questofile%>?tab=18&oper=list&iduser=<%=iduser%>">Articoli per spettanze</a></li>
	          <%if fogli_spettanze>0 then %>
	          <li><a href="<%=questofile%>?tab=23&oper=list&iduser=<%=iduser%>">Fogli spettanze (<%=fogli_spettanze%>)</a></li>
	          <%end if %>

	          <li><a href="<%=questofile%>?tab=20&oper=list&iduser=<%=iduser%>">Accessori</a></li>
	          <%if false then %>
	          <li><a href="<%=questofile%>?tab=17&oper=list&iduser=<%=iduser%>">Gradi</a></li>
	          <%end if %>
          <% end if%>

          <li><a href="<%=questofile%>?tab=11&iduser=<%=iduser%>">Registro</a></li>
          <%if utente_andrea then %>
          <li><a href="<%=questofile%>?tab=21&iduser=<%=iduser%>">Intestazioni e consegna</a></li>
          <%end if %>
          <%if secondari>0 then %>
          <li><a href="<%=questofile%>?tab=24&iduser=<%=iduser%>">Utenti secondari</a></li>
          <%end if %>
        </ul>
        <div id="tabs-1" class="tab-panel">
		<%end if	'if oper="view"
		in_tab=false
		call carica_tab(1,iduser)


        if oper<>"view" then %>
    <script src="Jquery/js/select2.js"></script>
<br /><input type='hidden' name='idpro' id='idpro' style="width:400px;"/><input type='hidden' name='idvara' id='vara' style="width:200px;"/>
<input type='hidden' name='idvarb' id='varb' style="width:200px;"/>
  <script type="text/javascript">
var form;
var termine;
function aggiungi(){
	$.get("pag_adm_tipologia.asp?ajax=true&tipologia="+termine, function(result){
	  if (result!="")
		{
			$(form).append($('<option>', {value:result, text: termine}));

			 $(form).trigger("change");
			 var data=$(form).select2("data");
			 data.push({"id":result,"text":termine});
			 $(form).select2("data", data);
		}
		else
		{
//				toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
//				toastr.error('Errore in aggiornamento prenotazione ordine');
		}
	});
	return true;
}
$(document).ready(function () {
	form=  $('#tipologia');
	$(form).select2({
	formatNoMatches:function(term){
			termine=term;
            return "Voce non in elenco, <button onclick='return aggiungi();'>aggiungi "+term+"  </button>";
		}}
	);
});
checkAnyFormFieldEdited();
function enableSaveBtn(t_this){
	//alert("enableSaveBtn");
		$("#salva").removeAttr("disabled");
	}
</script>
    <%end if %>





	<%if oper="view" then %>
	</DIV><!-- End .tabs-1 -->
	<script src="Jquery/js/select2.js"></script>

    </form>
            <div id="tabs-10" class="tab-panel">
    <%
	Set rs = Nothing

	'Blocco visualizzazione TICKET SINGOLO
	if false then
	if oper="view" then
			sql="select note_su_utenti.*, agenti.Nominativo, utenti.Cognome, utenti.Nome FROM (note_su_utenti LEFT JOIN agenti ON note_su_utenti.idagente = agenti.IDagente) INNER JOIN utenti ON note_su_utenti.iduser2 = utenti.iduser  where note_su_utenti.iduser=" & iduser
		  set rs=conn.execute(sql)
	if not rs.eof then
	%>
		<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
		  <tr>
			<td class="ui-widget-header">Note</td>
		  </tr>
		  <%
			do until rs.EOF%>
		  <tr align="left">
			<td style="border-bottom:1px solid;">Messaggio di: <b><%=rs("nome")&" "&rs("cognome")%></b> il <b><%=rs("data_nota")%></b> contatto <%=tipo_contatto(rs("tipo_contatto"))%> agente <%=rs("nominativo")%><br>
			  <br>
			<%=rs("nota")%></td>
		  </tr>
		  <%
		rs.movenext
		loop
		%>
		</table>
		<%end if
				Set rs = Nothing

	'Blocco immissione RISPOSTE

	if oper="view"  then
	oper="new"
	if request.form("id")<>"" then txt=request.form("id")
		sql=sql & " where id_attrezzatura =" & txt
		select case oper
		case "new"
			titolo= "Aggiungi nota"
		case "view_single"
			titolo= "Aggiungi messaggio"
			oper="reply"
		end select
		%>
		<script Language="JavaScript">
	pulsanteannulla=false;
	function valida_note(theForm)
	{
	  if (pulsanteannulla == true)  {
		return(true);
	  }
		if (tinyMCE.get('nota').getContent() == "")  {
		alert("Il campo messaggio è obbligatorio.");
		 tinyMCE.execCommand('mceFocus', false, "nota");
		return (false);
	  }
	  return (true);
	}
	</script>
		<form name="form2" method="post" action="<%=Request.ServerVariables("Script_Name")%>"  onSubmit="return valida_note(this)" style="margin-top:0px;">
		  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
			<tr>
			  <td colspan="2" class="ui-widget-header"><%response.write titolo%></td>
			</tr>
			<tr>
			  <td align="right" valign="top">Tipo contatto</td>
			  <td><input type="hidden" name="iduser" value="<%=iduser%>"/><select name="tipo_contatto">
				  <option value="">Selezionare</option>
				  <%
	for n=1 to tipo_contatto(0)
	response.write "<option value='"&n&"'>" & tipo_contatto(n)&"</option>"
	next
	%>
				</select>
				<select name="idagente" >
				  <option value="0">Nessuno</option>
				  <%
	set rsm=conn.execute("select * from agenti order by nominativo ;")
	do while not rsm.eof
	response.write "<option value=""" & rsm("idagente") &""""
	'if rs("idagente")=rsm("idagente") then response.write " selected"
	response.write ">" & rsm("nominativo")& "</option>"
	rsm.movenext
	loop
	rsm.close
	set rsm = Nothing
	%>
				</select></td>
			</tr>
			<tr>
			  <td align="right" valign="top">Messaggio</td>
			  <td><%if error<>"" then
						testo=request.form("messaggio_ticket")
					elseif oper="edit" then
						testo=rs("messaggio_ticket")
					elseif oper="new" then
						testo=""
					end if%>
				<!--#include file="tinymce4_conf2.asp"-->
				<textarea name="nota" rows="15"  class="mceEditor" id="nota" style="width: 500 px"  ><%=testo%></textarea>
			</tr>

			<tr>
			  <td colspan="2" align="center" valign="top"><input type="hidden" name="oper" value="add_<%=oper%>">
				<input type="submit" name="aggiungi_nota" value="<%=titolo%>">
				</td>
			</tr>
		  </table>
		</div>
		  <%
			  end if

		end if
		end if

	end if
	call connclose()
end if
%>
    </form>

    <!-- Box CORPO FINE-->
  </div>
  <%
	txt_timer=txt_timer&"T fine:"&FormatNumber(timer() - t_inizio, 2)&" "
%>
  <div id="footer"><%=txt_timer%></div>
<!--#include virtual="/pag_adm_footer_inc.asp" -->
</div>
<!-- OPER: <%=oper%>   -->
   <%if oper="view" then%>


<%end if%>
 <%if oper="list" then%>
<script>
document.title = "<%=document_title%>";
$(function(){
	var stringa=$("#cerca").val();
	console.log(window.location);
	$("#cerca").on("input", function() {
	if(parseInt($("#cerca").val(), 10) > 0) {
		$("#cercain").val("id");
	}

	if ($("#cerca").val().indexOf("@")>-1)
	{
		$("#cercain").val("email");
	}
	});
	$("#filtri").click(function(e){
		$("#tr_filtri").toggle();
	});
	$(".invia_pw").click(function(e){
		e.preventDefault();
		toastr.options = {"positionClass": "toast-bottom-right"};
		toastr.info("Invio email in corso");
		var id_user=$(this).closest("tr").attr("id").substring(7);
		var span=$(this).closest(".invia_pw_data");
		//var id_user=$(this).closest("tr").attr("id");
		console.log("invia_pw iduser:"+id_user);


		$.get("ajax_function.asp?oper=invia_pw&iduser=" + id_user, function (result) {
			//alert("qui");
			if (result.success==true)
			{
				//toastr.clear();
				toastr.options = {"positionClass": "toast-bottom-right"};
				toastr.success(result.Message);
				span.html(result.data_invio);
			}
			else
			{
				toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
				toastr.error(result.Message);
				span.html(result.Message);
			}
		});




	});


});
</script>
<%else %>
<script>
document.title = "<%=document_title%>";
<%if iduser<>"" then %>
var iduser=<%=iduser%>;
<%end if %>
				function aggiorna_tabella_dipendenti(){
					if($('#elenco_dipendenti').length){
						console.log("#elenco_dipendenti esiste");
					var tabs_panel=$(this).closest(".ui-tabs-panel");
					iduser=$("#iduser").val();
					var url="pag_adm_user.asp?tab=16&iduser=<%=iduser%>&oper=list"
					console.log('aggiorno tabella dipendenti iduser:'+iduser+" url:"+url);
					$.ajax({
						url     : url,
						type    : "post",
						cache	: false,
						//dataType: 'json',
						//data	: dati,
						success: function(data){
							var iddivparent=$("#elenco_dipendenti").closest(".ui-tabs-panel").attr("id");

							$("#"+iddivparent).html(data);
						}
						,error: function(xhr, textStatus, error){
							toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
							toastr.error('Errore nel caricamento della pagina');
							var txt="";
							txt+="Errore in "+location.href+"<br>";
							txt+="<br>querystring: ?"+querystring[1];
							txt+="<br>"+xhr.statusText;
							txt+="<br>"+xhr.responseText;
							txt+="<br>textStatus:"+textStatus;
							txt+="<br>error:"+error;

							$.ajax({
								url     : "searcher.asp",
								type    : "post",
								data	: "txt_errore="+encodeURIComponent(txt)
							});
						}
					});
					}



				}

$(function(){

	$("#reset_pw").click(function() {
	$.ajax({
		 url: "<%=questofile%>?reset_pw=<%=iduser%>"  ,
		 type: "post",
		 success: function(data) {
		 toastr.options = {
				"positionClass": "toast-bottom-right"
			};
			toastr.success( "Password resettata 123456");
			menu.toggle();
				}
			});
});


});



</script>
<%end if%>
</body>
</html>
<%
function suggerisci_sito(sito)
if sito<>"" then
suggerisci_sito="<br><a href='http:\\"&sito&"' target='_blank'><b>Sito suggerito</b></a>"
else
suggerisci_sito=""
end if
end function
%>
<%
function cerca_sito(email)
chiocciola=instr(1,email,"@")
cerca_sito=mid(email,chiocciola+1)
if instr(1,Application("escludi_sito_suggerito"),cerca_sito)=0 then
cerca_sito="www."&cerca_sito
else
cerca_sito=""
end if
end function
%>
<%
function tipo_contatto(stato)
	if not isnumeric(stato) then
		tipo_contatto=stato
	else
		select case stato
		case 0
		tipo_contatto=3
		case 1
		tipo_contatto="NOSTRA VISITA"
		case 2
		tipo_contatto="TELEFONATA"
		case 3
		tipo_contatto="ORDINE"
		end select
	end if
end function
sub carica_tab(tab,iduser)
	'add2log "Richiesto tab:"&tab,-1
	PaginazioneString="&tab="&tab&"&iduser="&iduser
	if oper="" then oper="view"
	select case tab
	case 1








		sql="select utenti.*, utenti_intestazioni.azienda, utenti.nome, utenti.cognome, utenti_intestazioni.cf, utenti_intestazioni.piva, utenti_intestazioni.indirizzo,utenti_intestazioni.cap, utenti_intestazioni.citta, utenti_intestazioni.provincia, utenti_intestazioni.regione, utenti_clienti.FatturaPACodiceDestinatario,email_inoltro_fattura,utenti_clienti.tipo, utenti_clienti.stato_estero FROM utenti LEFT JOIN utenti_clienti ON utenti.iduser = utenti_clienti.iduser left join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id "

		if oper="new" then
	    	'Apro recordset per acquisire la lunghezza dei campi

		else
			sql=sql&" where utenti.iduser=" & iduser
		end if
		'response.write sql
		set rs = Server.CreateObject("ADODB.Recordset")
	  	rs.Open sql, conn, 1, 3
	  	piva_editabile="SI"
		come_tipo="Default"
		'pres_tipo="S"
		secondario=0
		if oper<>"new" then
			'Determino tipo cliente
			pres_tipo=rs("tipo")
			secondario=clng(rs("secondario"))
			if utente_andrea then response.write "[soloio]oper:"&oper&" pres_tipo:"&pres_tipo&"[/soloio]"
			if valore_vuoto(pres_tipo) then
				'Se non ancora impostato determino possibile tipo cliente
				if utente_andrea then response.write "piva:"&valore_vuoto(rs("piva"))&" azienda:"&  valore_vuoto(rs("azienda"))
				pres_tipo=determina_tipo_cliente(rs("azienda"),rs("cf"),rs("piva"))
				testo_tipo="<br><span class=""rosso"">Tipo cliente non impostato ma determinato in automatico in base ai dati anagrafici</span>"
			end if

		end if
	%>



	<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
      <tr>
        <td colspan="2" valign="middle" class="ui-widget-header"><%
					if oper="view" then

						if rs("fornitore")=0 then
							if secondario>0 then
								response.write "Scheda account secondario di <a href=""pag_adm_user.asp?iduser="&rs("secondario")&""""">"&rs("azienda")&"</a>"
							else
								response.write "Scheda cliente"
							end if
						end if
						if rs("fornitore")=1 then%>Scheda fornitore<%end if
					elseif oper="edit" then
						if  rs("fornitore")=0 then%>Modifica dati cliente<%end if
						if  rs("fornitore")=1 then%>Modifica dati fornitore<%end if
          else%>
          Nuovo cliente
          <%end if%>
          <%if oper="view" and (iduser>1 or  utente_andrea) then %>
          <span style="float: right;">
          <%if not in_tab then %>
          <%if ha_il_permesso("E1") then %>
          <%if rs("secondario")=0 or utente_andrea then %>
            <input name="edit" type="submit"  style="FONT: 8px;" value="Modifica" class="ui-button ui-widget ui-state-default ui-corner-all" >
            <%end if %>
            <button class="puls_giu">Altro...</button>
          </span>
          <ul style="display:none;">
            <li><a href="#" id="reset_pw">Resetta password a 123456</a></li>
			<li><a href="#" id="new_pw">Invia nuova password</a></li>
			<%if mod_dip then %>
			<li><a href="#" id="new_dip">Aggiungi dipendente</a></li>
			<li><a href="ordine_excel.php?iduser=<%=iduser%>" >Modulo ordine Excel</a></li>
			<li><a href="foglio-spettanze.asp?iduser=<%=iduser%>" >Vedi foglio spettanze cliente</a></li>
			<%end if %>
			<%if rs("secondario")=0 then %>
			<li><a href="#" id="rendi_secondario">Rendi secondario</a></li>

			<% end if %>
          </ul>
          <%end if 	'ha_il_permesso("E1")  %>
          <%
	          end if	'not in_tab
	          end if	'oper="view" and iduser>1
	          %>
          </td>
      </tr>
      <%if true then%>
      <%
      if oper<>"new" then
      %>
      <tr>
        <td colspan="2" valign="top"></td>
      </tr>
      <tr>
        <td width="20%">&nbsp;ID utente:</td>
        <td width="80%"><%=iduser%>
          <input type="hidden" name="iduser" id="iduser" value="<%=iduser%>"></td>
      </tr>
        <td valign="top">&nbsp;Nominativo</td>
        <td><b><%=rs("nome")&" "&rs("cognome")%></b></td>
      </tr>
      <tr>
        <td valign="top">&nbsp;Data registrazione</td>
        <td><%=rs("data")%></td>
      </tr>
      <tr>
        <td valign="top">&nbsp;Numero visite</td>
        <td><%if rs("numvisit")=0 then%>
          Nessuna
          <%else%>
          <%=rs("numvisit")%>
          <%end if%></td>
      </tr>
            <tr>
        <td valign="top">&nbsp;Secondario</td>
        <td>
          <%=rs("secondario")%>
          </td>
      </tr>

            <tr>
        <td valign="top">&nbsp;Ultima visita</td>
        <td><%if isnull(rs("lastvisit")) then%>
          Mai
          <%else%>
          <%=rs("lastvisit")%>
          <%end if%></td>
      </tr>
      <tr>
        <td valign="top">&nbsp;Ultimo invio password</td>
        <td><%if isnull(rs("password_inviata")) then%>
          Mai
          <%else%>
          <%=rs("password_inviata")%>
          <%

						sql="select count(*) FROM password_reset where  iduser="& iduser
						set rs_tmp= conn.execute (sql)
						if clng(rs_tmp(0))>0 then
							response.write ", "&rs_tmp(0)&" password da usare"
						end if
						set rs_tmp = nothing
				      %>


          <%end if%></td>
      </tr>



      <tr>
        <td colspan="2"><strong>Dati anagrafici</strong></td>
      </tr>
		<%end if
		if secondario=0 then
		%>

      <tr>
        <td>&nbsp;Fornitore</td>
        <td><%val=""
				   if oper="new" then
				else
					if rs("fornitore") then val=" checked"
				end if
				if fatture_fornitore>0 then
				%>
				  <input  type="checkbox" value="si" disabled="disabled" checked>
				  <input name="fornitore" type="hidden" value="si">

				<%else %>
          <input name="fornitore" id="fornitore" type="checkbox" value="si" <%=disabled%> <%response.write val%>>
		  <%end if%>
		  </td>
      </tr>
      <tr>
        <td>&nbsp;PA</td>
        <td><%
				val=""
				   if oper="new" then
				else
					if rs("trattamento_iva")=10 then val=" checked"
				end if%>
          <input name="" type="checkbox" value="si" id="pa" <%=disabled%> <%=val%>>
		  </td>
      </tr>
      <tr>
        <td>&nbsp;Split Payment</td>
        <td><%
				val=""
				   if oper="new" then
				else
					if rs("fattura_sp")=1 then val=" checked"
				end if%>
          <input name="fattura_sp" type="checkbox" value="si"  <%=disabled%> <%=val%>>
		  </td>
      </tr>
		<tr>
        <td>&nbsp;Tipo cliente:</td>
        <td><%if oper="view" then
	        response.write TipoCliente(pres_tipo)
          else

		  %>
          <select name="tipo" id="tipo2">
		  <%
		  if oper="new" then
			pres_tipo=""
			%>
		  <option  value="" <%if pres_tipo="" then response.write "selected"%> class="richiesto" <%=style%>>Seleziona il tipo cliente</option>
		  <%end if %>
			<option  value="S" <%if pres_tipo="S" then response.write "selected"%>>Societ&agrave;</option>
			<option  value="D" <%if pres_tipo="D" then response.write "selected"%>>Ditta individuale</option>
			<option  value="P" <%if pres_tipo="P" then response.write "selected"%>>Privato</option>
			<option  value="A" <%if pres_tipo="A" then response.write "selected"%>>Associazione</option>
			<option  value="X" <%if pres_tipo="X" then response.write "selected"%>>Sconosciuto</option>
          </select>

		 <a href="#" id="tabella_tipo"><span class="ui-icon ui-icon-help" style="display: inline-block; "></span></a>

          <%end if %><%=testo_tipo %></td>
      </tr>
		<%end if ' if secondario=0 %>
      <tr>
        <td valign="top">&nbsp;Cognome:</td>
        <td><%if oper="view" then%>
          <%=rs("cognome")%>
          <%else
				if pres_tipo<>"P" or pres_tipo<>"D" then
					richiesto=""
				else
					richiesto="richiesto"
				end if
		  %>
          <%if oper="new" then
					val=request.form("cognome")
				else
					val=rs("cognome")
				end if
				cognome_DefinedSize=rs("cognome").DefinedSize
				%>
          <input name="cognome" type="text" class="<%=richiesto%> i_text" id="cognome" value="<%=val%>" size="40" maxlength="<%=cognome_DefinedSize%>" onblur="nome_cognome();">
          <%end if%></td>
      </tr>
      <tr>
        <td valign="top">&nbsp;Nome:</td>
        <td><%if oper="view" then%>
          <%=rs("nome")%>
          <%else %>
          <%if oper="new" then
					val=request.form("nome")
				else
					val=rs("nome")
				end if
				nome_DefinedSize=rs("nome").DefinedSize
				azienda_DefinedSize=rs("azienda").DefinedSize
				citta_DefinedSize=rs("citta").DefinedSize

				%>
          <input name="nome" type="text" class="<%=richiesto%> i_text" id="nome" value="<%=val%>" size="40" maxlength="<%=nome_DefinedSize%>" onblur="nome_cognome();">
          <%end if %></td>
      </tr>
      <%if secondario=0 then %>
         <tr id="tr_azienda">
        <td>&nbsp;Azienda:</td>
        <td><%if oper="view" then%>
          <b><%=rs("Azienda")%></b>
          <%else%>
          <%if oper="new" then
					val=request.form("azienda")
				else
					val=server.htmlencode( rs("azienda"))
				end if
				 %>
          <input name="azienda" id="azienda" type="text" class="richiesto i_text" value="<%=val%>" size="40" maxlength="<%=azienda_DefinedSize%>" onblur="nome_cognome();">
          <%end if %></td>
      </tr>
      <%end if %>
         <tr>
        <td valign="top">&nbsp;Email attuale:</td>
        <td><%if oper="view" then%>
          <a href='mailto:<%=rs("email")%>'><%=rs("email")%></a>
          <%else%>
          <%if oper="new" then
					val=request.form("email")
					val2=""
				else
					val=rs("email")
					if val="" then
						val2="checked"
					end if
				end if
				email_DefinedSize=rs("email").DefinedSize
				 %>
          <input name="email" type="text" id="email" value="<%=val%>" size="40" class="richiesto i_text"><span style="position:absolute;"><div id="email_error" class="box_errore"></div></span><br><input type="checkbox" id="noemail" value="1" <%=val2%>>Nessuna mail
          <%end if %></td>
      </tr>
      <%if secondario=0 then %>
      <tr>
        <td>&nbsp;Indirizzo:</td>
        <td><%if oper="view" then%>
          <%=rs("Indirizzo")%>
          <%else%>
          <%if oper="new" then
					val=request.form("indirizzo")
				else
					val=rs("indirizzo")
				end if %>
          <input name="Indirizzo" type="text" value="<%=val%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Citta:</td>
        <td><%if oper="view" then%>
          <%=rs("Citta")%>
          <%else%>
          <%if oper="new" then
					val=request.form("citta")
				else
					val=rs("citta")
				end if
				 %>
          <input name="Citta" id="citta" type="text" value="<%=val%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Cap:</td>
        <td><%if oper="view" then%>
          <%=rs("Cap")%>
          <%else%>
          <%if oper="new" then
					val=request.form("cap")
				else
					val=rs("cap")
				end if %>
          <input name="Cap" type="text" value="<%=val%>" size="40" id="cap" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Provincia:</td>
        <td>
        <%
		if oper="view" then%>
		  <%=rs("provincia")%>
		  <%else
		if oper="new"  then
			val=request.form("provincia")
		else
			val=ucase(rs("provincia"))
		end if
		%>
			<select name="provincia" id="provincia" <%=disabled%>>
					<option value="">Selezionare</option>
				  <%
			set rs_comuni=conn.execute("select * from elenco_province order by denominazione_provincia")
			do while not rs_comuni.eof
				response.write "<option value='"&rs_comuni("sigla_automobilistica")&"'"
				if rs_comuni("sigla_automobilistica")=val then response.write " selected "
				response.write ">" & rs_comuni("denominazione_provincia")&"</option>"
				rs_comuni.movenext
			loop
			rs_comuni.close
			set rs_comuni=nothing
		%>
          </select>
		  <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Regione:</td>
        <td><%if oper="view" then
		if rs("regione")=0 then response.write "Non specificata" else response.write regione(rs("regione"))
			  %>
          <%else%>
          <%if oper="new" then
					val=request.form("regione")
				else
					val=rs("regione")
				end if %>
          <select name="regione" id="regione"  <%=style%>>
            <option value="0">Selezionare</option>
            <%
			for n=1 to regione(-1)
			response.write "<option value='"&n&"'"
			if n=val then response.Write(" selected ")
			response.write ">" & regione(n)&"</option>"
			next
			%>
          </select>
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Stato estero:</td>
        <td><%val=""
			   if oper="new" then

				else
					if rs("stato_estero")=1 then val=" checked"
				end if%>
          <input name="stato_estero" type="checkbox" id="stato_estero" value="si" <%=disabled%> <%response.write val%>></td>
      </tr>
    <% end if %>

      <tr>
        <td>&nbsp;Telefono:</td>
        <td><%if oper="view" then%>
          <%=rs("Telefono")%>
          <%else%>
          <%if oper="new" then
					val=request.form("telefono")
				else
					val=rs("telefono")
				end if
				telefono_DefinedSize=rs("telefono").DefinedSize
				 %>
          <input name="telefono" type="text" class="i_text" value="<%=val%>" size="40" maxlength="<%=telefono_DefinedSize%>" >
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Fax:</td>
        <td><%if oper="view" then%>
          <%=rs("Fax")%>
          <%else%>
          <%if oper="new" then
					val=request.form("fax")
				else
					val=rs("fax")
				end if %>
          <input name="Fax" type="text" value="<%=val%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Cellulare:</td>
        <td><%if oper="view" then%>
          <%=rs("Cellulare")%>
          <%else%>
          <%if oper="new" then
					val=request.form("cellulare")
				else
					val=rs("cellulare")
				end if %>
          <input name="Cellulare" type="text" value="<%=val%>" size="40" maxlength="<%=rs("cellulare").DefinedSize%>" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Ruolo:</td>
        <td><%if oper="view" then%>
          <%=rs("ruolo")%>
          <%else %>
          <input name="ruolo" type="text" value="<%=rs("ruolo")%>" size="40" class="i_text">
          <%end if%></td>
      </tr>



      <% if secondario=0 then %>
      <%if oper="view" then
	    	if rs("stato_estero")=0 then
				chkpiva=ControllaPIVA(trim(RS("piva")))
				chkcf=controllacf(trim(RS("cf")))
				if RS("cf")=RS("piva") then
					if chkcf="" or chkpiva="" then
						chkcf=""
						chkpiva=""
					end if

				end if
			end if
		end if%>
      <tr id="tr_piva" >
        <td>&nbsp;Partita IVA:</td>
        <td><%if oper="view" then%>
          <%=rs("piva")%>
          <%
				if chkpiva<>"" then response.write "<span style='color: red;'>"&chkpiva&"</span>"

				%>
          <%else%>
          <%if oper="new" then
					val=request.form("piva")
				else
					val=rs("piva")
				end if %>
          <input name="piva" type="text" id="piva" value="<%=val%>" size="40" class="richiesto i_text"><span style="position:absolute;"><div id="piva_error" class="box_errore"></div></span>
		<input type="hidden" name="piva_editabile" id="piva_editabile" value="<%=piva_editabile%>"><br><input type="checkbox" id="pivaduplicata" name="pivaduplicata" value="si" >Consenti esistente
          <%end if%></td>
      </tr>
      <tr id="tr_cf">
        <td>&nbsp;Codice Fiscale<span id="numerico"></span>:</td>
        <td><%if oper="view" then%>
          <%=rs("cf")%>
          <%
					if chkcf<>"" then response.write "<br><span style='color: red;'>"&chkcf&"</span>"
%>
          <%else%>
          <%if oper="new" then
					val=request.form("cf")
				else
					val=rs("cf")
				end if %>
          <input name="cf" type="text" id="cf" value="<%=val%>" size="40" class="richiesto i_text"><span style="position:absolute;"><div id="cf_error" class="box_errore"></div></span>
          <%end if%></td>
      </tr>
	   <tr>
        <td>&nbsp;Codice univoco</td>
        <td><%if oper="view" then
        	if not isnull(rs("FatturaPACodiceDestinatario")) or rs("FatturaPACodiceDestinatario")<>"" then
        		response.write replace(rs("FatturaPACodiceDestinatario"),vbcrlf,"<br>")
			end if
        %>

          <%else%>
          <%if oper="new" then
					val=request.form("FatturaPACodiceDestinatario")
					row=0
				else
					row=0
					val=rs("FatturaPACodiceDestinatario")
					if not isnull(val) then
						txt=split(val,vbcrlf)
						row=ubound(txt)+1
					end if
					if row<2 then row=2
				end if %>
			<textarea name="FatturaPACodiceDestinatario" id="FatturaPACodiceDestinatario" cols="50" rows="<%=row%>" class="i_text"><%=val%></textarea><br>
			Per inserire più di un codice univoco utilizzare la seguente sintassi su ogni riga:<br>codice univoco:descrizione<br>
          <span style="position:absolute;"></span>
          <button id="cerca_pa">Cerca nella PA</button> per codice fiscale
           <script>
$(function () {
    $("#cerca_pa").on("click", function (e) {

				e.preventDefault();
				var cf=$("#cf").val();
				if (cf==""){
						var dialog=$("#dialog").dialog();
						var txt="La ricerca del codice univoco avviene tramite il codice fiscale.<br>Devi inserire il codice fiscale";
						dialog.html(txt);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 600,
							height: "auto",
							title: "Ricerca codice univoco"
						});
						return;


				}
				$.ajax({
					url     : "https://www.indicepa.gov.it/public-ws/WS01_SFE_CF.php",
					type    : "post",
					cache	: false,
					//dataType: 'json',
					data	: {
						AUTH_ID :"RPBJPION",
						CF : cf


					},
					success: function(data){
						console.log(data);
						var txt="";
						if (data.cod_err>0){
							txt="Errore nella ricerca codice univoco: "+cf+" errore:"+data.desc_err;
							$.ajax({
								url     : "searcher.asp",
								type    : "post",
								data	: "txt_errore="+encodeURIComponent(txt)
							});
						}
						if(data.num_items==0){
							txt="Nessuna corrispondenza trovara per il codice fiscale "+cf;
						}
						else
						{




							var result=data.result;
							var data2=data.data;
							if(result.num_items==1){
								txt="Trovata 1 corrispondenza<br>";
							}else
							{
								txt="Trovate "+result.num_items+" corrispondenze<br>";
							}

							for(var i = 0; i < data2.length; i++) {
							    var obj = data2[i];
								txt+="Descrizione: <b>"+obj.des_amm+"</b><br>";
								for(var n = 0; n < obj.OU.length; n++ ){
									var obj2 = obj.OU[n];
									txt+="&nbsp; Ufficio: "+obj2.des_ou+"<br>";
									txt+='&nbsp; Codice univoco: '+obj2.cod_uni_ou+'<input type="button" value="Aggiungi" onclick="javascript: codiceunivoco(\''+obj2.cod_uni_ou+'\'+\':\'+\''+obj2.des_ou+'\');"/><hr>';

								}


							}


						}
						var dialog=$("#dialog").dialog();
						dialog.html(txt);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 600,
							height: "auto",
							title: "Ricerca codice univoco"
						});


					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nella ricerca codice univoco');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>querystring: ?"+querystring[1];
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
				});
			  });
			  function codiceunivoco(stringa){
				  $("#salva").removeAttr("disabled");
				  var FatturaPACodiceDestinatario=$("#FatturaPACodiceDestinatario").val();
				  if (FatturaPACodiceDestinatario==""){
					  $("#FatturaPACodiceDestinatario").val(stringa);
				  }else{
					  $("#FatturaPACodiceDestinatario").val(FatturaPACodiceDestinatario+'\n'+stringa);
				  }
			  }
			  </script>
          <%end if%>
          </td>
      </tr>
      <%end if %>

           <tr>
        <td>&nbsp;Note su utente</td>
        <td><%if oper="view" then%>
          <%=rs("note_su_utente")%>
          <%else%>
          <%if oper="new" then
					val=request.form("note_su_utente")
				else
					val=rs("note_su_utente")
				end if %>
          <textarea name="note_su_utente" cols="50" rows="3" id="note_su_utente" class="i_text"><%=val%></textarea>
          <%end if%></td>
      </tr>
      <%if oper="edit" and false then%>
      <tr>
        <td colspan="2"><strong>Destinazione ordine</strong></td>
      </tr>
      <tr>
        <td>&nbsp;Azienda:</td>
        <td><%if oper="view" then%>
          <%=rs("D_Azienda")%>
          <%else %>
          <input name="D_Azienda" type="text" value="<%=rs("d_azienda")%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Indirizzo:</td>
        <td><%if oper="view" then%>
          <%=rs("D_Indirizzo")%>
          <%else %>
          <input name="D_Indirizzo" type="text" value="<%=rs("D_Indirizzo")%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Citta:</td>
        <td><%if oper="view" then%>
          <%=rs("D_Citta")%>
          <%else %>
          <input name="D_Citta" type="text" value="<%=rs("D_Citta")%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Provincia:</td>
        <td><%if oper="view" then%>
          <%=rs("D_provincia")%>
          <%else %>
          <input name="D_provincia" type="text" value="<%=rs("D_provincia")%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Cap:</td>
        <td><%if oper="view" then%>
          <%=rs("D_cap")%>
          <%else %>
          <input name="D_cap" type="text" value="<%=rs("D_cap")%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <%end if%>
      <%if secondario=0 then %>
            <tr>
        <td colspan="2"><strong>Dati per ordine e fatturazione</strong></td>
      </tr>
      <tr>
        <td>&nbsp;Pagamento concordato:</td>
        <td><%if oper="view" then%>
          <%=metodipagamento.descrizione(rs("pag_accordato"))%>
          <%else%>
          <%if oper="new" then
					val=request.form("pag_accordato")
				else
					val=rs("pag_accordato")
				end if %>
          <select name="pag_accordato">

            <option value="0">Nessuno</option>
                <%call metodipagamento.stampa_option(val)%>
          </select>
          <%end if%></td>
      </tr>
		<tr>
        <td>&nbsp;Modalità trasporto:</td>
        <td><%if oper="view" then%>
          <%=tipo_trasporto(rs("tipo_trasporto_2"))%>
          <%else%>
          <%if oper="new" then
					val=request.form("tipo_trasporto_2")
				else
					val=rs("tipo_trasporto_2")
					if val <>"" or not isnull(val) then val=cint(val)
				end if %>
          <select name="tipo_trasporto_2">
            <option value="0">Nessuno</option>
            <%
			for n=1 to tipo_trasporto(-1)
			response.write "<option value='"&n&"'"
			if val=n then response.write " selected "
			response.write ">" & tipo_trasporto(n)&"</option>"
			next
			%>
          </select>
          <%end if%></td>
      </tr>
   <tr>
        <td>&nbsp;Note spedizione:</td>
        <td><%if oper="view" then%>
          <%=rs("note_trasporto_2")%>
          <%else%>
          <%if oper="new" then
					val=request.form("note_trasporto_2")
				else
					val=rs("note")
				end if %>
          <input name="note_trasporto_2" type="text" value="<%=val%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Banca d'appoggio</td>
        <td><%if oper="view" then%>
          <%=rs("banca_appoggio")%>
          <%else%>
          <%if oper="new" then
					val=request.form("banca_appoggio")
				else
					val=rs("banca_appoggio")
				end if %>
          <input name="banca_appoggio" type="text" id="banca_appoggio" value="<%=val%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;IBAN</td>
        <td><%if oper="view" then%>
          <%
		  lunghezza_iban=len(rs("iban"))
		  if lunghezza_iban<>27 and lunghezza_iban>0 then errore_iban=true
		  if errore_iban then%>
          <font color="red">
          <%end if%>
          <%=rs("iban")%>
          <%if errore_iban then%>
		   &nbsp;lunghezza iban errata, <%=lunghezza_iban%> caratteri invece di 27
          </font>
          <%end if%>
          <%else%>
          <%if oper="new" then
					val=request.form("iban")
				else
					val=rs("iban")
				end if %>
          <input name="iban" type="text" id="iban" value="<%=val%>" size="40" class="i_text">
          <%end if%></td>
      </tr>
       <tr>
        <td valign="top">&nbsp;Email per invio fattura</td>
        <td><%if oper="view" then%>
          <%=rs("email_inoltro_fattura")%>
          <%else %>
          <%if oper="new" then
					val=request.form("email_inoltro_fattura")
				else
					val=rs("email_inoltro_fattura")
				end if %>
          <input name="email_inoltro_fattura" type="text" id="email_inoltro_fattura" value="<%=val%>" size="40" class="i_text"><br>
          Se valorizzato questo campo non viene più preso in considerazione il campo <b>Email attuale</b>, separare più indirizzi con <b>;</b>
          <%end if %></td>
      </tr>
      <tr>
        <td>&nbsp;No spese bancarie</td>
        <td><%val=""
							   if oper="new" then

				else
					if rs("spese_0") then val=" checked"
				end if%>
          <input name="spese_0" type="checkbox" id="spese_0" value="si" <%=disabled%> <%response.write val%>></td>
      </tr>
      <%
	      'Permesso per agente
	     if ha_il_permesso("Z1") or session("iduser")=4 then
	     'if ha_il_permesso("Z1") or oper="view" then

         %>
      <tr>
        <td>&nbsp;Agente:</td>
        <td><%
		if oper="view" then
			val=rs("idagente")
			if val<>"" then
				set rsm=conn.execute("select * from agenti where idagente= "&val)
				if not rsm.eof then response.write rsm("nominativo")

			end if
		else

		if oper="new" then
					val=request.form("idagente")
				else
					val=rs("idagente")
				end if%>
          <select name="idagente" <%=disabled%>>
            <option value="0">Nessuno</option>
            <%
			set rsm=conn.execute("select * from agenti order by nominativo ;")
			do while not rsm.eof
			response.write "<option value=""" & rsm("idagente") &""""
			if val=rsm("idagente") then response.write " selected"
			response.write ">" & rsm("nominativo")& "</option>"
			rsm.movenext
			loop
			rsm.close
			set rsm = Nothing
	%>
          </select>
<%end if
		  %>
		  </td>
      </tr>
      		  <%end if
		  %>
      <tr>
        <td>&nbsp;Trattamento iva:</td>
        <td><%if oper="view" then
				response.write trattamento_iva(rs("trattamento_iva"))
			else
				if oper="new" then
					val=request.form("trattamento_iva")
				else
					val=rs("trattamento_iva")
				end if
			%>
          <select name="trattamento_iva" id="trattamentoiva" <%=disabled%>>
			<%for n= 0 to trattamento_iva(-1) %>
            <option value="<%=n%>" <%if val=n then response.write " selected"%> ><%=trattamento_iva(n)%></option>
			<%next %>
          </select>
		  <%end if%></td>
      </tr>
      <tr>
        <td>&nbsp;Banca di appoggio:</td>
        <td><%if oper="view" then
				val=rs("idbanca")
				if val<>"" then
					set rsm=conn.execute("select * from banche where idbanca="&val)
					if not rsm.eof then response.write rsm("nome_breve")

				end if
			else
				if oper="new" then
					val=request.form("adbanca")
				else
					val=rs("idbanca")
				end if %>
          <select name="idbanca"  <%=disabled%>>
            <option value="0">Predefinita</option>
            <%
			set rsm=conn.execute("select * from banche order by nome_breve;")
			do while not rsm.eof
			%>
						<option value="<%=rsm("idbanca")%>" <%if val=rsm("idbanca") then response.write " selected"%>><%=rsm("nome_breve")%></option>
						<%
			rsm.movenext
			loop
			rsm.close
			set rsm = Nothing
			%>
              <option value="1" <%if val="" or isnull(val) then response.write " selected"%> >Non impostata</option>
          </select>
		  <%end if%>
		  </td>
      </tr>
            <tr>
        <td>&nbsp;Tipologia cliente</td>
        <td><%if oper="view" then

	        		response.write vedi_tipologia(RS("tipologia"))
        		else



		        if oper="new" then
					val=","&request.form("tipologia")&","
				else
					val=","&rs("tipologia")&","
				end if
				%>
          <select name="tipologia"  id="tipologia" multiple style="width:400px;" <%=disabled%>>
		  <%
		set rsm=conn.execute("select * from utenti_tipologia order by tipologia;")
		do while not rsm.eof
		%>
		            <option value="<%=rsm("id")%>" <%if instr(val,","&rsm("id")&",")>0 then response.write " selected"%>><%=rsm("tipologia")%></option>
		            <%
		rsm.movenext
		loop
		rsm.close
		set rsm = Nothing
		%>
          </select>
          <%end if%></td>
      </tr>
      <%end if %>
    <tr>
		<td colspan="2"><strong>Altro</strong></td>
	</tr>
      <tr>
        <td>&nbsp;Vedi prezzi riservati</td>
        <td><%val=""
			   if oper="new" then

				else
					if rs("vedi_prezzi")=1 then val=" checked"
				end if%>
          <input name="vedi_prezzi" type="checkbox" value="si" <%=disabled%> <%response.write val%>></td>
      </tr>


       <tr>
        <td>&nbsp;Non eliminare</td>
        <td><%val=""
			   if oper="new" then

				else
					if rs("non_eliminare") then val=" checked"
				end if%>
          <input name="non_eliminare" type="checkbox" value="si" <%=disabled%> <%response.write val%>></td>
      </tr>




       <tr>
        <td>&nbsp;Seleziona per mailing list</td>
        <td><%val=""
			   if oper="new" then

				else
					if rs("selezmailing") then val=" checked"
				end if%>
          <input name="selezmailing" type="checkbox" value="si" <%=disabled%> <%response.write val%>></td>
      </tr>

       <%if mod_larochelle then
  			   if oper="new" then
					permessi=permessi_utente_default()
				else
					permessi= rs("permessi_utente")

				end if %>
    <tr>
		<td colspan="2"><strong>Permessi cliente</strong></td>
	</tr>

       <tr>
        <td>&nbsp;Dipendenti e foglio spettanze:</td>
        <td><%
				val=""
				if instr(permessi,"A1=1")>0 then val=" checked"
				%>
			<input name="A1" type="checkbox" value="si" <%=disabled%> <%response.write val%>>A1</td>
      </tr>
       <tr>
        <td>&nbsp;Scarica ordine Excel:</td>
        <td><%
				val=""
				if instr(permessi,"B1=1")>0 then val=" checked"
				%>
			<input name="B1" type="checkbox" value="si" <%=disabled%> <%response.write val%>>B1</td>
      </tr>



      <%end if %>
	<%if oper<>"new" then %>
	<tr>
        <td>&nbsp;Lingua:</td>
        <td><%=rs("lingua")%></td>
      </tr>
      <tr>
        <td>&nbsp;Consenso a mailing:</td>
        <td><%
				if rs("mailing") then%>
          SI
          <%else%>
          NO
          <%end if%></td>
      </tr>


    <tr>
		<td colspan="2"><strong>Altre info</strong></td>
	</tr>

      <tr>
        <td>&nbsp;Ultima mailing:</td>
        <td><%=rs("lastmailing")%></td>
      </tr>
    <tr>
        <td valign="top">&nbsp;Email alla registrazione</td>
        <td><%=rs("1email")%></td>
      </tr>
      <tr>
        <td>&nbsp;Traccia log accessi:</td>
        <td><%=rs("idloged")%></td>
      </tr>
      <tr>
        <td valign="top">&nbsp;Modalità registrazione</td>
        <td><%=modalitaRegistrazione(rs("modalita_registrazione"))%>
        	<%if rs("modalita_registrazione")=1 then
        		if rs("registrato_da_iduser")>0 then
	        		set rs_admin=conn.execute ("select nominativo from admin where iduser="&rs("registrato_da_iduser"))
	        		if not rs_admin.eof then
		        		response.write " "&rs_admin("nominativo")
		        	end if
		        	set rs_admin=Nothing
		        end if


          end if%>
          </td>
      </tr>
      <tr>
        <td valign="top">&nbsp;Abilitato</td>
        <td><%
				if rs("abilitato") then%>
          SI
          <%else%>
          NO
          <%end if%></td>
      </tr>





      <%end if%>
      <%
		if oper<>"new" then

			if rs("fornitore") then
				classe="class=""trfornitore"""
				sql="select utenti_fornitori.* from utenti_fornitori where iduser="&iduser
				set rs_for=conn.execute(sql)
				if not rs_for.eof then importo_minimo=rs_for("importo_minimo") else importo_minimo=""
				set rs_for=nothing
			else
				classe="class=""trfornitore"" style=""display: none;"""
			end if
		else
			classe="class=""trfornitore"" style=""display: none;"""
		end if
	  %>
         <tr <%=classe%>>
           <td colspan="2"><strong>Dati fornitore</strong></td>
         </tr>
		<tr <%=classe%>>
	        <td >Importo minimo ordine</td>
	        <td><%
		if oper="view" then%>
          <%=importo_minimo%>
          <%else
		  %>
          <input name="importo_minimo" type="text" value="<%=importo_minimo%>" size="40" class="i_text"/>
          <%end if%>
          </td>
      </tr>
      <%

%>

		  <%if oper<>"new" then%>
	   <tr>
        <td>&nbsp;Record utenti_clienti</td>
        <td><%=not isnull(rs("stato_estero"))%></td>
      </tr>
	        <tr>
        <td>&nbsp;CHK</td>
        <td><%=chkDataUser(rs("data"))%></td>
      </tr>
	<%end if%>
	  <%if oper<>"new" and session("idadmin")="1" then%>

	        <tr>
        <td>&nbsp;HTTP_USER_AGENT</td>
        <td><%=rs("HTTP_USER_AGENT")%></td>
      </tr>
	<%end if%>

	<% if oper<>"new" then
		if rs("fornitore") then
			document_title="Fornitore "&server.htmlencode(denominazione(rs("nome"),rs("cognome"),rs("azienda")))
		else
			document_title="Cliente "&server.htmlencode(denominazione(rs("nome"),rs("cognome"),rs("azienda")))
		end if
	else
		document_title="Nuovo cliente o fornitore"
		end if
	%>
      <%end if%>
      <%if oper="edit" or oper="new" then%>
      <tr>
        <td colspan="2" align="center"><%
		if oper="view" then
			txt=puls_new
		elseif oper="new" then
			txt="Aggiungi"
		elseif oper="edit" then
			minuti=datediff("n",rs("lastvisit"),now())
			if minuti<20 then
				response.write "<span class=""rosso"">Attenzione, l'utente ha effettuato l'accesso "&minuti&" minuti fa</span><br>"
			end if
			txt="Salva modifiche"
		end if
		Set rs = Nothing
		%>
          <input type="submit" value="<%=txt%>" id="salva" <%if testo_tipo="" then %>disabled<%end if %>>
          <%if oper<>"view" then%>
          <input type="submit" name="annulla" value="Esci senza salvare" onClick="annulla=true;" class="cancel">
          <%end if%>
          <%if oper="edit" and iduser>1 then%>
          <input type="submit" name="elimina_utente" value="Elimina" <%if cancellabile=false then%> disabled <%end if%> class="cancel">
          <%end if%>
          <%if oper="edit" then
			oper="update"
		elseif oper="new" then
			oper="add"
		end if
		%>
          <input type="hidden" name="oper" value="<%=oper%>">
          </td>
      </tr>
      <%end if
			'end if
			%>
    </table>
<%if oper="view" then %>
<script>
	$(function() {




		puls_giu();
					$("#rendi_secondario").click(function(e){
				var dialog=$("#dialog").dialog();
				dialog.html("Attendi...");
				$.ajax({
					url     : "pag_adm_dialog_secondario.asp?iduser=<%=iduser%>",
					type    : "post",
					cache	: false,
					//dataType: 'json',
					//data	: dati,
					success: function(data){
						dialog.html(data);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 600,
							height: "auto",
							title: "Seleziona il cliente principale"
						});
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>querystring: ?"+querystring[1];
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});

	});
</script>
<%end if %>

<%if oper<>"view" then %>
<script type="text/javascript" src="Jquery/js/jquery.validate.min.js"></script>
<script type="text/javascript" src="Jquery/js/jquery.validate.min_it.js"></script>
<script>
			var tipo="";
			var piva_editabile="<%=piva_editabile%>";
			var txtErrori;
			var come_tipo="<%=come_tipo%>";
			imposta_form($("#tipo2").val());
			$("#tipo2").change(function () {
				var val=$("#tipo2").val();
						console.log("tipo onChange:"+val);
						tipo=val;
						imposta_form(tipo);
			});






			//$("#salva").click(function(e){
			//	console.log("click");
			//	$("#form_user").validate();
			//});
			$.validator.addMethod("codice_fiscale", function (value, element) {
			result = true;
			var tipo=$("#tipo2").val();
			if (piva_editabile=="SI"){
				if (tipo=="X"){
					chk="";
				}
				else if ($("#stato_estero").is(":checked")){
					chk="";
				}
				else if (tipo=="S" || tipo=="A")
				{
					chk=ControllaCFNumerico(value);
					txtErrori+="ControlloCF: ControllaCFNumerico:"+value;
				}
				else
				{
					chk=ControllaCF(value);
					txtErrori+="ControlloCF: ControllaCF:"+value;
				}
				if (chk != '' ) {
					$.validator.messages.codice_fiscale = chk;
					result = false;
				}
				if (value=="" && tipo!="X")
				{
					$.validator.messages.codice_fiscale = "Inserire il codice fiscale";
					result = false;
				}
			}
			return result;
		}, "");
		$.validator.addMethod("partita_iva", function (value, element) {
			result = true;
			var tipo=$("#tipo2").val();
			if (tipo=="X" || tipo=="A"){
				result=true;
			}
			else if ($("#stato_estero").is(":checked")){
				result=true;
			}
			else if (tipo!="P"){
				chk=ControllaPIVA(value);
				if (chk != '' ) {
					$.validator.messages.partita_iva = chk;
					result = false;
				}
				if ((tipo!="P") && (value=="")){
					$.validator.messages.partita_iva = "Inserire la partita iva, se &egrave; un privato seleziona privato nel campo tipo cliente";
					result = false;
				}
			}
			return result;
		}, "");
		$.validator.addMethod("almeno_uno", function (value, element) {
			result = true;
			var tipo=$("#tipo2").val();
			if (tipo=="X"){
				if($("#nome").val()=="" && $("#cognome").val()=="" && $("#azienda").val()==""){
					result= false;
					$.validator.messages.nome = "Compilare almeno uno dei tre campi nomi, cognome, azienda";
					console.log("qui");
			}
			}
			return result;
		}, "Compilare almeno uno dei tre campi nomi, cognome, azienda");



		var validator=$("#form_user").validate({
			rules: {
				tipo: { required: true},
				email: {required:{ depends: function(){
					return !$("#noemail").prop("checked");
				}


				},
						email: true,maxlength:<%=email_DefinedSize%>,
				        remote:{
							url: "searcher.asp?check_email=si&iduser=<%=iduser%>", //make sure to return true or false with a 200 status code
							type: "post",
							dataType: 'json',
							cache:false,
							dataFilter: function(data) {
								var json = JSON.parse(data);
								if(json.isError == "true") {
									return "\"" + decodeURIComponent(escape(json.errorMessage)) + "\"";
								} else {
									return true;
								}
							},
							error:function(xhr, textStatus, error,data){
									txt="Chiamata AJAX validazione email<br>";
									txt+="<br>iduser:<%=iduser%>, email:"+$("#email").val();
									  txt+="<br>xhr.statusText:"+xhr.statusText;
									  txt+="<br>xhr.responseText:"+xhr.responseText;
									  txt+="<br>textStatus:"+textStatus;
									  txt+="<br>error:"+error;
									  txt+="<br>data:"+data;
								    $.ajax({
									cache: false,
									url     : "searcher.asp?titolo=Validazione%20form%20non%20riuscita",
									type    : "post",
									data	: "txt_errore="+encodeURIComponent(txt)
									});
							}
						}},
				nome: {required:{
							depends: function(element){

								var tipo=$("#tipo2").val();
								if (tipo=="S" || tipo=="A" ){
									return false;
								}else if (tipo=="X"){
									if($("#nome").val()=="" && $("#cognome").val()=="" && $("#azienda").val()==""){
										return true;

									 }else{
										 return false;
									 }}

								else {
										return true;
									}

				        }},minlength:3,maxlength:<%=nome_DefinedSize%>},
				cognome: {required:{
							depends: function(element){
								var tipo=$("#tipo2").val();
								if (tipo=="S" || tipo=="A"){
									return false;
								}
								else if (tipo=="X"){
									if($("#nome").val()=="" && $("#cognome").val()=="" && $("#azienda").val()==""){
										$.validator.messages.cognome = "Compilare almeno uno dei tre campi nomi, cognome, azienda";
										return true;

									 }else{
										 return false;
									 }}

								else {
										$.validator.messages.cognome = "Campo obbligatorio, se &egrave; un privato o ditta individuale";
										return true;
									}

				        }},minlength: 3,maxlength:<%=cognome_DefinedSize%>},
				azienda: {required: {
							depends: function(element){
								var richiesto=$("#azienda").hasClass("richiesto");
								var result=false;
								console.log("azienda, richiesto:"+richiesto);
								var tipo=$("#tipo2").val();
								if (richiesto){
									$.validator.messages.azienda = "Campo obbligatorio, se non &egrave; un azienda seleziona privato o ditta individuale nel campo tipo cliente";
										result= true;
									}
									else
									{
										if( tipo=="X" && ($("#nome").val()=="" && $("#cognome").val()=="" && $("#azienda").val()=="")){
											$.validator.messages.azienda = "Compilare almeno uno dei tre campi nomi, cognome, azienda";
										result= true;

									 }else{
										 result= false;
									 }





									}
									console.log ("Azienda richiesto:"+result);
								return result;

							//return true;
				        }},maxlength:<%=azienda_DefinedSize%>},
				<%if mod_larochelle="" then %>
				cf: { codice_fiscale: true,
					remote:{
							url: "searcher.asp?chkcf=si&iduser=<%=iduser%>", //make sure to return true or false with a 200 status code
							type: "post",
							dataType: 'json',
							cache:false,
							dataFilter: function(data) {
								var json = JSON.parse(data);
								if(json.isError == "true") {
									return "\"" + decodeURIComponent(escape(json.errorMessage)) + "\"";
								} else {
									return true;
								}
							},
							}
					},
				<%end if %>
				piva: { partita_iva: true ,
					remote:{
							url: "searcher.asp?chkpiva=si&iduser=<%=iduser%>", //make sure to return true or false with a 200 status code
							type: "post",
							dataType: 'json',
							cache:false,
							dataFilter: function(data) {
								var json = JSON.parse(data);
								if(json.isError == "true") {
									if($("#pivaduplicata").is(":checked")){
										return true;
									}
									else{
									return "\"" + decodeURIComponent(escape(json.errorMessage)) + "\"";
									}
								} else {
									return true;
								}
							},
							}
					},
				indirizzo: {required: true,minlength: 5},
				citta: {required: true,minlength: 3,maxlength:<%=citta_DefinedSize%>},
				cap: {required: true,number: true,minlength: 5, maxlength: 5}
				//,provincia: { required: {
				//			depends: function(){
				//				if ($('#stato_estero').is(':checked')){
				//					return false;
				//				} else
				//				{ return true;
				//				}			}}
				//		}
				},

				messages:{ azienda:{required: "Campo obbligatorio, se &egrave; un privato seleziona privato nel campo tipo cliente."}

				},
			ignore: "",
		    invalidHandler: function(e, validator){
	           if(validator.errorList.length){
		           var txt_errori='';
		           //Apro tutti i panel con errori
		           for (var i=0;i<validator.errorList.length;i++){
				        txt_errori+=$(validator.errorList[i].element).attr("name")+',';

				    }
		        console.log("errore in:"+txt_errori);
				}
	        },
	        submitHandler: function(form){
				form.submit();
			}
		});

		    $("#pa").change(function () {
			    if ($(this).prop("checked")){
				    $("#trattamentoiva").val(10);


			    }else{
				    $("#trattamentoiva").val(0);
			    }
			});
		    $("#fornitore").change(function () {
			    $(this).prop("checked")?$(".trfornitore").show():$(".trfornitore").hide();
			});
		    $("#noemail").change(function () {
			    $(this).prop("checked")?$("#email").removeClass("richiesto"):$("#email").addClass("richiesto");
			});


			$("#tabella_tipo").click(function(e){
				var dialog=$("#dialog").dialog();
				dialog.html("Attendi...");
				$.ajax({
					url     : "ajax_function.asp?oper=tabella_tipo",
					type    : "post",
					cache	: false,
					//dataType: 'json',
					//data	: dati,
					success: function(data){
						dialog.html(data);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 600,
							height: "auto",
							title: "Dati richiesti per tipo cliente "
						});
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>querystring: ?"+querystring[1];
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});






			function imposta_form(tipo){
			switch(tipo) {
				case "S":
					$("#tr_piva").show();
					$("#tr_azienda").show();
					$("#nome").removeClass("richiesto");
					$("#cognome").removeClass("richiesto");
					$("#piva").addClass("richiesto");
					$("#cf").addClass("richiesto");
					$("#azienda").addClass("richiesto");
					$("#numerico").html(" numerico");

					break;
				case "D":
					$("#tr_piva").show();
					$("#tr_azienda").show();
					$("#azienda").removeClass("richiesto");
					$("#nome").addClass("richiesto");
					$("#cognome").addClass("richiesto");
					$("#piva").addClass("richiesto");
					$("#cf").addClass("richiesto");
					$("#numerico").html("");
					//Nascondo messaggi di errore
					$("#azienda-error").hide();
					break;
				case "P":
					$("#tr_piva").hide();
					$("#tr_azienda").hide();
					$("#azienda").removeClass("richiesto");
					$("#piva").val("");
					$("#nome").addClass("richiesto");
					$("#cognome").addClass("richiesto");
					$("#piva").removeClass("richiesto");
					$("#cf").addClass("richiesto");
					$("#numerico").html("");
					//Nascondo messaggi di errore
					$("#piva-error").hide();
					$("#azienda-error").hide();
					break;
				case "A":
					$("#tr_azienda").show();
					$("#azienda").addClass("richiesto");
					$("#tr_piva").hide();
					$("#piva").val("");
					$("#nome").removeClass("richiesto");
					$("#cognome").removeClass("richiesto");
					$("#piva").removeClass("richiesto");
					$("#cf").addClass("richiesto");
					$("#numerico").html("");
					//Nascondo messaggi di errore
					$("#piva-error").hide();
					$("#azienda-error").hide();
					break;
				case "X":
					console.log("sconosciuto");
					$("#tr_piva").show();
					$("#tr_azienda").show();
					$("#azienda").removeClass("richiesto");
					$("#nome").removeClass("richiesto");
					$("#cognome").removeClass("richiesto");
					$("#piva").removeClass("richiesto");
					$("#cf").removeClass("richiesto");
					$("#numerico").html("");
					//Nascondo messaggi di errore
					$("#piva-error").hide();
					$("#azienda-error").hide();
					break;
			}
	}
	</script>
    <%end if %>

	<%

	'***************************************************************************************************************************************************************************************
		case 2	'ordini
	'***************************************************************************************************************************************************************************************

			sql="select ordini.idord, ordini.nord, ordini.data, ordini.stato, ordini.totale, sommadiincassi.SommaDiimporto, ordini_fatture.idfat, fatture.Nfat FROM (((ordini LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) ) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDFat) LEFT JOIN (select incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idord ) AS sommadiincassi ON ordini.idord = sommadiincassi.idord where ordini.tipo_documento='ordine' and ordini.eliminato=0 and ordini.iduser="&iduser
			if not ha_il_permesso("C5") then
				sql=sql&" and creato_da_admin="&sessioniduser
			end if
			sql=sql&" order by ordini.data desc"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Vedi</td><td align=""center"">N&deg; ordine</td><td align=""center"">Data</td><td align=""center"">stato</td><td align=""center"">Totale</td></tr>"&vbcrlf

				For i = LBound(arrRS, 2) To UBound(arrRS, 2)

					testo_differenza=""
					sommadiincassi=arrRS(5, i)
					if isnull(sommadiincassi) then
						sommadiincassi=0
					else
						sommadiincassi=cdbl(sommadiincassi)
					end if
					totale_ordine=cdbl(arrRS(4, i))
					if totale_ordine<0 then	'totale<0
						colore="bgcolor='#99FF99'"
					elseif sommadiincassi=0 then 'isnull(sommadiimporto)

						colore=""
					elseif totale_ordine=sommadiincassi then 'totale=sommadiimporto
						colore="bgcolor='#99FF99'" 'Verdino saldata
					elseif totale_ordine>sommadiincassi then 'totale>sommadiimporto
						colore="bgcolor='#FFCCCC'" 'Rosino da saldare
						testo_differenza="("&formatcurrency(sommadiincassi-totale_ordine,2)&")"
					else
						colore="bgcolor='#CC00CC'" 'Viola
						testo_differenza="(+"&formatcurrency(sommadiincassi-totale_ordine,2)&")"
					end if
					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center""><input type=""radio"" onClick=""location.href='pag_adm_ordini.asp?idord="&arrRS(0, i)&"'""></td><td align=""center"">"&arrRS(1, i)&"</td><td align=""center"">"&formatdatetime(arrRS(2, i),2)&"</td><td align=""center"">"&stato_ordine(arrRS(3, i))&"</td><td align=""center"">"&formatnumber(arrRS(4, i),2)&"</td></tr>"


				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If
	'***************************************************************************************************************************************************************************************
		case 3	'fatture
	'***************************************************************************************************************************************************************************************
			sql="select fatture.IDFat, fatture.Nfat, fatture.data, fatture.causale, cast(fatture.totale_fattura as decimal (10,2)), sommadiincassi.SommaDiimporto FROM fatture LEFT JOIN (select incassi.idfat, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idfat  )  AS sommadiincassi ON fatture.IDFat = sommadiincassi.idfat WHERE fatture.iduser="&iduser&" order by fatture.data desc"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			Set ccausale_ddt = new cl_causale_ddt
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Vedi</td><td align=""center"">N&deg; fattura</td><td align=""center"">Data</td><td align=""center"">Causale</td><td align=""center"">Totale</td></tr>"&vbcrlf

				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					sommadiincassi=arrRS(5, i)
					totale_fattura=cdbl(arrRS(4, i))
					if isnull(sommadiincassi) then
						sommadiincassi=0
					else
						sommadiincassi=cdbl(sommadiincassi)
					end if
					testo_differenza=""
					if totale_fattura<0 then	'totale<0
						colore="bgcolor='#99FF99'"
					elseif sommadiincassi=0 then 'isnull(sommadiimporto)

						colore=""
					elseif totale_fattura=sommadiincassi then 'totale=sommadiimporto
						colore="bgcolor='#99FF99'" 'Verdino saldata
					elseif totale_fattura>sommadiincassi then 'totale>sommadiimporto
						colore="bgcolor='#FFCCCC'" 'Rosino da saldare
						testo_differenza="("&formatcurrency(sommadiincassi-totale_fattura,2)&")"
					else
						colore="bgcolor='#CC00CC'" 'Viola
						testo_differenza="(+"&formatcurrency(sommadiincassi-totale_fattura,2)&")"
					end if
					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center""><input type=""radio"" onClick=""location.href='pag_adm_fatture.asp?idfat="&arrRS(0, i)&"'""></td><td align=""center"">"&pre_fattura(arrRS(2, i))&arrRS(1, i)&"</td><td align=""center"">"&formatdatetime(arrRS(2, i),2)&"</td><td align=""center"">"&ccausale_ddt.elemento(arrRS(3, i))&"</td><td align=""center"">"&totale_fattura&"</td></tr>"


				Next
				set ccausale_ddt = Nothing
				stringa=stringa&"</table>"
				Erase arrRS
			End If
	'***************************************************************************************************************************************************************************************
		case 4	'DDT
	'***************************************************************************************************************************************************************************************
			sql="select ddt.idord, ordineddt.nord as nddt, ordineddt.data, ddt.causale, ordini.nord, ordini.idord, fatture.IDfat, fatture.Nfat, ddt.reso_tutto, ordini.stato, ordini.totale, sommadiincassi.SommaDiimporto, ordini.stato FROM ((((ordini LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) INNER JOIN ddt ON ordini.idord = ddt.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDFat) INNER JOIN utenti ON ordini.iduser = utenti.iduser) LEFT JOIN (select incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idord  )  AS sommadiincassi ON ordini.idord = sommadiincassi.idord inner join ordini ordineddt on ddt.idord=ordineddt.idord WHERE ordineddt.iduser="&iduser&" order by ordineddt.data desc"
			'sql="select ddt.IDddt, ddt.Nddt, ddt.data, ddt.causale, ordini.Nord, ordini.idord, fatture.IDFat, fatture.Nfat FROM ((ordini INNER JOIN ddt ON ordini.idord = ddt.idord) LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDfat WHERE ddt.iduser="&iduser&" order by ddt.data desc"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				Set ccausale_ddt = new cl_causale_ddt
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Vedi</td><td align=""center"">N&deg; ddt</td><td align=""center"">Data</td><td align=""center"">Causale</td></tr>"&vbcrlf

				For i = LBound(arrRS, 2) To UBound(arrRS, 2)


					verde=ccausale_ddt.verde(arrRS(3, i))
					sommadiimporto=arrRS(11, i)
					if isnull(sommadiimporto) then
						sommadiimporto=0
					else
						sommadiimporto=cdbl(sommadiimporto)
					end if
					totale=cdbl(arrRS(10, i))
					if verde=false and arrRS(8, i) then	'reso tutto
						colore="bgcolor='#CCFF99'"	'Verdino chiaro reso tutto
					elseif verde then
						colore="bgcolor='#CCFF99'"	'Verdino chiaro reso tutto
					'elseif objPagingRS("causale")>1 then
					'	colore=""
					elseif totale=0 and arrRS(12, i)>5 then
						colore="bgcolor='#99FF99'" 'Verdino saldata
					elseif totale<0 then
						colore="bgcolor='#99FF99'"
					elseif isnull(sommadiimporto) then
						colore=""
					elseif totale=sommadiimporto then
						colore="bgcolor='#99FF99'" 'Verdino saldata
					elseif totale>sommadiimporto then
						colore="bgcolor='#FFCCCC'" 'Rosino da saldare
						testo_differenza="("&formatcurrency(sommadiimporto-totale,2)&")"
					else
						colore="bgcolor='#CC00CC'" 'Viola
						testo_differenza="(+"&formatcurrency(sommadiimporto-totale,2)&")"
					end if


					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center""><input type=""radio"" onClick=""location.href='pag_adm_ddt.asp?idddt="&arrRS(0, i)&"'""></td><td align=""center"">"&arrRS(1, i)&"</td><td align=""center"">"&formatdatetime(arrRS(2, i),2)&"</td><td align=""center"">"&ccausale_ddt.elemento(arrRS(3, i))&"</td></tr>"


				Next
				stringa=stringa&"</table>"
				Erase arrRS
				set ccausale_ddt = Nothing
			End If
	'***************************************************************************************************************************************************************************************
		case 5	'Preventivi
	'***************************************************************************************************************************************************************************************
			sql="select idord, nord, data, stato, totale FROM preventivi WHERE eliminato=0 and iduser="&iduser&" order by data desc"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Vedi</td><td align=""center"">N&deg; preventivo</td><td align=""center"">Data</td><td align=""center"">Causale</td><td align=""center"">Totale</td></tr>"&vbcrlf

				For i = LBound(arrRS, 2) To UBound(arrRS, 2)

					stringa=stringa&"<tr style=""border-bottom:1px solid;""><td align=""center""><input type=""radio"" onClick=""location.href='pag_adm_preventivi.asp?idord="&arrRS(0, i)&"'""></td><td align=""center"">"&arrRS(1, i)&"</td><td align=""center"">"&formatdatetime(arrRS(2, i),2)&"</td><td align=""center"">"&stato_preventivo(arrRS(3, i))&"</td><td align=""center"">"&formatnumber(arrRS(4, i),2)&"</td></tr>"


				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If
	'***************************************************************************************************************************************************************************************
		case 6	'Carrello
	'***************************************************************************************************************************************************************************************

			sql="select idcar, idpro,codice, articolo, variante1, variante_a, variante2, variante_b, prezzo, promozione, sconto, prezzo_ve_va, quantita FROM carrello_esteso where iduser="&iduser
			sql="select idcar, carrello.idpro, codice, articolo, variante1, variante_a, variante2, variante_b, prezzo, promozione, sconto, prezzo_ve_va, quantita FRom varianti_b RIGHT JOIN (varianti_a RIGHT JOIN (carrello INNER JOIN prodotti oN carrello.idpro = prodotti.IDpro) oN varianti_a.IDvara = carrello.idvara) oN varianti_b.IDvarb = carrello.idvarb where iduser="&iduser
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Codice</td><td>Articolo</td><td align=""right"">Prezzo<br>Unitario</td><td align=""center"">Quantit&agrave;</td><td align=""right"">Prezzo<br>Totale</td></tr>"&vbcrlf
				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					stringa_tmp="<tr style=""border-bottom:1px solid;""><td align=""center""><a href=""product.asp?idpro="&arrRS(1, i)&"""><b>"&arrRS(2, i)&"</b></a></td><td>"&arrRS(3, i)
					varianti=""
					if arrRS(5, i)<>"" then varianti=arrRS(4, i)&": "&arrRS(5, i)
					if arrRS(7, i)<>"" then
						if varianti<>"" then varianti=varianti&", "
						varianti=varianti&arrRS(6, i)&": "&arrRS(7, i)
					end if
					if varianti<>"" then stringa_tmp=stringa_tmp&"<br>"&varianti

					if arrRS(11, i)<>"" then
						prezzo=arrRS(11, i)
					else
						prezzo=arrRS(8, i)
					end if


					if arrRS(9, i)=true then
						prezzoeff=prezzo*(1-arrRS(10, i)/100)
						txtprezzo="<s>"&simbolo_valuta&formatnumber(prezzo,2,true)&"</s><br><b>" & simbolo_valuta&formatnumber(prezzoeff,2,true)&"</b>"
					else
						prezzoeff=prezzo
						if isnull(prezzoeff) or prezzoeff="" then
							txtprezzo="NULL"
						else
							txtprezzo=formatnumber(prezzoeff,2)
						end if
					end if

					stringa_tmp=stringa_tmp&"</td><td align=""right"">"&txtprezzo&"</td><td align=""center"">"&arrRS(12, i)&"</td><td align=""right"">"&totale_riga(prezzo,arrRS(10, i),arrRS(12, i))&"</td><tr>"
					stringa=stringa&stringa_tmp
				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If

	'***************************************************************************************************************************************************************************************
		case 7	'Articoli ordinati
	'***************************************************************************************************************************************************************************************

		anno=request("anno")
		if anno="" then
			anno=year(date())
		end if
		report=request.querystring("report")
		if report="" then
			report="articolo"
		end if
		%>

		Tipologia report
			<select name="tipo_report" id="tipo_report">
                <option value="ordine" <%if report="ordine" then response.write "selected"%>>Raggruppa per ordine</option>
				<option value="articolo" <%if report="articolo" then response.write "selected"%>>Raggruppa per articolo</option>

			</select>
               Anno
              <input name="anno" type="text" value="<%=anno%>" size="4" id="anno">
              <input type="submit" class="aggiorna" value="Aggiorna" style="FONT: 12px;" href="pag_adm_user.asp?tab=7&iduser=<%=iduser%>"></form></div>
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
	              <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella>n° ordini</td>
	              <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Totale</b></td>
	            </tr>
		<%if report="articolo" then
			sql_ordine="select ordini_dett.idpro, ordini_dett.var1,ordini_dett.var2, ordini_dett.variante1_ordine, ordini_dett.variante2_ordine, Sum(ordini_dett.quantita) AS quantita, Count(ordini.idord) AS ordini,ordini_dett.codice_ordine,ordini_dett.articolo_ordine FROM prodotti INNER JOIN (ordini INNER JOIN ordini_dett ON ordini.idord = ordini_dett.idord) ON prodotti.IDpro = ordini_dett.idpro WHERE (((ordini.iduser)="&iduser&") AND ((Year(ordini.data))="&anno&")) GROUP BY ordini_dett.codice_ordine, ordini_dett.idpro,ordini_dett.var1,ordini_dett.var2, ordini_dett.variante1_ordine, ordini_dett.variante2_ordine,ordini_dett.codice_ordine,ordini_dett.articolo_ordine HAVING (((ordini_dett.codice_ordine)<>''));"
			set rs=conn.execute(sql_ordine)
			do while not rs.eof
		%><tr valign="middle">
              <td align="center" style="border-bottom: 1px solid gray;">
              <b><a href="product.asp?idpro=<%=RS("idpro")%>" title="Clicca per la scheda prodotto" ><%=RS("codice_ordine")%></a></b>
                </td>
              <td style="border-bottom: 1px solid gray;"><%=rs("articolo_ordine")%><br>
              <%
					txt=""
					if rs("variante1_ordine")<>"" then
						txt ="&nbsp;" & rs("variante1_ordine") & ": " & rs("var1")
					end if
					if rs("variante2_ordine")<>"" then
						txt =txt & "&nbsp;" & rs("variante2_ordine") & ": " & rs("var2")
					end if
					response.write txt
			  %></td>
              <td align="center" style="border-bottom: 1px solid gray;"><%'=rs("um")%></td>
              <td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%=rs("quantita")%></td>
              <td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%
			  	'prezzoeff=rs("prezzo")
            	response.write formatcurrency(prezzoeff,2)
				'prezzoeff=prezzoeff-(rs("sconto_prodotto")/100)*prezzoeff
				'totale_riga=prezzoeff*rs("quantita")
				%></td>
              <td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%=rs("ordini")%></td>
              <td align="right" valign="middle" style="border-bottom: 1px solid gray;"><%'=formatcurrency(totale_riga,2)%></td>
            </tr>
		<%
			rs.movenext
		loop
		end if%>
            <%
			if report="ordine" then
			sconto_totale=0
			totord=0
			sql_ordine="select ordini.* FROM ordini where eliminato=0 and iduser="&iduser&" and year(data)="&anno&" order by idord desc"
			set rs_ordini=conn.execute(sql_ordine)
			do while not rs_ordini.eof
			%>
                     <tr valign="middle">
              <td colspan="7" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><strong><a href="pag_adm_ordini.asp?idord=<%=rs_ordini("idord")%>">Ordine: <%=rs_ordini("nord")%></a></strong> del <%=formatdatetime(rs_ordini("data"),2)%></td>
            </tr>
		<%
		sql_dettaglio="select ordini_dett_esteso.* FROM ordini_dett_esteso  where idord="&rs_ordini("idord")&" order by iddett;"
		set rs=conn.execute(sql_dettaglio)
		  do while not rs.eof
			%>

            <tr valign="middle">
              <td align="center" style="border-bottom: 1px solid gray;">
              <b><a href="product.asp?idpro=<%=RS("idpro")%>" title="Clicca per la scheda prodotto" ><%=RS("codice_ordine")%></a></b>
                </td>
              <td style="border-bottom: 1px solid gray;"><%=rs("articolo_ordine")%><br>
              <%
					txt=""
					if rs("var1")<>"" then
						txt ="&nbsp;" & rs("variante1_ordine") & ": " & rs("var1")
					end if
					if rs("var2")<>"" then
						txt =txt & "&nbsp;" & rs("variante2_ordine") & ": " & rs("var2")
					end if
					response.write txt
			  %></td>
              <td align="center" style="border-bottom: 1px solid gray;"><%=rs("um")%></td>
              <td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%=rs("quantita")%></td>
              <td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%
			  	prezzoeff=rs("prezzo")
            	response.write formatcurrency(prezzoeff,2)
				prezzoeff=prezzoeff-(rs("sconto_prodotto")/100)*prezzoeff
				%></td>
              <td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%=FormatNumber(rs("sconto_prodotto"),2)%></td>
              <td align="right" valign="middle" style="border-bottom: 1px solid gray;"><%=formatcurrency((prezzoeff*rs("quantita")),2)%></td>
            </tr>
            <%
			'totord=totord+prezzoeff*rs("quantita")
						totord=totord+round(prezzoeff*rs("quantita"),2)
            rs.movenext
			loop%>
            <%
			  rs_ordini.movenext
			  loop
			totord=round(totord,2)
			end if

			  %>
    </table>
    <%
	'***************************************************************************************************************************************************************************************
		case 8	'Attività sul sito
	'***************************************************************************************************************************************************************************************

	%>
	<table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1">
	<%
	strsql="select log_action.data, log_action.idloged, log_action.testo, log_action.idpro, prodotti.codice, prodotti.articolo,  settori.Nome_Settore, log_action.ip FROM (prodotti RIGHT JOIN (log_action LEFT JOIN utenti ON log_action.iduse = utenti.iduser) ON prodotti.IDpro = log_action.idpro) LEFT JOIN settori ON log_action.idpro = settori.idsettore   where log_action.iduse="& iduser &" ORDER BY log_action.data DESC"
	call paginazione_start(strsql,50,"access")
	%>
		<tr bgcolor="#E5E5E5">
            <td><p><b>&nbsp;Data</b></p></td>
            <td>IP</td>
            <td><strong>Azione</strong></td>
            <td><b>&nbsp;</b></td>
		</tr>
		<%
		Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
			testo1=""
			select case objPagingRS("Testo")
			case "ric"
				testo="Cercato parole:" & id
			case "pro"
				testo="Visualizzato prodotto:"
				testo1= objPagingRS("codice")&" " &objPagingRS("articolo")
			case "cate"
				testo="Selezionato CATEGORIA "
				testo1= objPagingRS("categoria")
			case "reg"
				testo="Registrato l'utente "
				testo1="<a href='pag_adm_user.asp?iduser=" & objPagingRS("idpro")&"'>"&objPagingRS("nome")&" "&objPagingRS("cognome")&"</a>"
			case "log"
				testo="Ha effettuato il login"
			case "flog"
				testo="Login automatico"
			case "new"
				testo="Nuovo"
			case "rit"
				testo="Ritornato "
			case "wap"
				testo="Accesso al Wap"
			case "logwap"
				testo="Login Wap al Wap"
			case "pagl"
				testo="Accesso area riservata"
			case "sett"
				testo="Selezionato SETTORE "
				testo1= objPagingRS("nome_settore")
			case else
				testo=objPagingRS("Testo")
			end select
			%>
          <tr>
            <td nowrap class=corpobox><%=objPagingRS("data")%></td>
            <td class=corpobox>
	            <%= objPagingRS("ip")%>

	            </td>
            <td class=corpobox><%=testo%></td>
            <td class=corpobox><%=testo1%></td>
          </tr>
          <%
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	call paginazione_end(4,false)
	%>
        </table>
<%
		'***************************************************************************************************************************************************************************************
		case 9 'scadenze
	'***************************************************************************************************************************************************************************************
	sql="select scadenze.data, scadenze.importo, scadenze.note_pagamento, scadenze.tipo_pagamento, scadenze.riba FROM fatture INNER JOIN scadenze ON fatture.IDFat = scadenze.idfat WHERE fatture.iduser="&iduser & " order by scadenze.data"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Data</td><td align=""center"">Descrizione</td><td align=""right"">Importo</td></tr>"&vbcrlf

				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					if arrRS(0, i)<date() then
						colore="bgcolor='#FFCCCC'"
					else
						colore=""
					end if
					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center"">"&formatdatetime(arrRS(0, i),2)&"</td><td align=""center"">"&arrRS(2, i)&"</td><td align=""right"">"&simbolo_valuta&" "&formatnumber(arrRS(1, i),2)&"</td></tr>"


				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If


	'***************************************************************************************************************************************************************************************
		case 11 'Registro
	'***************************************************************************************************************************************************************************************
			if mysql_server="" then
				sql="select * from log where instr(evento,'[utente="& iduser&"]')>0 orde by data desc"
				'set rs=conn.execute (sql)
			else
				'MYSQL
				'Set conn = Server.CreateObject("ADODB.Connection")
				'conn.Open "server="&mysql_server&";uid="&mysql_uid&";pwd="&mysql_pwd&";database="&mysql_database&";driver=MySQL ODBC 3.51 Driver"
				sql="select * from log where instr(evento,'[utente="& iduser&"]')>0 order by data desc"
				'set rs=conn.execute (sql)
			end if
	%>
      <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
        <tr>
          <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;Data</td>
          <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;</td>
          <td align="left" bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;Evento&nbsp;</td>
        </tr>
        <%
		call paginazione_start(sql,30,"mysql")
		Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
%>
        <tr>
          <td style="border-bottom:1px solid;"><%=objPagingRS("data")%></td>
          <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%
		  	testo=objPagingRS("causale")
		  select case testo
		  case 1
			  txt="sem_verde_t.gif"
		  case 2
			  txt="sem_giallo_t.gif"
		  case 3
		  	txt="sem_rosso_t.gif"
		  case else
		  	txt="sem_rosso_t.gif"
		  end select

		  %>
            <img src="/images/<%=txt%>" width="20" height="20" /></td>
          <td style="border-bottom:1px solid;"><%txt=objPagingRS("evento")
response.write FormatText(left(txt,instr(txt,vbcrlf)))
		  'response.write left(txt,instr(txt,vbcrlf))%></td>
        </tr>
        <%
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	call paginazione_end(3,false)
	%>
      </table>
			<%
	'***************************************************************************************************************************************************************************************
		case 12	'fatture Fornitore
	'***************************************************************************************************************************************************************************************
			sql="select fatture_for.IDFat, fatture_for.Nfat, fatture_for.data, fatture_for.totale_fattura, sommadiincassi.SommaDiimporto FROM fatture_for LEFT JOIN (select scadenze_for.idfat, Sum(scadenze_for.importo) AS SommaDiimporto FROM scadenze_for WHERE (((scadenze_for.Pagato)=True)) GROUP BY scadenze_for.idfat  )  AS sommadiincassi ON fatture_for.IDFat = sommadiincassi.idfat WHERE fatture_for.idfor="&iduser&" order by fatture_for.data desc"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Vedi</td><td align=""center"">N&deg; fattura</td><td align=""center"">Data</td><td align=""center"">Totale</td><td align=""center"">Da saldare</td></tr>"&vbcrlf

				For i = LBound(arrRS, 2) To UBound(arrRS, 2)

					testo_differenza=""
					sommadiincassi=arrRS(4, i)
					if isnull(sommadiincassi) then
						sommadiincassi=0
					else
						sommadiincassi=cdbl(sommadiincassi)
					end if
					if cdbl(arrRS(4, i))<0 then	'totale<0
						colore="bgcolor='#99FF99'"
					'elseif isnull(arrRS(4, i)) then 'isnull(sommadiimporto)

						'colore=""
					elseif cdbl(arrRS(3, i))=sommadiincassi then 'totale=sommadiimporto
						colore="bgcolor='#99FF99'" 'Verdino saldata
					elseif cdbl(arrRS(3, i))>sommadiincassi then 'totale>sommadiimporto
						colore="bgcolor='#FFCCCC'" 'Rosino da saldare
						testo_differenza="("&formatcurrency(arrRS(3, i)-sommadiincassi,2)&")"
					else
						colore="bgcolor='#CC00CC'" 'Viola
						testo_differenza="(+"&formatcurrency(arrRS(3, i)-sommadiincassi,2)&")"
					end if
					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center""><input type=""radio"" onClick=""location.href='pag_adm_fatture_for.asp?idfat="&arrRS(0, i)&"'""></td><td align=""center"">"&arrRS(1, i)&"</td><td align=""center"">"&formatdatetime(arrRS(2, i),2)&"</td><td align=""center"">"&formatnumber(arrRS(3, i),2)&"</td><td align=""center"">"&formatnumber(cdbl(arrRS(3, i))-sommadiincassi,2)&"</td></tr>"
				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If
			'***************************************************************************************************************************************************************************************
		case 13 'scadenze fornitore
	'***************************************************************************************************************************************************************************************
	sql="select scadenze_for.scadenza, fatture_for.nfat, scadenze_for.importo  FROM fatture_for INNER JOIN scadenze_for ON fatture_for.IDFat = scadenze_for.idfat WHERE (((scadenze_for.Pagato)=False) AND ((fatture_for.idfor)="&iduser & ")) ORDER BY scadenze_for.scadenza;"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Data</td><td align=""center"">Fattura</td><td align=""right"">Importo</td></tr>"&vbcrlf
				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					if arrRS(0, i)<date() then
						colore="bgcolor='#FFCCCC'"
					else
						colore=""
					end if
					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center"">"&formatdatetime(arrRS(0, i),2)&"</td><td align=""center"">"&arrRS(1, i)&"</td><td align=""right"">"&simbolo_valuta&" "&formatnumber(arrRS(2, i),2)&"</td></tr>"
				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If

		'***************************************************************************************************************************************************************************************
		case 14	'Articoli forniti
	'***************************************************************************************************************************************************************************************
			sql="select idpro,codice, articolo,codice_fornitore FROM prodotti where idfor="&iduser&" order by codice"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Codice</td><td align=""center"">Codice fornitore</td><td>Articolo</td></tr>"&vbcrlf
				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					stringa_tmp="<tr style=""border-bottom:1px solid;""><td align=""center""><a href=""product.asp?idpro="&arrRS(0, i)&"""><b>"&arrRS(1, i)&"</b></a></td><td>"&arrRS(3, i)&"</td><td>"&arrRS(2, i)&"</td><tr>"
					stringa=stringa&stringa_tmp
				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If
'***************************************************************************************************************************************************************************************
		case 15	'ordini fornitore
	'***************************************************************************************************************************************************************************************

			sql="select ordini_fornitori.idord, ordini_fornitori.nord, ordini_fornitori.data, ordini_fornitori.stato, ordini_fornitori.totale from ordini_fornitori where ordini_fornitori.iduser="&iduser&" order by ordini_fornitori.data desc"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Vedi</td><td align=""center"">N&deg; ordine</td><td align=""center"">Data</td><td align=""center"">stato</td><td align=""center"">Totale</td></tr>"&vbcrlf
				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center""><input type=""radio"" onClick=""location.href='pag_adm_ordini_fornitori.asp?idord="&arrRS(0, i)&"'""></td><td align=""center"">"&arrRS(1, i)&"</td><td align=""center"">"&formatdatetime(arrRS(2, i),2)&"</td><td align=""center"">"&stato_ordine_fornitore(arrRS(3, i))&"</td><td align=""center"">"&formatnumber(arrRS(4, i),2)&"</td></tr>"
				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If
'***************************************************************************************************************************************************************************************
		case 16	'Dipendenti
	'***************************************************************************************************************************************************************************************

		action=request.form("action")
		iddip=request("iddip")
		iduser=request("iduser")
		if oper="delete" then

				set rs_dip=conn.execute ("select * from dipendenti where iddip="&iddip)
				nominativo_dip=rs_dip("cognome")&" "&rs_dip("nome")
				'iduser=rs_dip("iduser")
				set rs_dip = Nothing
				conn.execute("delete from dipendenti_misure where iddip="&iddip)
				sql="delete from dipendenti where iddip="&iddip
				conn.execute (sql)
				sql="delete from utenti_dipendenti where iddip="&iddip
				conn.execute (sql)
				call add2log("Eliminato dipendente "&nominativo_dip&" di "&utente_log(iduser),2)
				'call add2log("Eliminato dipendente "&nominativo_dip,2)
				action=""
				oper="list"
		end if
		if oper="dimetti" then
				set rs_dip=conn.execute ("select * from dipendenti where iddip="&iddip)
				nominativo_dip=rs_dip("cognome")&" "&rs_dip("nome")
				set rs_dip = Nothing
				sql="delete from utenti_dipendenti where iddip="&iddip&" and iduser="&iduser
				conn.execute (sql)
				call add2log("Dimesso dipendente "&nominativo_dip&" da "&utente_log(iduser),2)
				action=""
				oper="list"
		end if
		if oper="aggiungiesistente" then
				set rs_dip=conn.execute ("select * from dipendenti where iddip="&iddip)
				nominativo_dip=rs_dip("cognome")&" "&rs_dip("nome")
				set rs_dip = Nothing
				sql="insert into utenti_dipendenti (iddip, iduser) values ("&iddip&","&iduser&")"
				response.write sql
				conn.execute (sql)
				call add2log("Aggiunto dipendente esistente "&nominativo_dip&" da "&utente_log(iduser),2)
				action=""
				oper="list"
		end if



		if action<>"" then
			iduser=request("iduser")
			sql="select dipendenti.* from dipendenti"
			if action="update" then
				sql=sql&" where iddip="&iddip
				txt_oper="Modificato "
			end if
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql, conn, 3, 3
			if action="add" then
				txt_oper="Aggiunto "
				rs.addnew
				'rs("iduser")=iduser
				'setta spettanze in ordini
				setta_spettanze_in_ordini(iduser)
			end if
			if action="update" then
				set modificheRS= (new ClasseModificheRS)(oper)
				modifichers.leggi(rs)
			end if
			rs("nuovo")=0
			rs("cognome")=ucase(request.form("cognome"))
			rs("nome")=ucase(request.form("nome"))
			rs("matricola")=request.form("matricola")
			rs("Contatto")=request.form("Contatto")
			rs("sesso")=request.form("sesso")
			rs("grado")=request.form("grado")
			rs("arma")=request.form("arma")
			rs("lato_arma")=request.form("lato_arma")&""
			rs("nota_dipendente")=request.form("nota_dipendente")
			lingue=request.form("lingue")
			if request.form("lingue-altro")<>"" then lingue=lingue&","&request.form("lingue-altro")
			rs("lingue")=lingue
			if action="update" then
				modifiche_rs=modificheRS.confronta(rs)
				set modificheRS=Nothing
			end if

			nominativo_v=rs("cognome")&" "&rs("nome")
			rs.update
			if action="add" then
				iddip=Get_last_id("dipendenti")
				conn.execute ("insert into utenti_dipendenti (iddip,iduser) values ("&iddip&","&iduser&")")
			end if
			if action="update" then
				modifiche_rs=modifiche_rs&aggiorna_misure_in_ordini(iddip)
			end if
			rs.close
			set rs=nothing
			call add2log(txt_oper&" dipendente "&nominativo_v&" di "&utente_log(iduser)&vbcrlf&modifiche_rs,2)
			oper="view"
			aggiorna_elenco_dipendenti=true

		end if
	'***************************************************************************************************************************************************************************************
			'Dipendenti - LIST
	'***************************************************************************************************************************************************************************************
			'response.write "OPER:"&oper
			dialog=request.querystring("dialog")
			if oper="" then oper="list"
			if oper="list" then
				secondario=clng(conn.execute("select secondario from utenti where iduser="&iduser)(0))

				if secondario>0 then
					iduser=secondario
				end if


				sql="select dipendenti.*, misure.data_rilievo, utenti_gradi.nome_grado, utenti_gradi.colore, utenti_gradi.colore_nominativo,utenti_dipendenti.dimesso FROM (dipendenti left JOIN utenti_gradi ON dipendenti.grado = utenti_gradi.id) left join (SELECT iddip,max(data_rilievo) as data_rilievo FROM dipendenti_misure group by iddip  ) as misure ON dipendenti.iddip = misure.iddip inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.iduser="&iduser

				iPageSize=1000
				ordine=" order by sesso desc, cognome,nome"
			 	strSql = sql &where&ordine
				PaginazioneString="&cercain="&cercain&"&cerca="&cerca&"&anno="&anno&"&cerca_hidden="&request("cerca_hidden")
				'response.write strSql

				%>



		          <table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" class="tabella1" id="elenco_dipendenti">
		            <tr>
		              <td colspan="6">
		                <span style="float: right;"><a href="pdf_taglie.asp?iduser=<%=iduser%>">PDF taglie</a>&nbsp; <a href="pag_adm_rep23.asp?iduser=<%=iduser%>" target="_blank">Elenco taglie</a>
		                <%if ha_il_permesso("E3") then %>
		                <input type="button" value="Aggiungi nuovo" class="apri-in-dialog" data-id="<%=iduser%>" data-oper="new" data-nomeid="iduser" data-url="<%=questofile%>?tab=16" data-titolo="Aggiungi dipendente">
		                <input type="button" value="Aggiungi esistente" class="apri-in-dialog" data-id="<%=iduser%>" data-oper="newbis" data-nomeid="iduser" data-url="<%=questofile%>?tab=16" data-titolo="Aggiungi esistente">
		                <%end if %>
		                </span>
		                </td>
		            </tr>
		            <tr>
		              <td align="center" bgcolor="#E5E5E5"><p> Vedi</p></td>
		              <td align="center" bgcolor="#E5E5E5">Nominativo</td>
		              <td align="center" bgcolor="#E5E5E5">Matricola</td>
		              <td align="left" bgcolor="#E5E5E5">Grado</td>
		              <td bgcolor="#E5E5E5">Contatto</td>
		              <td bgcolor="#E5E5E5">Ultimo rilievo</td>
		            </tr>
		            <%
		call paginazione_start(strsql,ipagesize,"access")

		Do While iRecordsShown < iPageSize And Not objPagingRS.EOF

		%>            <tr class="coprobox" <%=colore%>>
		              <td align="center" style="border-bottom:1px solid;">
		                <%if ha_il_permesso("E3") then %>
			              <input type="radio" name="modifica" value="<%=objPagingRS("iddip")%>" class="apri-in-dialog" data-id="<%=objPagingRS("iddip")%>" data-oper="view" data-nomeid="iddip" data-url="<%=questofile%>?tab=16" data-titolo="Dipendente">
			            <%end if %>
			              </td>
		              <td style="border-bottom:1px solid;"><%=dipendente_colorato(  objPagingRS("cognome")&" "&objPagingRS("nome"),objPagingRS("sesso"), objPagingRS("colore_nominativo") )%><%if objPagingRS("nuovo")=1 then response.write balloon("00ff00","nuovo") %><%if objPagingRS("dimesso")=1 then response.write balloon("FF0000","dimesso") %> </td>
		              <td style="border-bottom:1px solid;" align="center"><%=objPagingRS("matricola")%></td>
		              <td style="border-bottom:1px solid;"><%=grado_colorato(objPagingRS("nome_grado"),objPagingRS("colore"))%></td>
		              <td style="border-bottom:1px solid;"><%=objPagingRS("contatto")%></td>
		              <td style="border-bottom:1px solid;"><%=objPagingRS("data_rilievo")%></td>
		            </tr>
		            <%
				iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
				objPagingRS.MoveNext
			Loop
		    call paginazione_end(4,true)

		%>
		          </table>
					<%
					elseif oper="newbis" then
					%>
					<form id="form-dip" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;">
						<input type="hidden" name="iduser" value="<%=iduser%>">
					Aggiungi dipendente: <input type="text" id="dipendente-esistente" name="iddip" value="" style="width: 300px;">
					<input id="aggiungi_dipendente" name="salva_dipendente" type="button"  style="FONT: 12px;" value="Aggiungi" class="" data-id="<%=iduser%>" data-url="<%=questofile%>?tab=16" data-nomeid="iduser" data-oper="aggiungiesistente" data-form="form-dip">

					</form>
					<script>

						$(function() {

							$('#dipendente-esistente').select2({
								dropdownCssClass: 'ui-dialog',
								placeholder: 'Seleziona dipendente esistente',
								minimumInputLength: 0,
								allowClear: true,
								ajax: {
									quietMillis: 150,
									url: "ajax_function.asp?select2=tuttidipendenti",
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


			$("#aggiungi_dipendente").click(function(e){
				e.preventDefault();

				$.ajax({
					url     : "ajax_function.asp?oper=aggiungi_dipendente_esistente",
					type    : "post",
					cache	: false,
					dataType: 'json',
					data	: $("#form-dip").serialize(),
					success: function(data){


						$("#dialog").dialog("close");
						if (typeof aggiorna_tabella_dipendenti === "function") {
							aggiorna_tabella_dipendenti();
						}




					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});


						});


						</script>



					<%





					else






	'***************************************************************************************************************************************************************************************
			'Dipendenti - VIEW e EDIT
	'***************************************************************************************************************************************************************************************
			if iddip="" then
				iddip=request.querystring("iddip")
			end if
			if oper="view" or oper="edit" then
				'response.write idfat
				if request.form("iduser")<>"" then
					iduser=request.querystring("iduser")
				end if
				sql1="select dipendenti.iddip, dipendenti.dummy1, dipendenti.cognome, dipendenti.nome, dipendenti.contatto, dipendenti.sesso, dipendenti.grado, dipendenti.nota_dipendente , dipendenti.matricola, dipendenti.arma,dipendenti.lato_arma,   dipendenti.lingue, i.azienda, utenti_dipendenti.iduser FROM dipendenti INNER JOIN  utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip inner join utenti ON utenti_dipendenti.iduser = utenti.iduser inner join utenti_intestazioni i on utenti.idintestazione = i.id  WHERE dipendenti.iddip= " & iddip &";"
				set rs1=conn.execute( sql1)
			elseif oper="new" then
				iduser=request.querystring("iduser")
				iddip=0
				dim rs1(12)

			end if

				%>

        <form id="form-dip" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" onSubmit="<%=convalida%>">

	        <%if oper="view" and ha_il_permesso("E3") then%><span style="float: right;"><input name="edit" type="button" onclick="modifica_dipendente();" style="FONT: 12px;" value="Modifica" class="" data-id="<%=iddip%>" data-url="<%=questofile%>?tab=16" data-nomeid="iddip" data-oper="edit" data-form="form-dip"></span><%end if%><br>

		<%if oper="view" then %>
		<label>ID</label><%=rs1("iddip")%><br>
		<%end if%>


		<label>Cognome</label><%if oper="view" then%>
	                <%=rs1("cognome")%>
                <%else %>
	                <input name="cognome" type="text" value="<%=rs1(2)%>">
				<%end if%>
		<br>
		<label>Nome</label><%if oper="view" then%>
	                <%=rs1("nome")%>
                <%else %>
	                <input name="nome" type="text" value="<%=rs1(3)%>">
				<%end if%>
		<br>
		<label>Matricola</label><%if oper="view" then%>
	                <%=rs1("matricola")%>
                <%else %>
	                <input name="matricola" type="text" value="<%=rs1(8)%>">
				<%end if%>
		<br>
		<label>Contatto</label><%if oper="view" then%>
	                <%=rs1("Contatto")%>
                <%else %>
	                <input name="Contatto" type="text" value="<%=rs1(4)%>">
				<%end if%>
		<br>
		<label>Sesso</label><%if oper="view" then %>
	                <%=rs1("sesso")%>
                <%else %>
                 <%
					 val_si=""
					 val_no=""
					if rs1(5)="F" then val_F=" checked" else val_M=" checked"
 %>
        Maschio <input name="sesso" type="radio" value="M" <%=val_M%>> Femmina<input name="sesso" type="radio" value="F" <%=val_F%>>
              <%end if%>
		<br />
		<label>Grado</label><%if oper="view" then
			if rs1("grado")=0 then
				response.write "Non specificato"
			else
				set grado=conn.execute("select nome_grado from utenti_gradi where id="&rs1("grado"))
				if grado.eof then
					response.write "Non trovato ("&rs1("grado")&")"
				else
					response.write grado("nome_grado")
				end if
			end if
		%>
                <%else %>
        	<select name="grado">
	        <option value="0">Non specificato</option>
			<%=elenco_option_gradi("select id, nome_grado as testo, colore,ordine from utenti_gradi order by ordine",rs1(6))%>

			</select>
              <%end if%>
      		<br>
      		<%if oper="view" then
	      		if not isnull(rs1("lingue")) then
	      			lingue=split(rs1("lingue"),",")
	      		end if
      			disabled=" disabled"
      		elseif oper="edit" then
	      		if not isnull(rs1("lingue")) then
	      			lingue=split(rs1("lingue"),",")
      			end if
      		else
	      		lingue=""
	      	end if
      		 %>
      	<label>Bandierina:</label>
      	Francese <input type="checkbox" name="lingue" value="fr" <%=checked_lingua("fr",lingue)%> <%=disabled%>>&nbsp;
      	Inglese<input type="checkbox" name="lingue" value="en" <%=checked_lingua("en",lingue)%> <%=disabled%>> &nbsp;
      	Tedesco<input type="checkbox" name="lingue" value="de" <%=checked_lingua("de",lingue)%> <%=disabled%>> &nbsp;
      	Spagnolo<input type="checkbox" name="lingue" value="sp" <%=checked_lingua("sp",lingue)%> <%=disabled%>> &nbsp;
      	<%
	      if oper="view" or oper="edit" then
		      if IsArray(lingue) then
			    if UBound(lingue)>=0 then
					linguealtro=trim(lingue(ubound(lingue)))
					if len (linguealtro)=2 then
						linguealtro=""
					end if
				end if
		    end if
	      end if
	    %>
      	Altro:<input type="text" name="lingue-altro" value="<%=linguealtro%>" <%=disabled%>>

      	<br>

		<label>Arma</label><%if oper="view" then%>
	                <%=rs1("arma")%> <strong>Lato</strong> <%=ucase(rs1("lato_arma"))%>
                <%else
					lato_arma=trim(rs1(10))
				%>
	                <input name="arma" type="text" value="<%=rs1(9)%>"> <strong>Lato</strong>  DX <input type="radio" value="dx" name="lato_arma" <%if lato_arma="dx" then response.write " checked"%>> SX<input type="radio" value="sx" name="lato_arma" <%if lato_arma="sx" then response.write " checked"%>>
				<%end if%>
		<br />
		<label>Note</label>
		<%if oper="view" then%>
                <%=rs1("nota_dipendente")%>
                <%else


			if oper="edit" then
			val=rs1("nota_dipendente")
		else
	      	val=request.form("nota_dipendente")
		end if
		%>
		<br/>
		<textarea name="nota_dipendente" style="width:100%; height:100px;"><%=val%></textarea>
		<%end if
		if oper<>"view" then
			if oper="edit" then
				action="update"
				n_ordini=clng(conn.execute("select count(*) from ordini_dett_spettanze where iddip="&iddip)(0))
				n_utenti=clng(conn.execute("select count(*) from utenti_dipendenti where iddip="&iddip)(0))
				if n_ordini>0 or n_utenti>1 then
					disabled=" disabled"
				end if
			end if
			if oper="new" then action="add"
			%>
			<input type="hidden" name="action" value="<%=action%>">
			<input type="hidden" name="iddip" value="<%=iddip%>">
			<input type="hidden" name="iduser" value="<%=iduser%>">
			<input id="salva_dipendente" name="salva_dipendente" type="button"  style="FONT: 12px;" value="Salva" class="" data-id="<%=iddip%>" data-url="<%=questofile%>?tab=16" data-nomeid="iddip" data-oper="edit" data-form="form-dip">
			<%if oper="edit" then %>
			<input id="dimetti_dipendente" name="dimetti_dipendente" type="button"  style="FONT: 12px;" value="Dimetti dipendente" class="" data-id="<%=iddip%>" data-url="<%=questofile%>?tab=16" data-nomeid="iddip" data-oper="dimetti" data-form="form-dip">
			<input id="elimina_dipendente" name="elimina_dipendente" type="submit"  style="FONT: 12px;" value="Elimina dipendente" class="" data-id="<%=iddip%>" data-url="<%=questofile%>?tab=16" data-nomeid="iddip" data-oper="delete" data-form="form-dip" <%=disabled%>>
			<%end if %>
			<%
		end if
			response.write "<br>"
		if oper="view"  then
			if iduser<>"" then
				call tabella_accessori(iduser)
			end if

			call tabella_misure(iddip, dialog)%>


				<%end if %>

		<script>

		function carica_tabella_misure(){
			$.ajax({
				url     : "pag_adm_user.asp?tabella_misure=<%=iddip%>",
				type    : "post",
				cache	: false,
				//dataType: 'json',
				//data	: dati,
				success: function(data){
					$("#tabella_misure").html(data);
				}
				,error: function(xhr, textStatus, error){
					toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
					toastr.error('Errore nel caricamento della pagina');
					var txt="";
					txt+="Errore in "+location.href+"<br>";
					txt+="<br>querystring: ?"+querystring[1];
					txt+="<br>"+xhr.statusText;
					txt+="<br>"+xhr.responseText;
					txt+="<br>textStatus:"+textStatus;
					txt+="<br>error:"+error;

					$.ajax({
						url     : "searcher.asp",
						type    : "post",
						data	: "txt_errore="+encodeURIComponent(txt)
					});
				}
			});
		}


		function modifica_dipendente(){

				var iduserTxt="";
				if (typeof iduser !== 'undefined') {
					iduserTxt="&iduser="+iduser;
				    // the variable is defined
				}
				var url="pag_adm_user.asp?tab=16&oper=edit&iddip=<%=iddip%>"+iduserTxt;
				console.log(url)
				$.ajax({
					url     : url,
					type    : "post",
					cache: false,
					//dataType: 'json',
					//data	: dati,
					success: function(data){
						$("#dialog").dialog().html(data);
						$("#dialog").dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 800,
							height: "auto",
							title: "Modifica dipendente ",
							position: {my: "center", at: "center", of: window}
						});
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});





		}




		$(function() {







		$("#salva_misura").click(function(e){
			e.preventDefault();
			form.validate();
			if (! form.valid()){
				return false;
			}

			var dati="oper=salva_misura_dipendente&"+$( "#form_misura" ).serialize();
			$.ajax({
				url     : "ajax_function.asp",
				type    : "post",
				cache	: false,
				dataType: 'json',
				data	: dati,

				success: function(data){

					if (data.success){


					console.log ("Chiamo carica_tabella_misure");
					if (dialog==""){
						carica_tabella_misure();
					}else
					{
						var dati = {};

						aggiorna_tabella(dati);
					}
					$("#dialog").dialog("close");
					}else{
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error("Errore nell'aggiornamento dei dati: "+ data.message);
					}

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
					invia_errore("Errore in salva_misura()",txt );
				}
			});

		});
			$("#dimetti_dipendente").click(function(e){
				e.preventDefault();
				var iduser=$("#iduser").val();
				var iddip=<%=iddip%>;
				console.log("dimetti dipendente iddip:"+iddip+" iduser:"+iduser)

				$.ajax({
					url     : "ajax_function.asp?oper=dimetti_dipendente",
					type    : "post",
					cache	: false,
					dataType: 'json',
					data	: {iduser:iduser, iddip: iddip},
					success: function(data){
						$("#dialog").dialog("close");
						if (typeof aggiorna_tabella_dipendenti === "function") {
							aggiorna_tabella_dipendenti();
						}



					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>querystring: ?"+querystring[1];
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});
			$("#elimina_dipendente").click(function(e){
				e.preventDefault();
				var iduser=$("#iduser").val();
				var iddip=<%=iddip%>;
				console.log("elimina dipendente iddip:"+iddip+" iduser:"+iduser)

				$.ajax({
					url     : "ajax_function.asp?oper=elimina_dipendente",
					type    : "post",
					cache	: false,
					dataType: 'json',
					data	: {iduser:iduser, iddip: iddip},
					success: function(data){
						$("#dialog").dialog("close");
						if (typeof aggiorna_tabella_dipendenti === "function") {
							aggiorna_tabella_dipendenti();
						}



					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>querystring: ?"+querystring[1];
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});



			$("#salva_dipendente").click(function(e){
				e.preventDefault();

				$.ajax({
					url     : "ajax_function.asp?oper=salva_dipendente",
					type    : "post",
					cache	: false,
					dataType: 'json',
					data	: $("#form-dip").serialize(),
					success: function(data){

						if (typeof aggiorna_tabella_dipendenti === "function") {
							aggiorna_tabella_dipendenti();
						}
						if (typeof aggiorna_tabella === "function") {

							var dati = {};
							//aggiorno ordine
							aggiorna_tabella(dati);
						}

						var action=$("#form-dip input[name=action]").val();
						console.log("action:"+action);
						if (action=='add'){
							var iddip=data.iddip;

							//Carico il dipendente in view
						var url="pag_adm_user.asp?tab=16&iddip="+iddip+"&iduser=<%=iduser%>";
				console.log(url)
				$.ajax({
					url     : url,
					type    : "post",
					cache: false,
					//dataType: 'json',
					//data	: dati,
					success: function(data){
						$("#dialog").dialog().html(data);
						$("#dialog").dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 800,
							height: "auto",
							title: "Scheda dipendente ",
							position: {my: "center", at: "center", of: window}
						});
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});











							$("#dialog").html(data);

						}else{
							console.log("chiudi dialog")
							$("#dialog").dialog("close");
						}



					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>querystring: ?"+querystring[1];
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});
			$(document).on("click",".radio_edit",function(e){
				var dialog=$("#dialog").dialog();
				dialog.html("Attendi...");
				var id=$(this).attr("id").replace("misura", "");
				console.log("ID:"+id);
				$.ajaxSetup({ cache: false });
				$.ajax({
					url     : "pag_adm_user.asp?modifica_misura="+id,
					type    : "post",
					//dataType: 'json',
					//data	: dati,
					success: function(data){
						dialog.html(data);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: "auto",
							height: "auto",
							title: "Modifica misura "
							,
							position: {my: "center", at: "center", of: window}
						});
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>querystring: ?"+querystring[1];
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});
		});
		</script>






		  </div>
			</div>
			<%
			%>
			          </form>
<%


					end if

	'***************************************************************************************************************************************************************************************
		case 17	'GRADI
	'***************************************************************************************************************************************************************************************
			sql="select id,nome_grado FROM utenti_gradi where iduser="&iduser&" order by nome_grado"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
				"<tr><td colspan=""2""><span style=""float: right;""><input type=""button"" value=""Nuovo"" id=""aggiungi_grado"" data-cosa=""grado""></span></td></tr>"&_
				"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td></td><td align=""center"">Grado</td></tr>"&vbcrlf
			If IsArray(arrRS) Then
				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					stringa_tmp="<tr style=""border-bottom:1px solid;""><td align=""center""><input type=""radio"" class=""edit_in_dialog"" data-cosa=""grado"" value=""grado"&arrRS(0, i)&"""></td><td>"&arrRS(1, i)&"</td><tr>"
					stringa=stringa&stringa_tmp
				Next
				Erase arrRS
			End If
			stringa=stringa&"</table>"

	'***************************************************************************************************************************************************************************************
		case 18	'ARTICOLI PER SPETTANZE
	'***************************************************************************************************************************************************************************************




			sql="select prodotti_spettanze.id, prodotti.codice, prodotti.articolo, varianti_a.variante_a, varianti_b.variante_b, prodotti_spettanze.ordine from ((prodotti_spettanze INNER JOIN prodotti on prodotti_spettanze.idpro=prodotti.idpro) left join varianti_a on prodotti_spettanze.idvara = varianti_a.idvara ) left join varianti_b on prodotti_spettanze.idvarb = varianti_b.idvarb where iduser="&iduser&" order by prodotti.codice"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Modifica</td><td align=""center"">Codice</td><td align=""center"">Articolo</td><td align=""center"">Varianti</td><td>Ordine</td></tr>"&vbcrlf
			If IsArray(arrRS) Then
				For i = LBound(arrRS, 2) To UBound(arrRS, 2)


					stringa=stringa&"<tr style=""border-bottom:1px solid;"" ><td align=""center""><input type=""radio"" class=""edit_in_dialog"" data-cosa=""articolo_spettanze"" value=""articolo_spettanze"&arrRS(0, i)&"""></td><td align=""center""><b>"&arrRS(1, i)&"</b></td><td >"&arrRS(2, i)&"</td><td align=""center"">"&arrRS(3, i)&" "&arrRS(4, i)&"</td><td align=""right"">"&arrRS(5, i)&"</td></tr>"


				Next
				Erase arrRS
			End If
			stringa=stringa&"<tr><td colspan=""5"">"

			stringa=stringa&"<input type=""button"" value=""Aggiungi articolo"" id=""dialog_aggiungi_articolo"" data-cosa=""articolo_spettanze""></td></tr>"
			stringa=stringa&"</table>"
			%>
		<script>

		$(function() {
			$("#dialog_aggiungi_articolo").click(function(e){
				var dialog=$("#dialog").dialog();
				dialog.html("Attendi...");
				var cosa=$(this).attr("data-cosa");
				console.log ("dialog_aggiungi_articolo cosa:"+cosa);
				$.ajax({
					url     : "pag_adm_user.asp?dialog=aggiungi_"+cosa+"&iduser=<%=iduser%>",
					type    : "post",
					cache	: false,
					//dataType: 'json',
					//data	: dati,
					success: function(data){
						dialog.html(data);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 600,
							height: "auto",
							title: "Aggiungi "+cosa
						});
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>querystring: ?"+querystring[1];
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});
			})
			</script>

			<%

	'***************************************************************************************************************************************************************************************
		case 19	'ordini_fornitori
	'***************************************************************************************************************************************************************************************

			sql="select ordini_fornitori.idord, ordini_fornitori.nord, ordini_fornitori.data, ordini_fornitori.stato, ordini_fornitori.totale FROM ordini_fornitori where ordini_fornitori.eliminato=0 and ordini_fornitori.iduser="&iduser&" order by ordini_fornitori.data desc"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Vedi</td><td align=""center"">N&deg; ordine</td><td align=""center"">Data</td><td align=""center"">stato</td><td align=""center"">Totale</td></tr>"&vbcrlf

				For i = LBound(arrRS, 2) To UBound(arrRS, 2)

					testo_differenza=""
					totale_ordine=cdbl(arrRS(4, i))
					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center""><input type=""radio"" onClick=""location.href='pag_adm_ordini_fornitori.asp?idord="&arrRS(0, i)&"'""></td><td align=""center"">"&arrRS(1, i)&"</td><td align=""center"">"&formatdatetime(arrRS(2, i),2)&"</td><td align=""center"">"&stato_ordine(arrRS(3, i))&"</td><td align=""center"">"&formatnumber(arrRS(4, i),2)&"</td></tr>"


				Next
				stringa=stringa&"</table>"
				Erase arrRS
			End If
	'***************************************************************************************************************************************************************************************
		case 20	'Accessori MOD_DIP
	'***************************************************************************************************************************************************************************************

			set rs= conn.execute ("select mod_dip from utenti_clienti where iduser="&iduser)
			if not rs.eof then
				jsonString=rs(0)
				if jsonString="" then
				jsonString="{}"
				end if
			else
				jsonString="{}"
			end if
			'jsonString="{""velcrocamicia"":1}"
			'Response.write(jsonString)
			set JSON = New JSONobject
			JSON.Parse(jsonString)
			%>
        <form id="form-accessori" name="" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" >
		<label>Velcro DX</label>
         <input name="velcrocamicia" type="checkbox" value="1" <%=checked_se_1(json("velcrocamicia"))%>>Camicia <input name="velcrogiacca" type="checkbox" value="1" <%=checked_se_1(json("velcrogiacca"))%>>Giacca <input name="velcrogiubbino" type="checkbox" value="1" <%=checked_se_1(json("velcrogiubbino"))%>>Giubbino

		<br />
		<label>Cintura e passanti</label><input name="cintura" type="radio" value="1" <%if json("cintura")="1" then response.write " checked" %>>SI
		&nbsp; &nbsp;
		<input name="cintura" type="radio" value="0" <%if json("cintura")="0" then response.write " checked" %>>NO

		<br>
		<label>Camicia M/L</label><input name="camiciaM" type="checkbox" value="1" <%=checked_se_1(json("camiciam"))%>>Militare <input name="camiciaC" type="checkbox" value="1" <%=checked_se_1(json("camiciac"))%>>Civile
		<br>
		<label>Camiciotto M/M militare</label><input name="camiciottob" type="checkbox" value="1" <%=checked_se_1(json("camiciottob"))%>>Tutto bottoni <input name="camiciottop" type="checkbox" value="1" <%=checked_se_1(json("camiciottop"))%>>Polo
		<br>
		<label>Interasse fori berretto</label><input name="interasse" type="text" value="<%=json("interasse")%>" >cm		<br>
		<label>Note</label>
		<textarea name="note_accessori" style="width:100%; height:100px;"><%=json("note")%></textarea>
		<input type="button" value="Salva" class="modifica-da-tab" data-id="<%=iduser%>" data-oper="update_accessori" data-nomeid="iduser" data-url="<%=questofile%>?tab=20" data-form="form-accessori">
        </form>
        <% if dati_salvati then %>
        <script>
        	toastr.options = {"positionClass": "toast-bottom-right"};
			toastr.success("Dati salvati");


	        </script>

        <%end if%>
			<%

	'***************************************************************************************************************************************************************************************
		case 21	'Intestazioni e destinazioni
	'***************************************************************************************************************************************************************************************
			response.write "<b>Intestazioni</b>"
			sql="select utenti_intestazioni.* FROM utenti_intestazioni where iduser="&iduser
			set rs=conn.execute (sql)
			if not rs.eof then
				call query_to_table(rs)
			End If

			response.write "<b>Consegna</b>"
			sql="select utenti_consegna.* FROM utenti_consegna where iduser="&iduser
			set rs=conn.execute (sql)
			if not rs.eof then
				call query_to_table(rs)
			End If

	'***************************************************************************************************************************************************************************************

		'***************************************************************************************************************************************************************************************
		case 22	'Preferiti
	'***************************************************************************************************************************************************************************************

			sql="select prodotti.idpro, codice, articolo, prezzo, promozione, sconto FRom prodotti inner join preferiti on prodotti.idpro = preferiti.idpro where iduser="&iduser& " order by codice"
			set rs=conn.execute (sql)
			if not rs.eof then arrRS = RS.GetRows()
			Set RS = Nothing
			If IsArray(arrRS) Then
				stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
					"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Codice</td><td>Articolo</td><td align=""right"">Prezzo<br>Unitario</td></tr>"&vbcrlf
				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					stringa_tmp="<tr style=""border-bottom:1px solid;""><td align=""center""><a href=""product.asp?idpro="&arrRS(0, i)&"""><b>"&arrRS(1, i)&"</b></a></td><td>"&arrRS(2, i)

					prezzo=arrRS(3, i)


					if arrRS(4, i)=true then
						prezzoeff=prezzo*(1-arrRS(5, i)/100)
						txtprezzo="<s>"&simbolo_valuta&formatnumber(prezzo,2,true)&"</s><br><b>" & simbolo_valuta&formatnumber(prezzoeff,2,true)&"</b>"
					else
						prezzoeff=prezzo
						if isnull(prezzoeff) or prezzoeff="" then
							txtprezzo="NULL"
						else
							txtprezzo=formatnumber(prezzoeff,2)
						end if
					end if

					stringa_tmp=stringa_tmp&"</td><td align=""right"">"&txtprezzo&"</td><tr>"
					stringa=stringa&stringa_tmp
				Next
				stringa=stringa&"</table>"
			sql="select count(*) from (select p.idpro,  p.iduser from preferiti p left join prodotti_spettanze on p.iduser = prodotti_spettanze.iduser and p.idpro = prodotti_spettanze.idpro where p.iduser="&iduser&" and prodotti_spettanze.idpro is null) as t"
			n_importa=clng(conn.execute(sql)(0))
				if n_importa>0 then
				stringa=stringa&"<input type=""button"" id=""copia_preferiti"" value=""Copia "&n_importa&" preferiti in spettanze"">"
				end if
				Erase arrRS

				%>
				<script>
		$(function() {
			$("#copia_preferiti").click(function(e){
				e.preventDefault();
				e.stopPropagation();
				console.log("copia_preferiti");
				var dati="oper=copia_preferiti&iduser=<%=iduser%>";

				$.ajax({
					url     : "ajax_function.asp?oper=copia_preferiti",
					type    : "post",
					cache	: false,
					dataType: 'json',
					data	: dati,

					success: function(data){


							toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
							toastr.info(data.numero+" preferiti copiati in articoli per spettanze");

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
						invia_errore("Errore in copia_preferiti()",txt );
					}
				});

			});
		});
				</script>

				<%



			End If


		'***************************************************************************************************************************************************************************************
		case 23	'FOGLI SPETTANZE
		'***************************************************************************************************************************************************************************************
			sql="select utenti_foglio_spettanze.*, ordini.nord, ordini.data, ordini.stato from utenti_foglio_spettanze left join ordini on utenti_foglio_spettanze.idord = ordini.idord WHERE utenti_foglio_spettanze.iduser="&iduser&" order by completo desc"
			set rs=conn.execute (sql)
			stringa="<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" bordercolor=""#CCCCCC"" class=""tabella1"">"&_
			"<tr bgcolor=""#E5E5E5"" style=""border-bottom:1px solid;""><td align=""center"">Foglio</td><td align=""center"">Vedi</td><td align=""center"">Creato</td><td align=""center"">Completato</td><td align=""center"">Ordine</td></tr>"&vbcrlf

			do while not rs.eof
					ordine=""
					classe=""
					if not isnull(rs("nord")) then
						ordine="<a href=""pag_adm_ordini.asp?idord="&rs("idord")&""">"&rs("nord")&" del "&rs("data")&"</a>"
						classe="class=""stato_"&RS("stato")&""""
					else
						if rs("completo")<>"" then
						ordine="<input type=""button"" value=""Trasforma in ordine"" onclick=""location.href='pag_adm_user.asp?trasformainordine="&RS("id")&"'"""
						end if
					end if

					stringa=stringa&"<tr style=""border-bottom:1px solid;"" "&colore&"><td align=""center"">"&rs("id")&"</td><td align=""center""><a href=""pdf-foglio-spettanze.asp?idfoglio="&rs("id")&""">PDF</a>&nbsp; <a href=""foglio-spettanze.asp?idfoglio="&rs("id")&""" target=""_blank"">foglio</a></td><td align=""center"">"&formatdatetime(RS("creato"),2)&"</td><td align=""center"">"&RS("completo")&"</td><td align=""center"" "&classe&">"&ordine&"</td></tr>"


			rs.movenext
			loop
			stringa=stringa&"</table>"
		'***************************************************************************************************************************************************************************************
		case 24	'UTENTI secondari
		'***************************************************************************************************************************************************************************************
			response.write "<b>Secondari</b>"
			sql="select utenti.nome, utenti.cognome, utenti.ruolo, utenti.email FROM utenti where secondario="&iduser
			set rs=conn.execute (sql)
			if not rs.eof then
				call query_to_table(rs)
			End If





	'***************************************************************************************************************************************************************************************
	end select
	'***************************************************************************************************************************************************************************************
	response.write stringa
	%>
	<script>

		$(function() {

			$(document).on("click",".edit_in_dialog",function(e){
				e.preventDefault();
				$(this).prop("checked",false);
				var dialog=$("#dialog").dialog();
				dialog.html("Attendi...");
				var cosa=$(this).attr("data-cosa");
				console.log ("edit_in_dialog cosa:"+cosa);
				dialog.html("Attendi...");
				var id=$(this).attr("value").replace(cosa, "");
				console.log("ID:"+id);
				$.ajaxSetup({ cache: false });
				$.ajax({
					url     : "pag_adm_user.asp?dialog=modifica_"+cosa+"&id"+cosa+"="+id,
					type    : "post",
					//dataType: 'json',
					//data	: dati,
					success: function(data){
						dialog.html(data);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 600,
							height: "auto",
							title: "Modifica "+cosa
						});
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});

		});
		</script>

	<%


	'response.end

end sub
function class_selected(val)
	if val<>"" then
		class_selected=" class=""cl_selected"""
	else
		class_selected=""
	end if
end function
sub tabella_misure(iddip,dialog)
%>
<%
			set rs_rilievi=conn.execute ("select dipendenti_misure.* from dipendenti_misure where iddip="&iddip&" order by data_rilievo desc")
			primo=true
			%>
		<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1" id="tabella_misure">
		            <tr >

						<td>Edit</td>

						<td></td>
		            </tr>
					<%

					do while not rs_rilievi.eof



					%>
		            <tr style="border-bottom: 1px solid black;">

			            <td >
				            <%if ha_il_permesso("E4") then %>
				            <input type="radio" id="misura<%=rs_rilievi("id")%>" class="radio_edit">
				            <%end if %>
			            </td>

			            <td
			             <%if rs_rilievi("tipo_rilievo")=1 then %>bgcolor="#d699ff"<%end if %>
			             <%if rs_rilievi("tipo_rilievo")=2 then %>bgcolor="#f9d0a6"<%end if %>


			            >
				            <%if primo then %>

							<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1" >
								<tr
								<%if rs_rilievi("tipo_rilievo")=1 then %>bgcolor="#d699ff"<%end if %>
								<%if rs_rilievi("tipo_rilievo")=2 then %>bgcolor="#f9d0a6"<%end if %>
								>
									<td align="center" ><strong>Rilievo misure del <%=rs_rilievi("data_rilievo")%></strong></td>
								</tr>
								<%if rs_rilievi("tipo_rilievo")=1 then%>
								<tr bgcolor="#d699ff">
									<td align="center">Nuovo metodo</td>
								</tr>


								<%end if%>
								<%if rs_rilievi("tipo_rilievo")=2 then%>
                                    <tr bgcolor="#f9d0a6">
                                        <td align="center">Nuovo metodo</td>
                                    </tr>


                                <%end if%>

							</table>
							<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1" >
					            <tr>
						            <td colspan="1" align="center" bgcolor="#a8dcf0">GIACCA<br>Tecnica</td>
						            <td colspan="7" align="center" bgcolor="#87CEEB">GIACCA</td>
						            <td colspan="1" align="center" bgcolor="#66c1e5">Cappotto</td>
					            </tr>
					            <tr>
						            <td class="cella-misura-head">Taglia</td>
						            <td class="cella-misura-head">TG</td>
						            <td class="cella-misura-head">Fondo</td>
						            <td class="cella-misura-head">Manica</td>
						            <td class="cella-misura-head">Torace</td>
						            <td class="cella-misura-head">Vita</td>
						            <td class="cella-misura-head">Bacino</td>
						            <td class="cella-misura-head">Spalle</td>
						            <td class="cella-misura-head">Lunghezza</td>
					            </tr>
					            <tr >
						            <td class="cella-misura"><%=rs_rilievi("giacca_tecnici")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_TG")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_fondo")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_manica")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_Torace")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_vita")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_bacino")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_spalle")%></td>
						            <td class="cella-misura"><%=rs_rilievi("cappotto")%></td>
					            </tr>
					            <tr style="border-bottom: 1px solid black;">
						            <td class="cella-misura" colspan="1" style="border-top: 1px solid black;">Variazioni</td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_tg2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_fondo2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_manica2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_Torace2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_vita2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_bacino2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("giacca_spalle2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("cappotto2")%></td>
					            </tr>
						            <%if rs_rilievi("note_giaccatec")<>"" then %>
						            <tr >
						            <td colspan="1">Note giacca tecnica:</td>
						            <td colspan="8"><%=rs_rilievi("note_giaccatec")%></td>
						            </tr>
						            <%end if %>
						            <%if rs_rilievi("note_giacca")<>"" then %>
						            <tr>
						            <td colspan="1">Note giacca:</td>
						            <td colspan="8"><%=rs_rilievi("note_giacca")%></td>
						            </tr>
						            <%end if %>
						            <%if rs_rilievi("note_cappotto")<>"" then %>
						            <tr>
						            <td colspan="1">Note cappotto:</td>
						            <td colspan="8"><%=rs_rilievi("note_cappotto")%></td>
						            </tr>
						            <%end if %>
					            </tr>
							</table>
							<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1" >
					            <tr>
						            <td  colspan="1" align="center" bgcolor="#F2F5A9">PANTALONI<br>Tecnici</td>
						            <td  colspan="7" align="center" bgcolor="#FFFF00">PANTALONI</td>
						            <td  colspan="1" align="center" bgcolor="#F4FA58">GONNA</td>
					            </tr>
					            <tr>
						            <td class="cella-misura-head">Tecnici</td>
						            <td class="cella-misura-head">TG</td>
						            <td class="cella-misura-head">Lunghezza</td>
						            <td class="cella-misura-head">Vita</td>
						            <td class="cella-misura-head">Bacino</td>
						            <td class="cella-misura-head">Cosce</td>
						            <td class="cella-misura-head">Cavallo</td>
						            <td class="cella-misura-head">Polpacci</td>
						            <td class="cella-misura-head">Lunghezza</td>
					            </tr>
					            <tr>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_tecnici")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_tg")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_lunghezza")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_vita")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_bacino")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_cosce")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_cavallo")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_polpacci")%></td>
						            <td class="cella-misura"><%=rs_rilievi("gonna")%></td>
					            </tr>
					            <tr style="border-bottom: 1px solid black;">
						            <td class="cella-misura" colspan="1" style="border-top: 1px solid black;">Variazioni</td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_tg2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_lunghezza2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_vita2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_bacino2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_cosce2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_cavallo2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("pantaloni_polpacci2")%></td>
						            <td class="cella-misura"><%=rs_rilievi("gonna2")%></td>
					            </tr>
						            <%if rs_rilievi("note_pantalonitec")<>"" then %>
						            <tr>
						            <td colspan="1">Note pantaloni tec:</td>
						            <td colspan="8"><%=rs_rilievi("note_pantalonitec")%></td>
						            </tr>
						            <%end if %>
						            <%if rs_rilievi("note_pantaloni")<>"" then %>
						            <tr>
						            <td colspan="1">Note pantaloni:</td>
						            <td colspan="8"><%=rs_rilievi("note_pantaloni")%></td>
						            </tr>
						            <%end if %>
						            <%if rs_rilievi("note_gonna")<>"" then %>
						            <tr>
						            <td colspan="1">Note gonna:</td>
						            <td colspan="8"><%=rs_rilievi("note_gonna")%></td>
						            </tr>
						            <%end if %>
							</table>
							<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1" >
								<tr>
						            <td  colspan="2" align="center" bgcolor="#00ff99">CAMICIA</td>
					            </tr>
								<tr>
						            <td class="cella-misura-head">Misura</td>
						            <td class="cella-misura-head" align="left">Note</td>
								</tr>
						            <td class="cella-misura"><%=rs_rilievi("camicia")%></td>
						            <td class="cella-misura"><%=rs_rilievi("camicia_note")%></td>
							</table>
							<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1" >
					            <tr>
						            <td colspan="8" align="center">ALTRO</td>
					            </tr>
					            <tr>
						            <td class="cella-misura-head">Scarpa</td>
						            <td class="cella-misura-head">Collant</td>
						            <td class="cella-misura-head">Berretto</td>
						            <td class="cella-misura-head">Guanto</td>
						            <td class="cella-misura-head">Maglieria</td>
						            <td class="cella-misura-head">Polo/t-shirt</td>
						            <td class="cella-misura-head">Cintura</td>
						            <td class="cella-misura-head">Cinturone</td>
					            </tr>
					            <tr>
						            <td class="cella-misura"><%=rs_rilievi("scarpa")%><%if rs_rilievi("note_scarpa")<>"" then response.write ":"&rs_rilievi("note_scarpa") end if %></td>
						            <td class="cella-misura"><%=rs_rilievi("collant")%><%if rs_rilievi("note_collant")<>"" then response.write ":"&rs_rilievi("note_collant") end if %></td>
						            <td class="cella-misura"><%=rs_rilievi("berretto")%><%if rs_rilievi("note_berretto")<>"" then response.write ":"&rs_rilievi("note_berretto") end if %></td>
						            <td class="cella-misura"><%=rs_rilievi("guanto")%><%if rs_rilievi("note_guanto")<>"" then response.write ":"&rs_rilievi("note_guanto") end if %></td>
						            <td class="cella-misura"><%=rs_rilievi("maglieria")%><%if rs_rilievi("note_maglieria")<>"" then response.write ":"&rs_rilievi("note_maglieria") end if %></td>
						            <td class="cella-misura"><%=rs_rilievi("polo")%><%if rs_rilievi("note_polo")<>"" then response.write ":"&rs_rilievi("note_polo") end if %></td>
						            <td class="cella-misura"><%=rs_rilievi("cintura")%><%if rs_rilievi("note_cintura")<>"" then response.write ":"&rs_rilievi("note_cintura") end if %></td>
						            <td class="cella-misura"><%=rs_rilievi("cinturone")%><%if rs_rilievi("note_cinturone")<>"" then response.write ":"&rs_rilievi("note_cinturone") end if %></td>
					            </tr>

								<%nota_misura=rs_rilievi("nota_misura")
								if nota_misura<>"" then%>
								<tr style="border-top: 1px solid black;">
						            <td colspan="8" align="left">Note: <%=nota_misura%></td>
								</tr>

								<%end if%>
				            </table>
				            <%primo=false

					            else


						            %>
					           <strong>Rilievo misure del <%=rs_rilievi("data_rilievo")%></strong>

					           <% end if %>

			            </td>
		            </tr>

				<%
				rs_rilievi.movenext
				loop

				%>
					<tr >
			            <td colspan="2" align="right" style="height: 1px; border-top:1px solid black;">
				            <%if ha_il_permesso("E4") then %>
				            <button type="button" id="aggiungi_misura">Nuovo rilievo misure</button>
				            <%end if %>

				            </td>
					</tr>

	            </table>
				<script>
					$(function() {

						$("#aggiungi_misura").click(function(e){
							e.preventDefault();

							$.ajax({
								url     : "pag_adm_user.asp?aggiungi_misura=<%=iddip%>",
								type    : "post",
								cache	: false,
								//dataType: 'json',
								//data	: $("#form-dip").serialize(),
								success: function(data){

									$("#dialog").html(data);
									$("#dialog").dialog({
										autoOpen: true,
										modal: true,
										resizable: false,
										width: "auto",
										height: "auto",
										title: "Aggiungi misure"
									});
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;

						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});
					}
				});
			});

				});

				</script>




<%
end sub
sub aggiungi_misura(oper)
if oper="add" then
	iddip=request.querystring("aggiungi_misura")
	dim rs_rilievi(70)
	rs_rilievi(2)=date()
elseif oper="edit" then
	id=request.querystring("modifica_misura")
	set rs_rilievi=conn.execute("select dipendenti_misure.* from dipendenti_misure where id="&id )
	iddip=rs_rilievi("iddip")
end if
	set rs_dipendente=conn.execute("select cognome,nome, nome_grado, ornamento from dipendenti left join utenti_gradi on dipendenti.grado = utenti_gradi.id where iddip="&iddip)
	%>
	Dipendente: <b><%=rs_dipendente("cognome")&" "&rs_dipendente("nome")%></b> <%=rs_dipendente("nome_grado")&" "&rs_dipendente("ornamento")%><br>
	<form id="form_misura">
		<input name="id" type="hidden" id="id" value="<%=id%>">
		<input name="iddip" type="hidden" id="iddip" value="<%=iddip%>">
		<label>Data rilievo misure</label>
		<input name="campo2" type="text" id="data_rilievo" class="richiesto" value="<%=rs_rilievi(2)%>" tabindex="-1">Metodo rilievo:
		Vecchio<input type="radio" name="campo46" value="0" <%if rs_rilievi(46)=0 then response.write "checked"%>>
		Nuovo<input type="radio" name="campo46" value="1" <%if rs_rilievi(46)=1 then response.write "checked"%>>
		4 bottoni<input type="radio" name="campo46" value="2" <%if rs_rilievi(46)=2 then response.write "checked"%>>
		<br />
		<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">

	        <tr>
	            <td colspan="1" align="center" bgcolor="#a8dcf0">GIACCA<br>Tecnica</td>
	            <td colspan="7" align="center" bgcolor="#87CEEB">GIACCA</td>
	            <td colspan="1" align="center" bgcolor="#66c1e5">Cappotto</td>
	        </tr>
	        <tr>
	            <td>TG Tecnica</td>
	            <td>TG</td>
	            <td>Fondo</td>
	            <td>Manica</td>
	            <td>Torace</td>
	            <td>Vita</td>
	            <td>Bacino</td>
	            <td>Spalle</td>
	            <td>Cappotto</td>
	        </tr>
	        <tr>
	            <td><input type="text" name="campo23" class="taglie" id="campo23" value="<%=rs_rilievi(23)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo3" class="taglie" value="<%=rs_rilievi(3)%>"></td>
	            <td><input type="text" name="campo4" class="taglie" value="<%=rs_rilievi(4)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo5" class="taglie" value="<%=rs_rilievi(5)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo6" class="taglie" value="<%=rs_rilievi(6)%>"></td>
	            <td><input type="text" name="campo7" class="taglie" id="vitagiacca" value="<%=rs_rilievi(7)%>"></td>
	            <td><input type="text" name="campo42" class="taglie" value="<%=rs_rilievi(42)%>"></td>
	            <td><input type="text" name="campo8" class="taglie" value="<%=rs_rilievi(8)%>"></td>
	            <td><input type="text" name="campo25" class="taglie" value="<%=rs_rilievi(25)%>" style="border-color: #00ccff;"></td>
	        </tr>
	        <tr>
	            <td colspan="1">Variazioni</td>
	            <td><input type="text" name="campo40" class="taglie" value="<%=rs_rilievi(40)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo27" class="taglie" value="<%=rs_rilievi(27)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo28" class="taglie" value="<%=rs_rilievi(28)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo29" class="taglie" value="<%=rs_rilievi(29)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo30" class="taglie" value="<%=rs_rilievi(30)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo43" class="taglie" value="<%=rs_rilievi(43)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo31" class="taglie" value="<%=rs_rilievi(31)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo32" class="taglie" value="<%=rs_rilievi(32)%>" style="border-color: #00ccff;"></td>
	        </tr>
	        <tr>
		        <td colspan="2">Note giacca tecnica</td>
		        <td colspan="7"><input type="text" name="campo50" class="" value="<%=rs_rilievi(50)%>" style=" width:95%;"></td>
	        </tr>
	        <tr>
		        <td colspan="2">Note giacca</td>
		        <td colspan="7"><input type="text" name="campo48" class="" value="<%=rs_rilievi(48)%>" style=" width:95%;"></td>
	        </tr>
	        <tr>
		        <td colspan="2">Note cappotto</td>
		        <td colspan="7"><input type="text" name="campo52" class="" value="<%=rs_rilievi(52)%>" style=" width:95%;"></td>
	        </tr>
	    </table>
		<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
	        <tr>
	            <td  colspan="1" align="center" bgcolor="#F2F5A9">PANTALONI<br>Tecnici</td>
	            <td  colspan="7" align="center" bgcolor="#FFFF00">PANTALONI</td>
	            <td  colspan="1" align="center" bgcolor="#F4FA58">GONNA</td>
	        </tr>
	        <tr>
	            <td>TG Tecnica</td>
	            <td>TG</td>
	            <td>Lunghezza</td>
	            <td>Vita</td>
	            <td>Bacino</td>
	            <td>Cosce</td>
	            <td>Cavallo</td>
	            <td>Polpacci</td>
	            <td>Lunghezza</td>
	        </tr>
	        <tr>
	            <td><input type="text" name="campo22" class="taglie" value="<%=rs_rilievi(22)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo9" class="taglie" value="<%=rs_rilievi(9)%>"></td>
	            <td><input type="text" name="campo10" class="taglie" value="<%=rs_rilievi(10)%>"></td>
	            <td><input type="text" name="campo11" class="taglie" id="vitapantalone" value="<%=rs_rilievi(11)%>"></td>
	            <td><input type="text" name="campo12" class="taglie" value="<%=rs_rilievi(12)%>"></td>
	            <td><input type="text" name="campo18" class="taglie" value="<%=rs_rilievi(18)%>"></td>
	            <td><input type="text" name="campo19" class="taglie" value="<%=rs_rilievi(19)%>"></td>
	            <td><input type="text" name="campo20" class="taglie" value="<%=rs_rilievi(20)%>"></td>
	            <td><input type="text" name="campo24" class="taglie" value="<%=rs_rilievi(24)%>" style="border-color: #00ccff;"></td>
	        </tr>
	        <tr>
	            <td colspan="1">Variazioni</td>
	            <td><input type="text" name="campo41" class="taglie" value="<%=rs_rilievi(41)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo33" class="taglie" value="<%=rs_rilievi(33)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo34" class="taglie" value="<%=rs_rilievi(34)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo35" class="taglie" value="<%=rs_rilievi(35)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo36" class="taglie" value="<%=rs_rilievi(36)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo37" class="taglie" value="<%=rs_rilievi(37)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo38" class="taglie" value="<%=rs_rilievi(38)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo39" class="taglie" value="<%=rs_rilievi(39)%>" style="border-color: #00ccff;"></td>
	        </tr>
	        <tr>
		        <td colspan="2">Note pantaloni tecnici</td>
		        <td colspan="7"><input type="text" name="campo51" class="" value="<%=rs_rilievi(51)%>" style=" width:95%;"></td>
	        </tr>
	        <tr>
		        <td colspan="2">Note pantaloni</td>
		        <td colspan="7"><input type="text" name="campo47" class="" value="<%=rs_rilievi(47)%>" style=" width:95%;"></td>
	        </tr>
	        <tr>
		        <td colspan="2">Note gonna</td>
		        <td colspan="7"><input type="text" name="campo53" class="" value="<%=rs_rilievi(53)%>" style=" width:95%;"></td>
	        </tr>

	    </table>

		<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
	        <tr>
	            <td  colspan="2" align="center" bgcolor="#00ff99">CAMICIA</td>
	        </tr>
	        <tr>
	            <td>Misura</td>
	            <td>Note</td>
	        </tr>
	        <tr>
	            <td><input type="text" name="campo13" class="taglie" value="<%=rs_rilievi(13)%>" style="border-color: #00ccff;"></td>
	            <td><input type="text" name="campo26" class="" style="width: 95%;" value="<%=rs_rilievi(26)%>" style="border-color: #00ccff;"></td>
	        </tr>
	    </table>
		<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
	        <tr>
	            <td>Scarpa</td>
	            <td>Collant</td>

	            <td>Berretto</td>
	            <td>Guanto</td>
	            <td>Maglieria</td>
	            <td>Polo/t-shirt</td>

	            <td>Cintura</td>
	            <td>Cinturone</td>
	        </tr>
	        <tr>
	            <td><input type="text" name="campo14" class="taglie" value="<%=rs_rilievi(14)%>" style="border-color: #00ccff;">
	            <a href="#" class="nota_altro" data-campo="scarpa" data-nota="<%=rs_rilievi(56)%>"><span class="ui-icon ui-icon-note"></span></a></td>
	            <td><input type="text" name="campo54" class="taglie" value="<%=rs_rilievi(54)%>" style="border-color: #00ccff;">
	            <a href="#" class="nota_altro" data-campo="collant"  data-nota="<%=rs_rilievi(57)%>"><span class="ui-icon ui-icon-note"></span></a></td>
	            <td><input type="text" name="campo15" class="taglie" value="<%=rs_rilievi(15)%>" style="border-color: #00ccff;">
	            <a href="#" class="nota_altro" data-campo="berretto"  data-nota="<%=rs_rilievi(58)%>"><span class="ui-icon ui-icon-note"></span></a></td>
	            <td><input type="text" name="campo16" class="taglie" value="<%=rs_rilievi(16)%>" style="border-color: #00ccff;">
	            <a href="#" class="nota_altro" data-campo="guanto"  data-nota="<%=rs_rilievi(59)%>"><span class="ui-icon ui-icon-note"></span></a></td>
	            <td><input type="text" name="campo21" class="taglie" value="<%=rs_rilievi(21)%>" style="border-color: #00ccff;">
	            <a href="#" class="nota_altro" data-campo="maglieria"  data-nota="<%=rs_rilievi(60)%>"><span class="ui-icon ui-icon-note"></span></a></td>
	            <td><input type="text" name="campo55" class="taglie" value="<%=rs_rilievi(55)%>" style="border-color: #00ccff;">
	            <a href="#" class="nota_altro" data-campo="polo"  data-nota="<%=rs_rilievi(61)%>"><span class="ui-icon ui-icon-note"></span></a></td>

	            <td><input type="text" name="campo44" class="taglie" id="cintura" value="<%=rs_rilievi(44)%>" title="shift+click  per calcolo misura" style="border-color: #00ccff;">
	            <a href="#" class="nota_altro" data-campo="cintura"  data-nota="<%=rs_rilievi(62)%>"><span class="ui-icon ui-icon-note"></span></a></td>
	            <td><input type="text" name="campo45" class="taglie" id="cinturone" value="<%=rs_rilievi(45)%>" title="shift+click  per calcolo misura" style="border-color: #00ccff; display: inline;">
	            <a href="#" class="nota_altro" data-campo="cinturone"  data-nota="<%=rs_rilievi(63)%>" style="display: inline;"><span class="ui-icon ui-icon-note"></span></a></td>
	        </tr>
			<tr>
	            <td colspan="8">Note misurazione:<br>
				<textarea name="campo17" style="width:95%; height:30px;" ><%=rs_rilievi(17)%></textarea>
				</td>
			</tr>
	    </table>
	    <input type="hidden" id="nota_scarpa" value="<%=rs_rilievi(56)%>" name="campo56" >
	    <input type="hidden" id="nota_collant" value="<%=rs_rilievi(57)%>" name="campo57" >
	    <input type="hidden" id="nota_berretto" value="<%=rs_rilievi(58)%>" name="campo58" >
	    <input type="hidden" id="nota_guanto" value="<%=rs_rilievi(59)%>" name="campo59" >
	    <input type="hidden" id="nota_maglieria" value="<%=rs_rilievi(60)%>" name="campo60" >
	    <input type="hidden" id="nota_polo" value="<%=rs_rilievi(61)%>" name="campo61" >
	    <input type="hidden" id="nota_cintura" value="<%=rs_rilievi(62)%>" name="campo62" >
	    <input type="hidden" id="nota_cinturone" value="<%=rs_rilievi(63)%>" name="campo63" >


	    <%if false then %>

	            	<strong>Articoli in ordini in <%=stato_ordine(5)%> o antecedenti:</strong><br>

		            <%
			            dim rs_misure_in_ordini
			            sql="select ordini_dett_spettanze.taglia_misura, ordini_dett.articolo_ordine, ordini.nord, ordini.data, ordini.idord from (ordini_dett_spettanze inner join ordini_dett on ordini_dett_spettanze.iddett = ordini_dett.iddett) inner join ordini on ordini_dett.idord = ordini.idord where ordini.stato<=5 and ordini_dett_spettanze.iddip="&iddip
			            'Response.write sql
			            set rs_misure_in_ordini=conn.execute(sql)
			            if not rs_misure_in_ordini.eof then
			            	%>
			            	<table border="1" cellpadding="2" cellspacing="0" width="500px">
				            	<tr>
					            	<td>Ordine</td>
					            	<td>Data</td>
					            	<td>Articolo</td>
					            	<td>Taglia</td>
				            	</tr>
			            	<%
				            do while not rs_misure_in_ordini.EOF
				            	%>
				            	<tr>
					            	<td><%=rs_misure_in_ordini("nord")%></td>
					            	<td><%=FormatDateTime( rs_misure_in_ordini("data"),2)%></td>
					            	<td><%=rs_misure_in_ordini("articolo_ordine")%></td>
					            	<td><%=rs_misure_in_ordini("taglia_misura")%></td>
				            	</tr>
				            	<%
				            	rs_misure_in_ordini.MoveNext
				            loop
				            %>
			            	</table>
				            <strong>Salvando i dati, le misure esistenti negli ordini in <%=stato_ordine(5)%> o antecedenti verranno sostituite con le misure presenti in questa scheda:</strong>

				            <%
			            end if

			            %>
			            <%end if %>

		<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
			<tr>
	            <td align="center">
		            <input type="submit" id="salva_misura" name="salva_misura" value="Salva"><input type="submit" id="elimina_misura" name="elimina_misura" value="Elimina misura">
		        </td>

			</tr>
		</table>





	</form>
<script type="text/javascript" src="Jquery/js/jquery.validate.min.js"></script>
<script type="text/javascript" src="Jquery/js/jquery.validate.min_it.js"></script>
<script type="text/javascript" src="Jquery/Js/jquery.mobile-events.min.js"></script>

<script>
	$(function() {
		$( "#form_misura :input" ).tooltip({
	      position: {
	        collision: "none"
	      }
	    });
		$("#data_rilievo" ).datepicker();
		$.validator.addMethod(
		    "dateITA",
		    function(value, element) {
		        // put your own logic here, this is just a (crappy) example
		        return value.match(/^\d\d?\/\d\d?\/\d\d\d\d$/);
		    },
		    "Please enter a date in the format dd/mm/yyyy."
		);
		var form =$('#form_misura');
		form.validate({
	       ignore: "",
	       ignoreTitle: true,
	       rules: {

				campo2:{
					required:true,dateITA:true
		        }

			},
			invalidHandler: function(e, validator){
	           if(validator.errorList.length){

				}
	        }
	   });


	   $('.nota_altro').click(function(){
		   var campo=$(this).data('campo');

		   var $inputHidden=$("#nota_" + campo);
		   var testoNota=$inputHidden.val();

		   var nota = prompt("Nota per "+campo, testoNota);
		   if(nota==null)return;
		   console.log(nota);
		   $inputHidden.val(nota);


		   if(false){
		   $(this).data('nota',nota);
			$.ajax({
				url     : "ajax_function.asp",
				type    : "post",
				cache	: false,
				dataType: 'json',
				data	: {'oper':'salva_nota_campo',
					'campo':campo,
					'nota':nota,
					'id':$('#id').val()
				},

				success: function(data){

					if (data.success){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.info("Nota aggiornata");


					console.log ("Nota aggiornata");

					}else{
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error("Errore nell'aggiornamento dei dati: "+ data.message);
					}

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
					invia_errore("Errore in salva_misura()",txt );
				}
			});
}



	   });



		$("#salva_misura").click(function(e){
			e.preventDefault();
			form.validate();
			if (! form.valid()){
				return false;
			}

			var dati="oper=salva_misura_dipendente&"+$( "#form_misura" ).serialize();
			$.ajax({
				url     : "ajax_function.asp",
				type    : "post",
				cache	: false,
				dataType: 'json',
				data	: dati,

				success: function(data){

					if (data.success){


					console.log ("Chiamo carica_tabella_misure");
					if (dialog==""){
						carica_tabella_misure();
					}else
					{

							if (typeof aggiorna_tabella === "function") {

						var dati = {};

						aggiorna_tabella(dati);
						}
					}
					$("#dialog").dialog("close");
					}else{
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error("Errore nell'aggiornamento dei dati: "+ data.message);
					}

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
					invia_errore("Errore in salva_misura()",txt );
				}
			});

		});
		$("#cintura").click(function(e){
		    if(e.shiftKey){
			    var misura=$("#vitapantalone").val();
			    $(this).val(misura*2+15);
			}
		});
	    $('#cintura').bind('doubletap', function(e) {
			    var misura=$("#vitapantalone").val();
			    $(this).val(misura*2+15);
		});



		$("#cinturone").click(function(e){
		    if(e.shiftKey){
			    var misura=$("#vitagiacca").val();
			    $(this).val(misura*2+20);
			}
		});
	    $('#cinturone').bind('doubletap', function(e) {
			    var misura=$("#vitagiacca").val();
			    $(this).val(misura*2+20);
		});

		$("#elimina_misura").click(function(e){
			e.preventDefault();
			var dati="oper=elimina_misura_dipendente&id="+$( "#id" ).val();
			$.ajax({
				url     : "ajax_function.asp",
				type    : "post",
				cache	: false,
				dataType: 'json',
				data	: dati,

				success: function(data){
					console.log ("Chiamo carica_tabella_misure");
					carica_tabella_misure();
					$("#dialog").dialog("close");

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
					invia_errore("Errore in elimina_misura()",txt );
				}
			});

		});

	});
</script>
<%
end sub
sub dialog_grado(oper)
if oper="add" then
	iduser=request.querystring("iduser")
	dim rs_rilievi(17)
	rs_rilievi(2)=date()
elseif oper="edit" then
	idgrado=request.querystring("idgrado")
	set rs_rilievi=conn.execute("select utenti_gradi.* from utenti_gradi where id="&idgrado)
	iduser=rs_rilievi("iduser")
end if
%>
	<form id="form_grado">
	<input name="idgrado" type="hidden" id="id" value="<%=idgrado%>">
	<input name="iduser" type="hidden" id="iduser" value="<%=iduser%>">
	<label>Nome grado</label>
	<input name="campo1" type="text" value="<%=rs_rilievi(1)%>">
		<br />
		<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr>
	            <td></td>
	            <td align="center" bgcolor="#87CEEB">GIACCA</td>
	            <td align="center" bgcolor="#FFFF00">PANTALONI</td>
	            <td align="center" bgcolor="">CAMICIA</td>
	            <td align="center">Scarpa</td>
	            <td align="center">Berretto</td>
	            <td align="center">Guanto</td>
            </tr>
            <tr>
	            <td>Maschio</td>
	            <td><input type="text" name="campo3" class="taglie" value="<%=rs_rilievi(3)%>"></td>
	            <td><input type="text" name="campo4" class="taglie" value="<%=rs_rilievi(4)%>"></td>
	            <td><input type="text" name="campo5" class="taglie" value="<%=rs_rilievi(5)%>"></td>
	            <td><input type="text" name="campo6" class="taglie" value="<%=rs_rilievi(6)%>"></td>
	            <td><input type="text" name="campo7" class="taglie" value="<%=rs_rilievi(7)%>"></td>
	            <td><input type="text" name="campo8" class="taglie" value="<%=rs_rilievi(8)%>"></td>
            </tr>
            <tr>
	            <td>Femmina</td>
	            <td><input type="text" name="campo9" class="taglie" value="<%=rs_rilievi(9)%>"></td>
	            <td><input type="text" name="campo10" class="taglie" value="<%=rs_rilievi(10)%>"></td>
	            <td><input type="text" name="campo11" class="taglie" value="<%=rs_rilievi(11)%>"></td>
	            <td><input type="text" name="campo12" class="taglie" value="<%=rs_rilievi(12)%>"></td>
	            <td><input type="text" name="campo13" class="taglie" value="<%=rs_rilievi(13)%>"></td>
	            <td><input type="text" name="campo14" class="taglie" value="<%=rs_rilievi(14)%>"></td>
            </tr>
			<tr>
	            <td colspan="7" align="center"><input type="submit" id="salva_grado" value="Salva"><input type="submit" id="elimina" name="elimina_grado" value="Elimina grado"></td>
			</tr>
	            </table>
				</form>
<script>
	$(function() {
		$("#salva_grado").click(function(e){
			e.preventDefault();
			var dati="oper=salva_grado&"+$( "#form_grado" ).serialize();
			$.ajax({
				url     : "ajax_function.asp",
				type    : "post",
				cache	: false,
				dataType: 'json',
				data	: dati,

				success: function(data){
					console.log ("Chiamo carica_tabella_misure");
					aggiorna_tab();
					$("#dialog").dialog("close");

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
					invia_errore("Errore in salva_grado()",txt );
				}
			});

		});
		$("#elimina_grado").click(function(e){
			e.preventDefault();
			var dati="oper=elimina_grado&id="+$( "#id" ).val();
			$.ajax({
				url     : "ajax_function.asp",
				type    : "post",
				cache	: false,
				dataType: 'json',
				data	: dati,

				success: function(data){
					console.log ("Chiamo carica_tabella_misure");
					carica_tabella_misure();
					$("#dialog").dialog("close");

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
					invia_errore("Errore in elimina_grado()",txt );
				}
			});

		});

	});
</script>
<%
end sub

sub dialog_articolo_spettanze(oper)
if oper="add" then
	tipoprodotto=-1
	iduser=request("iduser")
elseif oper="edit" then

	'idgrado=request.querystring("idgrado")
	'set rs_rilievi=conn.execute("select utenti_gradi.* from utenti_gradi where id="&idgrado)
	'iduser=rs_rilievi("iduser")
	id=request("idarticolo_spettanze")
	sql="select prodotti_spettanze.id, prodotti_spettanze.iduser, prodotti.codice, prodotti.articolo, varianti_a.variante_a, varianti_b.variante_b,prodotti_spettanze.ordine from ((prodotti_spettanze INNER JOIN prodotti on prodotti_spettanze.idpro=prodotti.idpro) left join varianti_a on prodotti_spettanze.idvara = varianti_a.idvara ) left join varianti_b on prodotti_spettanze.idvarb = varianti_b.idvarb where id="&id
	set rs=conn.execute(sql)
	pres_ordine=clng(rs("ordine"))
	iduser=rs("iduser")
end if
%>
	<form id="form_articolo">
	<input name="iduser" type="hidden" id="iduser" value="<%=iduser%>">
	<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
	    <tr  >
	      <td    style="border-bottom:1px solid; ">Articolo:
	        <input type='hidden' name='id' id="id" value="<%=id%>"/>
	        <%if oper="add" then %>
			<input type='hidden' name='idpro' id='idpro' style="width:350px;" />
	        <input type='hidden' name='idvara' id='vara' style="width:150px;" class="step1"/>
	        <input type='hidden' name='idvarb' id='varb' style="width:150px;" class="step1"/>
	        <%elseif oper="edit" then %>
	        <%=rs("codice")&" "&rs("articolo")& " "&rs("variante_a")& " "&rs("variante_b")%>
	        <%end if %>
	      </td>
	    </tr>
	    <tr >
	      <td style="border-bottom:1px solid; ">
		  Posizione in elenco:<%=pres_ordine%>
		  <input type="text" name="ordine" value="<%=pres_ordine%>">
	      </td>
	    </tr>
	    <tr >
	      <td  align="center"  style="border-bottom:1px solid; ">
		    <%if oper="add" then %>
			<input type="button" id="salva_articolo_spettanze" value="Aggiungi articolo" disabled="disabled" />
			<%elseif oper="edit" then %>
			<input type="button" id="salva_articolo_spettanze" value="Salva" />
	        <input type="submit" id="elimina_articolo_spettanze" name="rimuovi_articolo" value="Rimuovi articolo">


	        <%end if %>
	      </td>
    </tr>
  </table>
	</form>



	  <script type="text/javascript">
  $(document).ready(function () {
	//Inizializzo select articolo
	var varianti = 0,
		variantea = false,
		varianteb = false,
		idpro = 0,
		idvara = 0,
		idvarb = 0,
		oper="<%=oper%>";


	$('#idpro').select2({
		dropdownCssClass: 'ui-dialog',
		placeholder: 'Seleziona articolo',
		minimumInputLength: 3,
		allowClear: true,
		ajax: {
			quietMillis: 150,
			url: "ajax_function.asp?select2=articoli_sel2",
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
	$("#idpro").on('change', function (e) {
		//console.log("valore idpro val:"+$("#idpro").select2('val'))
		//console.log("valore idpro data:"+$("#idpro").select2('data').id)
		console.log("idpro change");
		posizione_errore="#idpro_change";
		var data = $("#idpro").select2('data');
		$.ajaxSetup({
			cache: false
		});
		idvara = 0,
		idvarb = 0;
		varianti=0;
		variantea=false;
		varianteb=false;
		posizione_errore="#idpro_change: P1";
		if (typeof data == "undefined") {
		   return true;
		}
		posizione_errore="#idpro_change: P1.1";
		if (!data){
			return true;
		}
		posizione_errore="#idpro_change: P1.2";
		idpro=data.id;
		posizione_errore="#idpro_change: P1.5";

		//quantitamagazzino();
		$.getJSON("ajax_function.asp?select2=varianti&term=" + data.id)
			.done(function (json) {
				//console.log(json);
				//Verifico se ha varianti
				posizione_errore="#idpro_change: P2";
				if (json.var_a == "true") {
					varianti += 1;
					variantea = true;
				}
				if (json.var_b == "true") {
					varianti += 1;
					varianteb = true;
				}
				verifica();
				console.log()
				if (json.var_a == "true") {
					posizione_errore="#idpro_change: P3";
					$.getJSON("ajax_function.asp?select2=vara&prezzo=si&term=" + data.id)
						.done(function (json) {
							//Inizializzo select variante A
							$('#vara').select2({
								dropdownCssClass: 'ui-dialog',
								data: json,
								placeholder: 'Variante A'
							});
							$("#vara").select2("val", "");

							$("#vara").on('change', function (e) {
								idvara=$("#vara").val();
								console.log("vara change:"+idvara);
								verifica();
							});
						});
				} else {
					$('#vara').select2("destroy");
				}
				if (json.var_b == "true") {
					posizione_errore="#idpro_change: P4";
					$.getJSON("ajax_function.asp?select2=varb&term=" + data.id)
						.done(function (json) {
							//Inizializzo select variante A
							$('#varb').select2({
								dropdownCssClass: 'ui-dialog',
								data: json,
								placeholder: 'Variante B'
							});
							$("#varb").select2("val", "");
							$("#varb").on('change', function (e) {
								console.log("varb change");
								idvarb=$("#varb").val();
								verifica();
							});
						});
				} else {
					$('#varb').select2("destroy");
				}
			})
			.fail(function (jqxhr, textStatus, error) {
				var err = textStatus + ", " + error;
				console.log("Request Failed: " + err);
			});
	});
	function verifica() {
		//console.log("Inizio verifica variantea:"+variantea +" idvara:"+ idvara);
		posizione_errore="function verifica";
		var verifica = true;

		if (variantea && idvara == 0) {
				verifica = false;
		}
		if (varianteb && idvarb == 0) {
				verifica = false;
		}
		console.log("verifica:"+verifica);
		if (verifica) {
			$("#salva_articolo_spettanze").removeAttr('disabled');
			$("#salva_articolo_spettanze").focus();
		} else {
			$("#salva_articolo_spettanze").attr('disabled', 'disabled');
		}
	}




		$("#salva_articolo_spettanze").click(function(e){
			e.preventDefault();
			var dati="oper=aggiungi_articolo_spettanze&"+$( "#form_articolo" ).serialize();
			$.ajax({
				url     : "ajax_function.asp",
				type    : "post",
				cache	: false,
				dataType: 'json',
				data	: dati,

				success: function(data){
					console.log ("Chiamo carica_tabella_spettanze");
					if (!data.success){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "2000","extendedTimeOut": "0", "closeButton": true};
						toastr.error("Questo articoli &egrave; già presente");



					}
					aggiorna_tab();
					if(oper=="edit"){
						$("#dialog").dialog("close");
					}
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
					invia_errore("Errore in salva_articolo_spettanze()",txt );
				}
			});

		});
		$("#elimina_articolo_spettanze").click(function(e){
			e.preventDefault();
			var dati="oper=elimina_articolo_spettanze&id="+$( "#id" ).val();
			$.ajax({
				url     : "ajax_function.asp",
				type    : "post",
				cache	: false,
				dataType: 'json',
				data	: dati,

				success: function(data){
					console.log ("Chiamo carica_tabella_misure");
					aggiorna_tab();
					$("#dialog").dialog("close");

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
					invia_errore("Errore in elimina_articolo_spettanze()",txt );
				}
			});

		});

	});
</script>
<%
end sub

function green(val)
	green=""
	if val=1 then green="green"
end function
function checked_lingua(lingua,lingue)
	dim n
	checked_lingua=""
	if IsArray(lingue) then
		for n=0 to UBound(lingue)


			if trim(lingue(n))=lingua then
				checked_lingua=" checked"
				exit for
			end if
		next
	end if
end function

function elenco_option_gradi(sql,selected)
	txt=""
	txt="<!-- parametro: "&parametro&"-->"
	if isnumeric(selected) then
		selected=cint(selected)
	else
		selected=0
	end if
	on error resume next
	set rs=conn.execute(sql)
	if err.number<>0 then
		response.write sql
	end if
	on error goto 0
		do while not rs.eof
			txt=txt&"<option value='"&rs("id")&"'"
			if cint(rs("id"))=selected then
				txt=txt& " selected style='color:red;'"
			end if
			if rs("colore")<>"" then
				txt=txt&"style=""background-color: #"&rs("colore")&";"""
			end if

			txt=txt& ">" & rs("testo")&"</option>"
			rs.MoveNext
		loop
        'txt=txt&"</select>"
		elenco_option_gradi=txt
end function

%>
