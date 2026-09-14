<!--#include virtual="/setup.asp" -->
<%
'-----------------------------------------
'-----------------Pagina: pag_cat.asp
'---------- Applicazione: Catalogo
'-----------------------------------------
file=Request.ServerVariables("Script_Name")
idcat=request("idcat")
if idcat="" then
	if request.form("idcathidden") <>"" then
		idcat=request.form("idcathidden")
	else
		idcat=0
	end if
end if
idsettore=request("idsettore")
if idsettore="" then
	if request.form("idsettorehidden") <>"" then
		idsettore=request.form("idsettorehidden")
	else
		idsettore=0
	end if
end if
if request("dove")<>"" then
	if request("dove")="su" then
		dove="-15"
	else
		dove="+15"
	end if
	sql="UPDATE prodotti SET prodotti.ordine = [ordine]"&dove&" WHERE (((prodotti.idpro)="&request("id")&"));"
	conn.execute sql
	Set rs = Server.CreateObject("ADODB.Recordset")
	call riordina(request("id"))
end if

sub riordina(id)
	'trovo categoria del prodotto
	sql="select * from prodotti where idpro="&id
	rs.open sql, conn,3,3
	idcat=rs("idcat")
	rs.close
	sql="select * from prodotti where idcat="&idcat&" order by ordine;"
	'response.write sql
	rs.open sql, conn,3,3
	'response.end
	ordine=10
	do while not rs.eof
	rs("ordine")=ordine
	rs.update
	ordine=ordine+10
	rs.movenext
	loop
	rs.close
end sub
if request.form("cerca")<>"" then
	testo=pulisci(trim(request.form("testo")))
	if len(testo)>0 and session("idadmin")="" then
		if request.form("idpro")<>"" then testo=request.form("idpro")
		sql="select * from ricerche where parola='" & testo &"'"
		rs.Open sql, conn, 3, 3
		if rs.eof then
			rs.addnew
			rs("parola")=left (testo,50)
		end if
		rs("click")=rs("click")+1
		rs.update
		rs.close
		logga "ric",trim(request.form("testo"))
	end if
	idsettore=0
	idcat=0
else
	testo=trim(request("testo"))
end if

sub mappasettoriselezione(layer,idsettore,selezionare,idcat)
	dim sql_order,str,MessageSpacing,mesid,spaceing,tid
	set rs_order=server.createObject("adodb.recordset")
	sql_order = "select * FROM settori WHERE idpadre = " & idsettore &"  order by nome_settore;"
	rs_order.open sql_order,conn
	nodo=true
	do until rs_order.eof
		spaceing=""
		for MessageSpacing=1 to layer
			spaceing=spaceing & "&nbsp;&nbsp;&nbsp;&nbsp;"
		next
		response.write "<option value=""" & rs_order("idsettore") &""""
if int(selezionare)=int(rs_order("idsettore")) then response.write " selected style='background-color:#007FFF;'"
response.write ">" &spaceing& rs_order("nome_settore")& "</option>"
	    call mappasettoriselezione(layer+1,rs_order("idsettore"),selezionare,escludi)
		rs_order.movenext
	loop
	rs_order.close
	set rs_order=nothing
End Sub	
sub mappacatselezione(layer,idcat,selezionare,idsettore)
	dim sql_order,str,MessageSpacing,mesid,spaceing,tid
	set rs_order=server.createObject("adodb.recordset")
	nascondi=" and categorie.nascondi=false " 
	if session("idadmin")<>"" then nascondi=""
	sql_order = "select * FROM categorie WHERE idpadre = " & idcat &nascondi & "  order by ordine;"
	rs_order.open sql_order,conn
	nodo=true
	do until rs_order.eof
		spaceing=""
		for MessageSpacing=1 to layer
			spaceing=spaceing & "&nbsp;&nbsp;&nbsp;&nbsp;"
		next
		response.write "<option value=""" & rs_order("idcat") &""""
if int(selezionare)=int(rs_order("idcat")) then response.write " selected style='background-color:#007FFF;'"

response.write ">" &spaceing& rs_order("categoria")& "</option>"
	    call mappacatselezione(layer+1,rs_order("idcat"),selezionare,escludi)
		rs_order.movenext
	loop
	rs_order.close
	set rs_order=nothing
End Sub	
sub boxsx(tipo)
sql="select * from news where stato='0' and tipo='"&tipo&"' order by ordine;"
select case tipo
case "ht","hb","ha"
	classtitolo="ui-widget-header"
	classcorpo="testotabella"
case else
	classtitolo="titolobox"
	classcorpo="corpobox"
end select
set rs=conn.execute(sql)
do while not rs.eof
response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class="tabella1">"
response.write "<tr><td class="&classtitolo&" valign='middle'>"& rs("titolo")
if session("idadmin")<>"" then
	response.write " <a href='pag_adm_pagine.asp?modifica="&rs("idnews")&"'><img src='images/edit3.gif' alt='Modifica' border='0'></a>"
end if
response.write"</td></tr>"
response.write "<tr><td class="&classcorpo&">"&rs("testo")&"</td></tr>"
response.write "</table>"
rs.movenext
loop
end sub



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

<!--#include virtual="/sub_head.asp" -->

<style type="text/css">
<!--
.separa { border-top:
widht:100%;
clear:both;	
	}
.casella{
	margin: 0px 0px 0px 0; padding: 1px; float: left; width: 32%;  position: relative;
	}	
div.descrizione {
	border-top:  1px solid gray;
	border-left:  1px solid gray;
	margin-top: 2px;
	FONT-SIZE: 10px;
	/*colore testo contenuto tabelle*/
	padding: 1px 1px 1px 1px;
}

-->
</style>
	<style type="text/css">
	/* Big box with list of options */
	#ajax_listOfOptions{
		position:absolute;	/* Never change this one */
		width:240px;	/* Width of box */
		height:250px;	/* Height of box */
		overflow:auto;	/* Scrolling features */
		border:1px solid #FA85F8;	/* Dark green border */
		background-color:#FFF;	/* White background color */
		text-align:left;
		font-size:0.9em;
		z-index:100;
	}
	#ajax_listOfOptions div{	/* General rule for both .optionDiv and .optionDivSelected */
		margin:1px;		
		padding:1px;
		cursor:pointer;
		font-size:0.9em;
	}
	#ajax_listOfOptions .optionDiv{	/* Div for each item in list */
		
	}
	#ajax_listOfOptions .optionDivSelected{ /* Selected item in the list */
		background-color:#FA85F8;
		color:#FFF;
	}
	#ajax_listOfOptions_iframe{
		background-color:#F00;
		position:absolute;
		z-index:5;
	}
	
	form{
		display:inline;
	}
	
	</style>
	<script type="text/javascript" src="js/ajax.js"></script>
	<script type="text/javascript" src="js/ajax-dynamic-list.js">
	/************************************************************************************************************
	(C) www.dhtmlgoodies.com, April 2006
	
	This is a script from www.dhtmlgoodies.com. You will find this and a lot of other scripts at our website.	
	
	Terms of use:
	You are free to use this script as long as the copyright message is kept intact. However, you may not
	redistribute, sell or repost it without our permission.
	
	Thank you!
	
	www.dhtmlgoodies.com
	Alf Magne Kalleland
	
	************************************************************************************************************/	
	</script>
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
	

<style>
DIV.list_item_container {
    height: 102px;
    padding: 3px;
}
DIV.image {
    width: 100px;
    height: 100;
    float: left;
}
DIV.description {
    font-style: italic;
    font-size: 0.9em;
    color: gray;
}
DIV.label {
	font-weight:bold;
    font-size: 1.2em;
}
</style>
<script>
	$(function() {

    $("#cercanew").autocomplete({
		source: "ajax_function.asp?autocomplete=articolinew",
		cache: false ,
        minLength: 2,
		maxCacheLength:0,
		select: function( event, ui ) {
				$('#cerca_hidden').val(ui.item.id); 
				alert();
			}

    }).data( "ui-autocomplete" )._renderItem = function( ul, item ) {
       var inner_html = '<a><div class="list_item_container"><div class="image"><img src="send_img.asp?s=100&id=' + item.id + '"></div><div class="label">' + item.codice + '</div><div class="description">' + item.articolo + '</div></div></a>';
        return $( "<li></li>" )
            .data( "item.autocomplete", item )
            .append(inner_html)
            .appendTo( ul );
    };
});

</script>

</head>
<body>
<%
'response.Write("<br>IDCAT:"&request("idcat")&"IDSETTORE:"&request("idsettore"))
%>
<div id="wrap">
<%barra=3%>
<div id="header"><!--#include virtual="/sub_top.asp" --></div>
<div id="sidebar"><!--#include virtual="/sub_menu.asp" --> </div>
<div id="content_large">
<form name="form1" method="post" action="<%=questofile%>" style="margin:0 0 0 0;"> 
                     <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tabella1">
          <tr> 
            <td width="33%" align="center" class="ui-widget-header"><b><%=Application("testo_settore")%></b><%if session("idadmin")<>"" then%> <a href="pag_adm_testi_catalogo.asp" title="Modifica titoli"><img src="images/edit3.gif" alt="" border="0"></a><a href="pag_adm_settori.asp" title="Modifica settore"><img src="images/edit3.gif" alt="" border="0"></a>
      <%end if%></td>
            <td width="33%" align="center" class="ui-widget-header" style="border-left: 1px solid black; border-right: 1px solid black;"><b><%=Application("testo_categoria")%></b><%if session("idadmin")<>"" then%><a href="pag_adm_categorie.asp" title="Modifica categorie"><img src="images/edit3.gif" alt="" border="0"></a>
      <%end if%><%if session("iduser")<>"" then
			sqlcount="select * from prodcat where idcat="&idcat
		Set rscount=conn.execute(sqlcount)
		if not rscount.eof then
			%> 
            <a href="pag_adm_testi_catalogo.asp"><img src="images/edit3.gif" alt="Modifica titolo" border="0"></a> <a href="pdf_listino.asp?cat=<%=idcat%>">Scarica listino</a>
            <%end if
			rscount.close
			end if%></td>
            <td width="33%" align="center" class="ui-widget-header">Ricerca libera</td>
          </tr>
          <tr> 
           
           
            <td align="center" colspan=3> Inserisci il testo da cercare o selezionalo tra i risultati proposti<br>
        
		<input name="carcanew" id="cercanew" type="text" style="font-size:12px; height: 18px; width:75%;" size="20" autocomplete=off >
		<input type="submit" name="puls_cerca" value="Cerca" style="font-size:12 px; width:20%;">
         <input type="hidden" id="cerca_hidden" name="idpro">
       
		        
    </td>
          </tr>
        </table>
      </form>
        <!-- ELENCO prodotti INIZIO   -------------------------------------------------------- -->
        
<%
prezzo=false
if session("iduser") <> "" or Application("scopri_prezzi") then prezzo=true
if nascondi_prezzi_set or nascondi_prezzi_cat then prezzo=false
%>        
      </div>
    
    <div id="footer"></div><!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
