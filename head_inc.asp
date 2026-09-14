<!--#include virtual="/cart_small_inc.asp" -->
<%
if session("iduser")<>"" then
	where_carrello="iduser="&session("iduser")
else
	if SessionIdGuest="" then
		where_carrello="idloged=0"
		'add2log "manca idloged",3
		'response.end
	else
		where_carrello="idloged="&SessionIdGuest
	end if
end if

if meta_description="" or isnull(meta_description)then
	meta_description=Application("tag_desc")
end if
menu_admin= "<li><a href=""pag_adm_menu.asp"" title=""Modifica menu"" style=""padding-left:10px;padding-right:10px;""><i class=""fa fa-pencil""></i></a><ul><li><a href=""pag_adm_pagine.asp"">Contenuti</a></li><li><a href=""pag_adm_layout.asp"">Organizza home</a></li><li><a href=""pag_adm_gruppihome.asp"">Gruppi articoli</a></li></ul></li>"
call carica_dizionario()
if meta_title="" then
	meta_title=Application("brwstitle")
else
	meta_title_tmp=traduci(meta_title)
	if meta_title_tmp<>"" then
		meta_title=meta_title_tmp
	end if
	meta_title=meta_title&" - "&nomesito
end if
%>
<!DOCTYPE html>
<!--[if IE 8]> <html class="ie8"> <![endif]-->
<!--[if IE 9]> <html class="ie9"> <![endif]-->
<!--[if !IE]><!--> <html> <!--<![endif]-->
    <head>
		<meta charset="UTF-8">
        <title><%=Server.HTMLEncode(meta_title)%></title>
        <meta name="description" content="<%=Server.HTMLEncode(meta_description)%>">
        <%if not mod_larochelle then %>
        <meta name="author" content="Andrea Spotorno">
        <meta name="copyright" content="Andrea Spotorno">
        <%end if %>
        <meta name="ROBOTS" content="index,follow">

       <!--[if IE]> <meta http-equiv="X-UA-Compatible" content="IE=edge"> <![endif]-->
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <%=Application("head")%>
        <%if lingua="" then%>
        <meta http-equiv="Content-Language" content="it" />
        <link rel="alternate" hreflang="en" href="index.asp?lang=en" />
        <%else %>
        <meta http-equiv="Content-Language" content="en" />
        <link rel="alternate" hreflang="it" href="index.asp?lang=it" />
	    <%end if%>
        
        <link href='//fonts.googleapis.com/css?family=PT+Sans:400,700,400italic,700italic%7CPT+Gudea:400,700,400italic%7CPT+Oswald:400,700,300' rel='stylesheet' id="googlefont">
        
        <link rel="stylesheet" href="css/bootstrap.min.css">
        <link rel="stylesheet" href="css/font-awesome.min.css">
        <link rel="stylesheet" href="css/prettyPhoto.css">
        <link rel="stylesheet" href="css/revslider.css">
        <link rel="stylesheet" href="css/owl.carousel.css">        
        <link rel="stylesheet" href="css/style.css">
        <link rel="stylesheet" href="css/mycss.css">
        <link rel="stylesheet" href="css/responsive.css">
        
        <%if utente_admin then %>
		<link rel="stylesheet" href="jquery/css/superfish.css" media="screen">
        <style>
	        .sf-menu a {background: #fff}
	    </style>
        <%end if %>
        
        
        

		<%if select2 then%>        
	        <!-- Select2bootstrap -->
			<link rel="stylesheet" href="jquery/js/select2/select2.css"  />
	        <link rel="stylesheet" href="jquery/js/select2/select2-bootstrap.css">
	        <link rel="stylesheet" href="css/jquery.selectbox.css">
	        <style>
	        .select2-container .select2-choice {height: 38px;}
			</style>
		<%end if%>    

		<%if css_category then%>        
	        <style>
	        .list-group.panel > .list-group-item {
			  border-bottom-right-radius: 4px;
			  border-bottom-left-radius: 4px
			}
			.list-group-submenu {
			  margin-left:20px;
			}
			.strong { font-weight: bold; }
			</style>
		<%end if%>    
			<link rel="stylesheet" href="Jquery/Js/jquery.cookiebar.css"  />
		
		
		<%if css_form then%>        
		<link rel="stylesheet" href="css/bootstrap-switch.css"> 
		<%end if%>
		
		<%if css_FileUpload then%>        
		<!-- CSS to style the file input field as button and adjust the Bootstrap progress bars -->
		<link rel="stylesheet" href="jquery/JQuery-File-Upload/css/jquery.fileupload.css">
		<%end if%>     
		<%if js_toastr then%>        
		<link rel="stylesheet" href="Jquery/css/toastr.min.css" />
		<%end if%>
		<%if fancybox then%>
		<link href="Jquery/Js/fancyBox/jquery.fancybox.css" rel="stylesheet">
		<%end if%>
		
		<%if foglio_spettanze then%>        
		<link rel="stylesheet" href="https://cdn.datatables.net/1.10.15/css/jquery.dataTables.min.css">
		<link rel="stylesheet" href="https://cdn.datatables.net/fixedcolumns/3.2.2/css/fixedColumns.dataTables.min.css">
		<link rel="stylesheet" href="foglio-spettanze.css?<%=ver_js_css%>">
		
		<%end if %>
		
		
		
		
        <!-- Favicon and Apple Icons -->		
		
		<!--#include virtual="/favicon.asp" -->

		
		<%=meta_facebook%>
		<!--- jQuery -->
        <script src="//ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
        <script>window.jQuery || document.write('<script src="js/jquery-1.11.1.min.js"><\/script>')</script>

		<!--[if lt IE 9]>
			<script src="js/html5shiv.js"></script>
			<script src="js/respond.min.js"></script>
		<![endif]-->
	    
		<style id="custom-style">

		/*ui-autocomplete*/
			.ui-autocomplete-loading {
			}
			.ui-autocomplete {
				max-height: 500px;
				overflow-y: auto;
				/* prevent horizontal scrollbar */
				overflow-x: hidden;
				z-index: 1000;
			}
			/* IE 6 doesn't support max-height
			 * we use height instead, but this forces the menu to always be this tall
			 */
			* html .ui-autocomplete {
				height: 500px;
			}
			/*ui-autocomplete con immagini*/
			DIV.list_item_container {
			    padding: 3px;
			}
			DIV.image {
			    width: 100px;
			    height: 100px;
			    float: left;
				margin-right:5px;
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
		<%if session("idadmin")<>"" then %>
			var id_navContainer="#admin-nav-container";
			var id_header="#admin-nav-container";
		<%else %>
			var id_navContainer="#main-nav-container";
			var id_header="#header";
		<%end if%>
		</script>
    </head>
    <!-- <%=TipoUtente&" idguest:"&SessionIdGuest%> -->
<%
sub menu(layer,id,testo,testo_dopo)
	dim sql_order,str,MessageSpacing,mesid,spaceing,tid
	set rs_menu=server.createObject("adodb.recordset")
	sql_order="select menu.idmen, menu.Voce,menu.Voce_en, menu.idpadre, menu.link, Count(menu_1.idmen) AS ConteggioDiidmen, menu.ordine FROM menu AS menu_1 RIGHT JOIN menu ON menu_1.idpadre = menu.idmen GROUP BY menu.idmen, menu.Voce, menu.Voce_en, menu.idpadre, menu.ordine,menu.link HAVING  (menu.idpadre)= " & id 
	'sql_order = "select * FROM menu WHERE idpadre = " & id 
	if id=0 then sql_order=sql_order&" or (menu.idpadre) Is Null "
	sql_order=sql_order&" order by menu.ordine"
	rs_menu.open sql_order,conn
	nodo=true
	testo=testo&testo_dopo
	testo_dopo=""
	if not 	rs_menu.eof then 
		if layer>0 then testo=testo& "<ul>"
	else
		nodo=false
	end if
	do until rs_menu.eof
		link=rs_menu("link")
		if instr(link,"[settore=")>0 then
			val=replace(replace(link,"[settore=",""),"]","")
			testo=testo&  "<li class=""mega-menu-container"">"
			testo_dopo="<a href=""category.asp?idsettore="&val&""">"&rs_menu("voce"&lingua)&"</a>"
			call menusettori(layer+1,val,testo,testo_dopo)
						  
		else
			testo=testo&  "<li>"
			testo_dopo="<a href="""&link&""">"&rs_menu("voce"&lingua)&"</a>"
		end if
		'calling again to find more childs - if not rs_menu=nothing
	    call menu(layer+1,rs_menu("idmen"),testo,testo_dopo)
	'testo=testo& "</li>"&vbcrlf
		rs_menu.movenext
	loop
	if nodo=true and layer>0 then testo=testo& "</ul>"&vbcrlf
	'closing object
	rs_menu.close
	set rs_menu=nothing
	'if layer=0 then response.write testo
End Sub	
sub menusettori(layer,id,testo,testo_dopo)
	dim sql_order,str,MessageSpacing,mesid,spaceing,tid,MaxRow,NumRow
	maxRow=6
	'sql_order="select settori.* from settori where nascondi=0 and idpadre="&id 
	'sql_order=sql_order&" order by ordine"
	
	sql_order="select settori.idsettore, settori.nome_settore,  settori.nascondi_prezzi, settori_settori.idpadre from settori_settori inner join settori on settori_settori.idfiglio = settori.idsettore where settori.nascondi=0 and settori_settori.idpadre="&id& " order by settori_settori.ordine"

	
	set rs_settori=conn.execute(sql_order)
	testo=testo&"<div class=""mega-menu clearfix"">"
	testo=testo&"<div class=""row"">"
	row=0
	do while not rs_settori.eof
		testo=testo&"<div class=""col-5"">"
		testo=testo&"<a href=""category.asp?idsettore="&rs_settori("idsettore")&""" class=""mega-menu-title"">"&rs_settori("nome_settore")&"</a><!-- End .mega-menu-title -->"
		testo=testo&"<ul class=""mega-menu-list clearfix"">"
		
		'sql_order="select settori.* from settori where nascondi=false and idpadre="&rs_settori("idsettore")&" order by ordine"
		sql_order="select settori.idsettore, settori.nome_settore,  settori.nascondi_prezzi, settori_settori.idpadre from settori_settori inner join settori on settori_settori.idfiglio = settori.idsettore where settori.nascondi=0 and settori_settori.idpadre="&rs_settori("idsettore")& " order by settori_settori.ordine"

		
		set rs_sottosettori=conn.execute(sql_order)
		NumRow=0
		do while not rs_sottosettori.eof and NumRow<MaxRow
			testo=testo&"<li><a href=""category.asp?idsettore="&rs_sottosettori("idsettore")&""">"&rs_sottosettori("nome_settore")&"</a></li>"	
			NumRow=NumRow+1
			rs_sottosettori.movenext
		loop
		if not rs_sottosettori.eof and NumRow>=MaxRow then
			testo=testo&"<li><a href=""category.asp?idsettore="&rs_settori("idsettore")&"""><strong>Vedi tutti</strong></a></li>"	
		end if
		testo=testo&"</ul></div>"
		rs_settori.movenext
		row=row+1
		if row>=5 then
			testo=testo&"</div><!-- End .row -->"
			testo=testo&"<div class=""row"">"
			row=0
			
		end if
	loop
	testo=testo&"</div><!-- End .row -->"
	testo=testo&"</div>"
	rs_settori.close
	set rs_settori=nothing
End Sub	
%>