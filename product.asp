<%
'Verifica chiusure 30_11_2015
%>
<!--#include virtual="/setup.asp" -->
<%
idpro=replace(request("idpro"),"'A=0","")
'if Request.QueryString("idsettore")<>"" then
'	idsettore=Request.QueryString("idsettore")
'else
'	sql="select settori_prodotti.* from settori_prodotti where idpro="&idpro
'	set rs_settori=conn.execute (sql)
'	idsettore=rs_settori("idsettore")
'end if
if idpro="" or not isnumeric(idpro) then
	call determina_sqlinjection(idpro)

	session("codice_errore")=11
	session("oggetto_errore")="manca idpro"
	session("pagina_errore")=questofile
	call connclose()
	response.redirect "404.asp"
end if
idpro=int(idpro)
sql="select prodotti.*, settori.nascondi_prezzi, settori_prodotti.idsettore FROM settori RIGHT JOIN (prodotti LEFT JOIN settori_prodotti ON prodotti.IDpro = settori_prodotti.IDpro) ON settori.idsettore = settori_prodotti.IDsettore WHERE prodotti.idpro=" & idpro
Set rs_prodotto = Server.CreateObject("ADODB.Recordset")
idsettore=request("idsettore")
if  not isnumeric(idsettore) then
	call determina_sqlinjection(idsettore)
	idsettore=""
end if
rs_prodotto.Open sql, conn, 3, 3
if rs_prodotto.eof then
	rs_prodotto.close
	'rs.close
	set rs_prodotto=nothing
	session("codice_errore")=9
	session("oggetto_errore")=idpro
	session("pagina_errore")=questofile
	call connclose()
	response.redirect "404.asp"
end if
if (rs_prodotto("attivo")=0 or rs_prodotto("visibilita")>2) and (not utente_admin) then
	rs_prodotto.close
	'rs.close
	set rs_prodotto=nothing
	session("codice_errore")=10
	session("oggetto_errore")=idpro
	session("pagina_errore")=questofile
	call connclose()
	response.redirect "404.asp"
end if
meta_title=rs_prodotto("articolo")
meta_description=rs_prodotto("metadescription")
meta_facebook="<meta property=""og:title"" content="""&rs_prodotto("articolo")&""" />"&vbcrlf
meta_facebook=meta_facebook&"<meta property=""og:url"" content=""http://"&Request.ServerVariables("SERVER_NAME")&"/product.asp?idpro"&rs_prodotto("idpro")&""" />"
meta_facebook=meta_facebook&"<meta property=""og:description"" content="""&meta_description&""" />"
meta_facebook=meta_facebook&"<meta property=""og:image"" content=""http://"&Request.ServerVariables("SERVER_NAME")&upl_img_cat&rs_prodotto("fileimg")&""" />"
meta_facebook=meta_facebook&"<meta property=""og:type"" content=""product"" />"
meta_facebook=meta_facebook&"<meta property=""og:site_name"" content="""&Application("brwstitle")&""" />"

%>
<!--#include virtual="/config/header_inc.asp" -->
<%

'rs_prodotto("click")=rs_prodotto("click")+1
if idsettore="" then idsettore=rs_prodotto("idcat")

'rs_prodotto.update
call logga ("pro",idpro)

var1=rs_prodotto("var1val")
var2=rs_prodotto("var2val")
sql_pulisci_magazzino="delete FROM magazzino WHERE magazzino.idpro="&idpro
set rsvar_A=conn.execute("select * from varianti_a where idpro="&idpro)
set rsvar_B=conn.execute("select * from varianti_b where idpro="&idpro)
prezzo_minimo=0
if rsvar_A.eof then
	variante_A=false
	rsvar_A.close
else
	variante_A=true
	if not isnull(rs_prodotto("prezzo_ve_min")) then prezzo_minimo=cdbl(rs_prodotto("prezzo_ve_min"))
	sql_pulisci_magazzino=sql_pulisci_magazzino&" AND magazzino.idvara=0 "
end if
if rsvar_B.eof then
	variante_B=false
	rsvar_B.close
else
	variante_B=true
	sql_pulisci_magazzino=sql_pulisci_magazzino&" AND magazzino.idvarb=0"
end if
if variante_A or variante_B then
	conn.execute sql_pulisci_magazzino,n
	if n >0 then
			add2log "Corretto errore su magazzino "&" [articolo="& idpro&"]"&rs_prodotto("codice")&"[/articolo], eliminato "&n&" record",2
	end if
end if
promozione=0
prezzo=cdbl(rs_prodotto("prezzo"))
'Gestione prezzo con iva
if rs_prodotto("promozione")=1 and (rs_prodotto("prodata")>date() or isnull(rs_prodotto("prodata"))) then
	promozione=1
	if rs_prodotto("protesto")<>"" then pro_testo="<br>"&rs_prodotto("protesto")
	if rs_prodotto("prodata")<>"" then fino_al="<br><strong>"&traduci("finoal") &" "& rs_prodotto("prodata")&"</strong> "
end if
if promozione=1 and rs_prodotto("sconto")>0 then
	txtprezzoold=simbolo_valuta&formatnumber(prezzo,2)
	prezzo=prezzo*(1-rs_prodotto("sconto")/100)
	txt_prezzo_con_iva=simbolo_valuta&formatnumber(prezzo*(1+iva(date())/100),2)
	txt_sconto=" -"&rs_prodotto("sconto")&"% "
	txtprezzo=simbolo_valuta&formatnumber(prezzo,2)
else
	if prezzo_minimo>0 then
		txtprezzo="Da "&simbolo_valuta&formatnumber(prezzo_minimo,2)
		txt_prezzo_con_iva="Da "&simbolo_valuta&formatnumber(prezzo_minimo*(1+iva(date())/100),2)
	else
		txtprezzo=simbolo_valuta&formatnumber(prezzo,2)
		txt_prezzo_con_iva=simbolo_valuta&formatnumber(prezzo*(1+iva(date())/100),2)
	end if
	txt_sconto=""
end if
prezzo_visibile=rs_prodotto("prezzo_visibile")

		'Decisionale visualizza il prezzo
		if prezzo_visibile=0 or (prezzo_visibile=2 and session("vedi_prezzi")=1) then
			'Visualizza il prezzo
			vedi_prezzo=1
			'call add2log("VEDO",0)
			
		else 
			
			'Telefonare
			vedi_prezzo=0
			'call add2log("NON VEDO",0)

		end if 
if application("scopri_prezzi")=0 and sessioniduser="" then
			vedi_prezzo=0
			'call add2log("NON VEDO PIU",0)

end if





if rs_prodotto("novita")=1 and (rs_prodotto("novdata")>date() or isnull(rs_prodotto("novdata"))) then novita=1
articolohtml=Server.HTMLEncode(rs_prodotto("articolo"&lingua))
idconsigliato=rs_prodotto("id_pro_consigliati")
if utente_andrea then response.write "[soloio]vedi_prezzo:"&vedi_prezzo&", prezzo_visibile:"&prezzo_visibile&", session(vedi_prezzi):"&session("vedi_prezzi")&"[soloio]<br>"
	%>
			<!--- Modernizr -->
        <script src="js/modernizr.custom.js"></script>

        <section id="content">
        	<div id="breadcrumb-container">
        		<div class="container">
					<ul class="breadcrumb">
						<li><a href='category.asp?idsettore=0'>Tutte le categorie</a></li>

						<%if idsettore>0 then%>
						<%=fullcat(idsettore,"",lingua)%>
						<%end if%>
					</ul>
        		</div>
        	</div>
        	<div class="container">
	        	<%if rs_prodotto("attivo")=false then%>
									<div class="alert alert-warning"  >
										<strong>Articolo non attivo,</strong> visibile solo ad admin
									</div>
				<%end if%>

        		<div class="row">
        			<div class="col-md-12">
        				<div class="row" itemscope itemtype="http://schema.org/Product">
							<div class="col-md-6 col-sm-12 col-xs-12 product-viewer clearfix" >
								<div id="product-image-carousel-container">
									<%
									sql="select * from files where idcosa="&idpro &" and cosa=4 order by ordine"
									Set rs_files = Server.CreateObject("ADODB.Recordset")
									rs_files.Open sql, conn, 1, 3
									lngTotalRecords=clng(conn.Execute("Select Found_Rows();")(0).Value)
									
									%>
									<ul id="product-carousel" class="celastislide-list <%if lngTotalRecords>4 then%>product-carousel<%end if%>">
									<%
									do while Not rs_files.EOF
									%>
										<li <%if n=1 then%> class="active-slide"<%end if%> ><a data-rel="prettyPhoto[product]" href="<%=rs_files("path")&rs_files("filename")%>" data-image="<%=rs_files("path")&rs_files("filename")%>" data-zoom-image="<%=rs_files("path")&rs_files("filename")%>" class="product-gallery-item " ><img src="<%=rs_files("path")&"s/"&rs_files("filename")%>" alt="Immagine" class="add-tooltip" title="<%=rs_files("descrizione")%>"></a></li>
										<%
										rs_files.MoveNext
									Loop
									rs_files.close
									set rs_files=nothing
										%>
									</ul><!-- End product-carousel -->
								</div>
								<div id="product-image-container">
									<figure><img itemprop="image" src="<%=upl_img_cat&rs_prodotto("fileimg")%>" alt="<%=articolohtml%>" id="product-image" data-zoom-image="<%=upl_img_cat&rs_prodotto("fileimg")%>">
										<%if false then%>
										<figcaption class="item-price-container">
											<%if promozione=1 and rs_prodotto("sconto")>0 then
											prezzoarold=split(txtprezzoold,",")%>                                                        
											<span class="old-price"><%=prezzoarold(0)%><span class="sub-price">,<%=prezzoarold(1)%></span></span><%end if%>
											<%if vedi_prezzo=1 then%>
											<%prezzoar=split(txtprezzo,",")%>
											<span class="item-price"><%=prezzoar(0)%><span class="sub-price">,<%=prezzoar(1)%></span></span>	
											<%else %>
											<span class="item-price"><span class="sub-price"><%=traduci("noprezzop")%></span></span>
											<%end if%>
										</figcaption>
										<%end if%>
										<%
										set rs_brand=Server.CreateObject("ADODB.Recordset")
										rs_brand.Open "select prodotti_brand.idpro, brand.NomeBrand, brand.ImmagineBrand FROM prodotti_brand INNER JOIN brand ON prodotti_brand.idbrand = brand.Id where idpro="&idpro, conn, 1, 3
										if not rs_brand.eof then
										%>
										<figcaption class="item-brand" itemprop="brand" content="<%=rs_brand("NomeBrand")%>">
											<img src="<%=rs_brand("ImmagineBrand")%>">
										</figcaption>
										<%
										end if
										set rs_brand = nothing
										%>
									</figure>
								</div><!-- product-image-container -->        				 
							</div><!-- End .col-md-6 -->

							<div class="col-md-6 col-sm-12 col-xs-12 product">
								<div class="lg-margin visible-sm visible-xs"></div><!-- Space -->
								<h1 class="product-name" itemprop="name"><%=articolohtml%></h1>
								<%if utente_admin then%>
											<div class="btn-group">
												<%if ha_il_permesso("A1") then %>
												<a href="pag_adm_artic.asp?modifica=<%=idpro%>" class="btn btn-sm btn-custom">Modifica</a>
												<%end if %>
												<a href="pag_adm_magazzino.asp?idpro=<%=idpro%>" class="btn btn-sm btn-custom">Giacenze</a>
												<a href="pag_adm_movimenti.asp?idpro=<%=idpro%>" class="btn btn-sm btn-custom">Movimenti</a>
												<a href="pag_adm_rep08.asp?idpro=<%=idpro%>" class="btn btn-sm btn-custom">Storico articolo</a>
											</div>
								<%end if%>
								<ul class="product-list">
									<li><span><%=traduci("idpro")%>:</span><%=rs_prodotto("idpro")%></li>
									<li><span><%=traduci("code")%>:</span><%=rs_prodotto("codice")%></li>
									<%if rs_prodotto("disponibilita")<>4 then %>
									<li><span ><%=traduci("Availability")%>:</span><%
									disponibilita_var=rs_prodotto("disponibilita")
									if disponibilita_var>=0 then
										if Application("disponibilita_auto")=1 and variante_a=false and variante_a=false then
										'Disponibilità automatica, cerco nel magazzino
										set rs_magazzino=conn.execute ("select magazzino.* from magazzino where idpro="&idpro& " and idvara=0 and idvarb=0")
										if not rs_magazzino.eof then
											if rs_magazzino("quantita_magazzino")>0 then
												'Se a magazzino imposto disponibile
												disponibilita_var=0
											else
												'Se non a magazzino imposto almeno 1 altrimenti da prodotto
												if disponibilita_var=0 then
													disponibilita_var=1
												end if
											end if 
										end if
									end if
									
									%> <span itemprop="availability"><%=disponibilita(disponibilita_var,lingua)%></span>
									  <%end if%></li>
									  <%end if %>
									<%if rs_prodotto("um")<>"" then%>
									<li><span><%=traduci("txtum")%>:</span> <%=rs_prodotto("um")%></li>
									<%end if%>
									<%if txtprezzoold<>"" then%>
									<li><span><%=traduci("txtprezzo4")%>:</span><s><%=txtprezzoold%></s></li>
									<%end if%>
									<%if promozione=1 then%>
									<li><span><%=traduci("promo")%>:</span><%=txt_sconto&pro_testo&fino_al%></li>
									<%end if%>									  
									<%
										fascia_di_sconto=rs_prodotto("fascia_di_sconto")
										if isnull(fascia_di_sconto) then fascia_di_sconto=0
									if cint(fascia_di_sconto)>0 then
										fascia_di_sconto="&nbsp;<span style=""width:10px; height:10px; background-color:"&colore_sconto(rs_prodotto("fascia_di_sconto"))&"; float:none;"">&nbsp;&nbsp;&nbsp;</span>"
									else
										fascia_di_sconto=""
									end if
									if vedi_prezzo=1 then
									if prezzo>0 then
									%>
									<li><span><%=traduci("txtprezzo1")%>:</span><span itemprop="price"><%=txtprezzo%></span><%=fascia_di_sconto%> </li>
									<li><span><%=traduci("txtprezzo3")%>:</span><%=txt_prezzo_con_iva%></li>
																	
									<%elseif prezzo_minimo>0 then%>
									<li><span><%=traduci("txtprezzo1")%>:</span><span itemprop="price"><%=txtprezzo%></span><%=fascia_di_sconto%></li>
									<li><span><%=traduci("txtprezzo3")%>:</span>Da <%=simbolo_valuta&formatnumber(prezzo_minimo*(1+iva(date())/100),2)%></li>
									<%end if
										
									else 'telefonare
										 %>
										<li><span><%=traduci("txtprezzo1")%>:</span> <%=traduci("noprezzop")%><%=fascia_di_sconto%>

										
										</li>
										
									<%

										
										end if%>
										<%if rs_prodotto("flag")=1 then%>
										<li style="background:#00FF40; width:100%;">&nbsp;</li>
										<%end if%>
										
										<%
										if utente_admin then %>
										<hr>
										<li class="admin"><strong>Dati visibili solo admin:</strong></li>
										<li class="admin">Visibilità prezzo:<%
											select case rs_prodotto("prezzo_visibile")
												case 0
													response.write "visibile"
												case 1
													response.write "Nascosto a tutti"
												case 2
													response.write "Visibile solo ad utenti abilitati"
												end select
											%></li>
											<li>Prezzo: <%=simbolo_valuta& formatnumber(rs_prodotto("prezzo"),2)%>
											<li>Prezzo con iva: <%=simbolo_valuta& formatnumber(cdbl(rs_prodotto("prezzo"))*1.22,2)%>
										<li class="admin"><span>Fornitore:</span>
										<%
											if rs_prodotto("idfor")>0 then 
												response.write get_denominazione(rs_prodotto("idfor"))
											end if
											
											
											
											%> </li>
										
										
										<%end if
										
										%>
										

								</ul>

								<hr>
								<%if variante_a then%>
								<div class="product-color-filter-container">
									<span>Seleziona <%=rs_prodotto("variante1")%>:</span>
									<div class="xs-margin"></div>
									<ul class="filter-size-list clearfix">
									<%
									do while not rsvar_A.eof
									%>
									<li class="active"><a href="#" id="vara_<%=rsvar_a("idvara")%>" class="vara" price="<%=rsvar_a("prezzo_ve_va")%>"><%
											response.write codice_e_variante(rsvar_a("codicevara"),rsvar_a("variante_a"))  
											if prezzo_minimo>0 and vedi_prezzo=1 then
												if cdbl(rsvar_A("prezzo_ve_va"))=0 then
													response.write " - &#8364; "&formatnumber(prezzo,2)
												else
													response.write " - &#8364; "&formatnumber(rsvar_A("prezzo_ve_va"),2)
												end if
											end if
											%></a></li><%
										rsvar_A.movenext
									loop
								%>
									</ul>
								</div><!-- End .product-color-filter-container-->
						  
								<%end if
									set rsvar_A = Nothing
								%>
								<%if variante_B then%>
								<div class="product-size-filter-container">
									<span>Seleziona <%=rs_prodotto("variante2")%>:</span>
									<div class="xs-margin"></div>
									<ul class="filter-size-list clearfix">
									<%
									do while not rsvar_b.eof
									%>
									<li><a href="#" id="varb_<%=rsvar_b("idvarb")%>" class="varb"><%=rsvar_b("variante_b")%></a></li><%
										rsvar_b.movenext
										loop
										
										%>
									</ul>
								</div><!-- End .product-size-filter-container-->
								<%end if
								set rsvar_b = Nothing
								%>
								<%if variante_A or variante_B then%>
								<hr>
							<%end if%>
							<%if  rs_prodotto("attivo") and rs_prodotto("vendita") and vedi_prezzo=1 and application("consenti_acquisti")=1 then%>

								<div class="product-add clearfix" id="prova2">
									<div class="custom-quantity-input">
										<input type="text" name="quantity" value="1" id="quantity">
										<a href="#" onclick="return false;" class="quantity-btn quantity-input-up" id="incrementa_quantity"><i class="fa fa-angle-up"></i></a>
										<a href="#" onclick="return false;" class="quantity-btn quantity-input-down" id="decrementa_quantity"><i class="fa fa-angle-down"></i></a>
									</div>
									<button class="btn btn-custom-2" id="aggiungi"><%=traduci("addcart2")%></button>
								</div><!-- .product-add -->
								<div id="alert-aggiunto" style="display: none;">
									<div class="md-margin"></div><!-- Space -->
									<div class="alert alert-success"  >
										<strong><%=traduci("altit")%></strong>&nbsp; <%=traduci("altxt")%>
									</div>
								</div>
								<div class="md-margin"></div><!-- Space -->
								<%end if%>
								<div class="product-extra clearfix">
									<%if session("iduser")<>"" then%>
									<div class="product-extra-box-container clearfix">
										<div class="item-action-inner">
										<%
										set rs_preferiti=conn.execute ("select preferiti.* from preferiti where idpro="&idpro&" and iduser="&session("iduser"))
										if not rs_preferiti.eof then
											classe_preferiti=" active"
											txt_btn_pref="Rimuovi da"
										else	
											txt_btn_pref="Aggiungi a"
										end if
										set rs_preferiti=nothing
										%>
											<a id="favourite" href="#" class="btn btn-custom <%=classe_preferiti%>"><%=txt_btn_pref%> articoli preferiti</a>
											<%if utente_admin and rs_prodotto("idfor")>0 then %>
											<a id="ordina_a_for" href="#" class="btn btn-custom">Ordina a fornitore</a>
											<div id="Modalordina" class="modal fade">
											    <div class="modal-dialog">
											        <div class="modal-content" style="transform: translate(0, 50%) !important; -ms-transform: translate(0, 50%) !important;-webkit-transform: translate(0, 50%) !important;">
											            <div class="modal-header">
											                <h4 class="modal-title">Sei sicuro?</h4>
											            </div>
											            <div class="modal-body">
											                <p>Vuoi aggiungere questo articolo all'ordine fornitore?</p>
											            </div>
											            <div class="modal-footer">
											                <button type="button" class="btn btn-primary" data-dismiss="modal" id="ordina_a_for_conferma">SI</button><button type="button" class="btn btn-primary" data-dismiss="modal">NO</button>
											            </div>
											        </div>
											    </div>
											</div>
											<%end if %>
										</div><!-- End .item-action-inner -->
									</div>
									<%end if%>
									<%if false then
										'if not mod_larochelle then %>
									<div class="share-button-group">
										<!-- AddThis Button BEGIN -->
										<div class="addthis_toolbox addthis_default_style addthis_32x32_style" data-title="<%=rs_prodotto("fileimg")%>">
										<a class="addthis_button_facebook"></a>
										<a class="addthis_button_twitter"></a>
										<a class="addthis_button_email"></a>
										<a class="addthis_button_print"></a>
										<a class="addthis_button_compact"></a><a class="addthis_counter addthis_bubble_style"></a>
										</div>
										<script type="text/javascript">var addthis_config = {data_track_addressbar:true, ui_language: "it"};</script>
										<script type="text/javascript" src="//s7.addthis.com/js/300/addthis_widget.js#pubid=ra-52b2197865ea0183"></script>
										<!-- AddThis Button END -->
									</div><!-- End .share-button-group -->
									<%end if %>
								</div><!-- End .product-extra -->
								<div id="alert-add-pre" style="display: none;">
									<div class="md-margin"></div><!-- Space -->
									<div class="alert alert-success"  >
										<strong><%=traduci("preaddtit")%></strong>&nbsp; <%=traduci("preaddtxt")%>
									</div>
								</div>
								<div id="alert-add-remove" style="display: none;">
									<div class="md-margin"></div><!-- Space -->
									<div class="alert alert-success"  >
										<strong><%=traduci("preremtit")%></strong>&nbsp; <%=traduci("preremtxt")%>
									</div>
								</div>
								<div id="alert-add-order" style="display: none;">
									<div class="md-margin"></div><!-- Space -->
									<div class="alert alert-success"  >
										<strong>Articolo aggiunto all'ordine fornitore <span id="nord"></span></strong><br><a href="#" id="nord_href">Apri ordine forniore</a>
									</div>
								</div>
							</div><!-- End .col-md-6 -->
        				</div><!-- End .row -->
        				
        				<div class="lg-margin"></div><!-- End .space -->
        				<div class="row">
        					<div class="col-md-12 col-sm-12 col-xs-12">
        						
        						<div class="tab-container left product-detail-tab clearfix">
                <!-- TITOLI PANNELLI   -->
									<ul class="nav-tabs">
										<li class="active"><a href="#description" data-toggle="tab"><%=traduci("desc")%></a></li>
										<%
											set rs_modelli=conn.execute ("select marca_modello.Nome, prodotti_marca_modello.idpro, marca_modello.Id, marca_modello.IdPadre FROM prodotti_marca_modello INNER JOIN marca_modello ON prodotti_marca_modello.idmodello = marca_modello.Id where idpro="&idpro)
											if not rs_modelli.eof then
											%>
											<li><a href="#modelli" data-toggle="tab"><%=traduci("comptit")%></a></li>
											<%end if%>
										<%
										'Pulsanti di selezione dettaglio
										'Dettagli aggiuntivi
										sql_dettagli="select * from prodotti_dettagli where idpro="&idpro &" order by ordine"
										set rs_dettagli=conn.execute(sql_dettagli)
										do while not   rs_dettagli.eof
											%>
											<li><a href="#dettaglio<%=rs_dettagli("ordine")%>" data-toggle="tab"><%=rs_dettagli("titolo"&lingua)%></a></li>
											<%
											rs_dettagli.movenext
										loop
										set rs_dettagli=nothing

										'files allegati
											%>
											  <li><a href="#download" data-toggle="tab"><%=traduci("scarica")%></a></li>
											<%
										if utente_admin then%>
											<li><a href="pag_adm_dettagli.asp?oper=new&idpro=<%=idpro%>&goback=si" title="Aggiungi dettaglio"><i class="fa fa-plus"></i></a></li>
										<%end if
										%>
										</ul>
										
                <!-- CONTEUTI PANNELLI   -->
        								<div class="tab-content clearfix">
                <!-- PANNELLO DESCRIZIONE   -->
        									<div class="tab-pane active" id="description"><%if utente_admin then%><a href="pag_adm_artic.asp?modifica=<%=idpro%>" title="Modifica articolo" class="edit-item pull-right" style="margin-right: -30px;"><i class="fa fa-pencil"></i></a><%end if%>
												<p><%=rs_prodotto("descrizione"&lingua)%></p>
        									</div><!-- End .tab-pane -->
        									
											<%if not rs_modelli.eof then
											%>
                <!-- PANNELLO COMPATIBILE   -->
											<div class="tab-pane" id="modelli">
												<p><%=traduci("comptxt")%></p>
												<ul class="product-details-list">
												<%do while not rs_modelli.eof%>
												<li><%=gerarchia_modelli_bkw(rs_modelli("id"),"")%></li> 
												<%
													rs_modelli.MoveNext
													Loop
													%>
												</ul>
        									</div><!-- End .tab-pane -->
   											<%end if
	   											
	   											set rs_modelli=Nothing
   											%>
        								   <%
                                             set rs_dettagli=conn.execute(sql_dettagli)
                                             do while not rs_dettagli.eof
                                           %>
                <!-- PANNELLI DETTAGLI   -->
	        									<div class="tab-pane" id="dettaglio<%=rs_dettagli("ordine")%>"><%if utente_admin then%><a href="pag_adm_dettagli.asp?modifica=<%=rs_dettagli("id")%>&goback=si" title="Modifica dettaglio" class="edit-item pull-right" style="margin-right: -30px;"><i class="fa fa-pencil"></i></a><%end if%>
        										   <%=rs_dettagli("testo"&lingua)%>
	        									</div><!-- End .tab-pane -->
                                             <%
                                              rs_dettagli.movenext
                                              loop
                                              set rs_dettagli=Nothing
                                               %>
	        									<div class="tab-pane" id="download">
													<p><%=traduci("downtxt")%></p>
	                                            
												<ul class="product-details-list">
												<%if session("iduser")<>"" then %>
												<li><a href="pdf_dettaglio.asp?idpro=<%=idpro%>"><%=traduci("scheda")%>&nbsp;<%=rs_prodotto("codice")%></a></li><%
else
%>												<li><%=traduci("scheda")%>&nbsp;<%=rs_prodotto("codice")%></li>
												<%end if%>
												<%
												sql="select * from files where idcosa="&idpro &" and cosa=2"
												set rs_files=conn.execute(sql)
												do until rs_files.EOF
													if session("iduser")<>"" then
												%>
	                                                <li><a href="<%=rs_files("path")&rs_files("filename")%>" target="blank"><%=rs_files("filename")%></a></li>
	                                                <%
													else
													%>
													<li><%=rs_files("filename")%></li>
													<%
													end if
	                                                rs_files.movenext
	                                                loop
	                                                rs_files.close
	                                                set rs_files=Nothing
	                                                %>
												
												</ul>
												<%if session("iduser")="" then %>
												<div class="md-margin"></div><!-- Space -->
												<p><%=traduci("nodown")%>
												<%end if%>
	        									</div><!-- End .tab-pane -->
												
        									

        								</div><!-- End .tab-content -->
        						</div><!-- End .tab-container -->
        						<div class="lg-margin visible-xs"></div>
        					</div><!-- End .col-md-12 -->

        				</div><!-- End .row -->
						
						
        				<div class="lg-margin2x"></div><!-- Space -->
						<%
						if not isnull(idconsigliato) then
						set rs_articoli=conn.execute("select prodotti.*, prodotti_consigliati.idconsigliato FROM prodotti_consigliati INNER JOIN prodotti ON prodotti_consigliati.idpro = prodotti.IDpro  where attivo=true and  prodotti_consigliati.idconsigliato="&idconsigliato)
						if not rs_articoli.eof then
						%>

                <!-- CONSIGLIATI SLIDER CONSIGLIATI  CONSIGLIATI  CONSIGLIATI  CONSIGLIATI  CONSIGLIATI  CONSIGLIATI -->
						<div class="purchased-items-container carousel-wrapper">
                            <header class="content-title">
                                <div class="title-bg">
                                    <h2 class="title"><%=traduci("tit11")%></h2>
                                </div><!-- End .title-bg -->
                                <p class="title-desc"><%=traduci("tit11txt")%></p>
                            </header>

                            <div class="carousel-controls">
                                <div id="purchased-items-slider-prev" class="carousel-btn carousel-btn-prev"></div><!-- End .carousel-prev -->
                                <div id="purchased-items-slider-next" class="carousel-btn carousel-btn-next carousel-space"></div><!-- End .carousel-next -->
                            </div><!-- End .carousel-controllers -->
							
                            <div class="purchased-items-slider owl-carousel">
                            <%
							
									response.write box_rs_articoli(rs_articoli,0)
                            %>

                            </div><!--purchased-items-slider -->
                        </div><!-- End .purchased-items-container -->
                            <%
                            end if
							Set rs_articoli=Nothing

                            end if
                            if idsettore<>"" then
                            
							set rs_articoli=conn.execute("select prodotti.*, settori_prodotti.IDsettore FROM prodotti INNER JOIN settori_prodotti ON prodotti.IDpro = settori_prodotti.IDpro WHERE attivo=true and settori_prodotti.IDsettore="&idsettore&" LIMIT 0,12")
						if not rs_articoli.eof then
	                        %>
                            
                <!-- CATEGORIE SLIDER CATEGORIE  CATEGORIE  CATEGORIE  CATEGORIE  CATEGORIE  CATEGORIE  CATEGORIE -->
        				<div class="categoria-items-container carousel-wrapper">
                            <header class="content-title">
                                <div class="title-bg">
                                    <h2 class="title"><%=traduci("tit10")%></h2>
                                </div><!-- End .title-bg -->
                                <p class="title-desc"><%=traduci("tit10txt")%></p>
                            </header>
                            <div class="carousel-controls">
                                <div id="categoria-items-slider-prev" class="carousel-btn carousel-btn-prev"></div><!-- End .carousel-prev -->
                                <div id="categoria-items-slider-next" class="carousel-btn carousel-btn-next carousel-space"></div><!-- End .carousel-next -->
                            </div><!-- End .carousel-controllers -->
                            <div class="categoria-items-slider owl-carousel">
                                    <%
									response.write box_rs_articoli(rs_articoli,0)
                                    %>
                            </div><!--categoria-items-slider -->
                        </div><!-- End .categoria-items-container -->
                        <%
	                        end if
							Set rs_articoli=Nothing
							end if 'idsettore<>"" 
	                        %>

        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
			</div><!-- End .container -->
        </section><!-- End #content -->
        
<div id="ModalvarA" class="modal fade">
    <div class="modal-dialog">
        <div class="modal-content" style="transform: translate(0, 50%) !important; -ms-transform: translate(0, 50%) !important;-webkit-transform: translate(0, 50%) !important;">
            <div class="modal-header">
                <h4 class="modal-title">Attenzione</h4>
            </div>
            <div class="modal-body">
                <p>Devi selezionare la variante <strong><%=ucase(rs_prodotto("variante1"))%></strong></p>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-primary" data-dismiss="modal">OK</button>
            </div>
        </div>
    </div>
</div>
<div id="ModalvarB" class="modal fade">
    <div class="modal-dialog">
        <div class="modal-content" style="transform: translate(0, 50%) !important; -ms-transform: translate(0, 50%) !important;-webkit-transform: translate(0, 50%) !important;">
            <div class="modal-header">
                <h4 class="modal-title">Attenzione</h4>
            </div>
            <div class="modal-body">
                <p>Devi selezionare la variante <strong><%=ucase(rs_prodotto("variante2"))%></strong></p>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-primary" data-dismiss="modal">OK</button>
            </div>
        </div>
    </div>
</div>

<%
	codice=rs_prodotto("codice")
	idfor=rs_prodotto("idfor")
	set rs_prodotto=Nothing
	
	
%>
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
    <%page_script="product"%>
	<!--#include virtual="/script_inc.asp" -->
	
	<script>
	var idvara;
	var idvarb;
	var prezzo;
	var idpro=<%=idpro%>;
	var codice="<%=codice%>";
	<%if session("iduser")<>"" then%>
	var iduser=<%=session("iduser")%>;
	<%else%>
	var iduser=null;
	<%end if%>
	<%if SessionIdGuest<>"" then%>
	var idloged=<%=SessionIdGuest%>;
	<%else%>
	var idloged=null;
	<%end if%>
		$(function() {
			$.ajaxSetup({
				// Disable caching of AJAX responses
				cache: false
			});	
		   //mio codice
			$("#incrementa_quantity").click(function() {
				var quantity=parseInt($("#quantity").val());
				quantity=quantity+1;
				$("#quantity").val(quantity);	
				
			});
			$("#decrementa_quantity").click(function() {
				var quantity=parseInt($("#quantity").val());
				quantity=quantity-1;
				if(quantity<1){quantity=1;}
				$("#quantity").val(quantity);	
				
			});			
			$("#aggiungi").click(function(e) {
				$("#alert-aggiunto").hide();
				console.log("aggiungi");
				<%if variante_a then%>
				if(!idvara){
					//alert("seleziona variante");
			        $("#ModalvarA").addClass('in').modal('show');
                    console.log("manca variante A");
			        return;
				}
                <%end if%>
				<%if variante_b then%>
				if(!idvarb){
					//alert("seleziona variante");
			        $("#ModalvarB").addClass('in').modal('show');
                    console.log("manca variante A");
			        return;
				}
				<%end if%>
				quantita=$("#quantity").val();
				$.get("cart_ajax.asp?add=si&idpro=" + idpro + "&idvara=" + idvara+ "&idvarb=" + idvarb+"&quantita=" + quantita+"&codice="+codice+"&iduser="+iduser+"&idloged="+idloged, function (result) {
					$("#carrellino").html(result);
					$("#alert-aggiunto").slideDown();
					$(".vara").removeClass("myhover");
					$(".varb").removeClass("myhover");
			        
				})

				  .fail(function() {
				    alert( "error" );
				  });
				
				
			});
			<%if utente_admin and idfor>0 then %>
			$("#ordina_a_for").click(function(e) {
				<%if variante_a then%>
				if(!idvara){
					//alert("seleziona variante");
			        $("#ModalvarA").addClass('in').modal('show');
                    console.log("manca variante A");
			        return;
				}
                <%end if%>
				<%if variante_b then%>
				if(!idvarb){
					//alert("seleziona variante");
			        $("#ModalvarB").addClass('in').modal('show');
                    console.log("manca variante A");
			        return;
				}
				<%end if%>
				$("#Modalordina").addClass('in').modal('show');
			});
			
			$("#ordina_a_for_conferma").click(function(e) {

				quantita=$("#quantity").val();
				
				
				$.ajaxSetup({ cache: false });

			    $.ajax({
					url     : "ajax_function.asp?oper=ordina_for&idpro=" + idpro + "&idvara=" + idvara+ "&idvarb=" + idvarb+"&quantita=" + quantita+"&idfor=<%=idfor%>&tabella=ordini_fornitori",
					type    : "post",
					dataType: 'json',
					//data	: order,
					success: function(data){
						//alert(data.Message);
						$("#nord").html(data.nord);
						$("#nord_href").attr("href","pag_adm_ordini_fornitori.asp?idord="+data.idord);
						$("#alert-add-order").slideDown();
						$(".vara").removeClass("myhover");
						$(".varb").removeClass("myhover");
						
					}
					,error:function(xhr, textStatus, error){
					      console.log("xhr.statusText:"+xhr.statusText);
					      console.log("xhr.responseText:"+xhr.responseText);
					      console.log("textStatus:"+textStatus);
					      console.log("error:"+error);
						  }
				});
			});
			<%end if %>

			
			
			<%if session("iduser")<>"" then%>
			$("#favourite").click(function(e) {
				e.preventDefault();
				$.get("searcher.asp?oper=favourite&idpro=" + idpro , function (result) {
					if (result.success==true)
					{
						if (result.action=="add")
						{
							$("#favourite").addClass("active");
							$("#alert-add-remove").hide();
							$("#alert-add-pre").slideDown();
						}
						else
						{
							$("#favourite").removeClass("active");
							$("#alert-add-pre").hide();
							$("#alert-add-remove").slideDown();
						}

					}
					else
					{
						alert(result.Message);
					}
				})
			  .fail(function() {
				    alert( "error" );
				  });
			});
			<%end if%>
			
			$(".delete-item").click(function(){
				var idcar=$(this).closest('li').attr('id').replace('item-car-sm_', '');
				console.log("Elimina IDcar:"+idcar);
				
			})
			
			$(".vara").click(function(e) {
				e.preventDefault();
				$(".vara").removeClass("myhover");
				$(this).addClass("myhover");
				idvara=$(this).attr("id").replace("vara_", "");
				console.log("idvariantea:"+idvara);				
				
			})
			$(".varb").click(function(e) {
				e.preventDefault();
				$(".varb").removeClass("myhover");
				$(this).addClass("myhover");
				idvarb=$(this).attr("id").replace("varb_", "");
				console.log("idvarianteb:"+idvarb);				
				
			})
			
			//fine mio codice
			
			

		});
	</script>

    </body>
</html>
<!--#include virtual="/dettaglio_inc.asp" -->
<%
function fullcat(id,tmp,lingua)
	set rs_settore=conn.execute ("select settori.idsettore, settori.nome_settore, settori_settori.idpadre FROM settori inner join settori_settori on settori.idsettore =  settori_settori.idfiglio WHERE idsettore = " & id&";")
	if not rs_settore.eof then
		if rs_settore("idsettore")<>cint(Application("settore_sup")) then  
			if tmp="" then txtclass="class=""active"""
			str="<li><a href='category.asp?idsettore="&rs_settore("idsettore")&"'>"&rs_settore("Nome_Settore"&lingua)&"</a></li>"
			str=str& tmp
			tmp=fullcat(rs_settore("idpadre"),str,lingua)
		end if
	end if
	fullcat=tmp
	set rs_settore=nothing
End function

%>