<%
'Verifica chiusure 30_11_2015
%>
<!--#include virtual="/setup.asp" -->
<%
if utente_user then
	where_carrello="iduser="&sessionIDUser
	if sessionIDUser="" then
		'Verifica temporanea 02/12/2015, poi togliere
		if SessionIdGuest="" then
			add2log "sessionIDUser non impostato",0
		end if
	end if
else
	where_carrello="idloged="&SessionIdGuest
end if
' Aggiorna carrello
IF Request( "update_cart" ) <> "" THEN
	sql = "select carrello.* FROM carrello WHERE "&where_carrello
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	WHILE NOT RS.EOF
	newQ = TRIM( Request( "quantity_" & RS( "idcar" ) ) )
	IF newQ = "" OR newQ = "0" or TRIM(Request("elimina"&RS( "idcar")))="SI" THEN
		sql = "select prodotti.idpro, prodotti.codice FROM prodotti WHERE idpro=" & rs("idpro")
		set rs_prodotti=conn.execute (sql)
		add2log "[articolo="& rs_prodotti("idpro")&"]"&rs_prodotti("codice")&"[/articolo] rimosso dal carrello da "&testoDa()&" iduser:"&rs("iduser")&" idloged:"&rs("idloged")&" id:"&rs("idcar"),1
		set rs_prodotti=Nothing
		RS.Delete
	ELSE
	  IF isNumeric( newQ ) THEN
		RS("quantita") = newQ
		rs.update
	  END IF
	END IF
	RS.MoveNext
	WEND
	RS.Close
	set rs=Nothing
	session("cache_carrellino")=""
END IF
IF Request( "cart_quantity_delete" ) <> "" THEN
	sql = "select * FROM carrello WHERE idcar=" & request("cart_quantity_delete")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	sql = "select prodotti.idpro, prodotti.codice FROM prodotti WHERE idpro=" & rs("idpro")
	set rs_prodotti=conn.execute (sql)
	add2log "[articolo="& rs_prodotti("idpro")&"]"&rs_prodotti("codice")&"[/articolo] rimosso dal carrello da "&testoDa()&" idcar:"&rs("idcar")&" iduser:"&rs("iduser")&" idloged:"&rs("idloged"),1
	set rs_prodotti=Nothing
	rs.Delete
	rs.close
	set rs=Nothing
	session("cache_carrellino")=""
END IF
function testoDA()
	if utente_user then
	testoDa="[utente="&session("iduser")&"]"&session("nominativo")&"[/utente]"
else
	testoDa="utente non loggato"
end if
end function
'meta_title=traduci("carrello")
%>
<!--#include virtual="/config/header_inc.asp" -->
<%
'on error resume next	
%>
        <section id="content">
        	<div id="breadcrumb-container">
        		<div class="container">
					<ul class="breadcrumb">
						<li><a href="/">Home</a></li>
						<li class="active"><%=traduci("carrello")%></li>
					</ul>
        		</div>
        	</div>
        	<div class="container">
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=traduci("carrello")%></h1>
							<p class="title-desc"><%=traduci("cartxt")%></p>
						</header>
        				<div class="xs-margin"></div><!-- space -->
        				<div class="row">
        					
        					<div class="col-md-12 table-responsive" id="order-detail-content">
								<form name="form1" id="form1" method="post" action="<%=questofile%>">  
								<input type="hidden" value="update_cart" name="update_cart"/>
        						<table class="table cart-table" >
        						<thead>
        							<tr>
										<th class="table-title"><%=traduci("artic")%></th>
										<th class="table-title"><%=traduci("code")%></th>
										<th class="table-title"><%=traduci("priceun")%></th>
										<th class="table-title"><%=traduci("quant")%></th>
										<th class="table-title"><%=traduci("subtot")%></th>
        							</tr>
        						</thead>
								<tbody>
								<%
								tot_carrello=0
    							sql="select carrello.*, prodotti.codice, prodotti.articolo, prodotti.variante1, prodotti.variante2, prodotti.prezzo, prodotti.costo, prodotti.Promozione, prodotti.Sconto, prodotti.Prodata, varianti_a.variante_a, varianti_a.prezzo_ve_va, varianti_b.variante_b, prodotti.idfor, prodotti.fileimg, prodotti.aggiungi_a_ordine FROM varianti_b RIGHT JOIN (varianti_a RIGHT JOIN (carrello INNER JOIN prodotti ON carrello.idpro = prodotti.IDpro) ON varianti_a.IDvara = carrello.idvara) ON varianti_b.IDvarb = carrello.idvarb where "&where_carrello
								Set rs_carrello=conn.Execute(sql)
								do while not rs_carrello.EOF
								if converti_typevar14(rs_carrello("prezzo_ve_va"))>0 then
									prezzo=rs_carrello("prezzo_ve_va") 
								else
									prezzo=rs_carrello("prezzo")
								end if
								
								prezzo=cdbl(prezzo)
								'Calcolo lo sconto								
								if rs_carrello("promozione")=1 and rs_carrello("sconto")>0 then
									prezzoeff=prezzo*(1-rs_carrello("sconto")/100)
								else
									prezzoeff=prezzo
								end if
								
								'Eseguo il conteggio
								subtot=totale_riga(prezzoeff,0,rs_carrello("quantita"))
								subtoteff=subtot
								'Aggiungo iva per visualizzazione								
								if Application("prezzi_con_iva")>0 then
									prezzoeff=applica_iva(prezzo)
									subtoteff=applica_iva(subtoteff)
								end if
								'Preparo la stringa
								if rs_carrello("promozione")=1 and rs_carrello("sconto")>0 then
									txtprezzo="<s>"&simbolo_valuta&formatnumber(prezzo,2)&"</s><br><b>" & formatnumber(prezzoeff,2)&"</b>"
								else
									txtprezzo=simbolo_valuta&formatnumber(prezzoeff,2)
								end if
								
								
								
								%>
									<tr>
										<td class="item-name-col">
											<figure>
												<a href="product.asp?idpro=<%=rs_carrello("idpro")%>"><img src="send_img.asp?imgs1=<%=rs_carrello("fileimg")%>" alt="<%=rs_carrello("articolo")%>"></a>
											</figure>
											<header class="item-name"><a href="product.asp?idpro=<%=rs_carrello("idpro")%>"><%=rs_carrello("articolo")%></a></header>
											 <%
											'varianti 
											txt=""
											if rs_carrello("idvara")<>""  and rs_carrello("idvara")<>0 then
												txt ="<li>" & rs_carrello("variante1") & ": " &  rs_carrello("variante_a")&"</li>"
											end if
											if rs_carrello("idvarb")<>"" and rs_carrello("idvarb")<>0 then
												txt =txt & "<li>" & rs_carrello("variante2") & ": "  &  rs_carrello("variante_b")&"</li>"
											end if
											if txt<>"" then%><ul><%=txt%></ul><%end if%>
										</td>
										<td class="item-code"><%=rs_carrello("codice")%></td>
										<td class="item-price-col"><span class="item-price-special"><%=simbolo_valuta&FormatNumber(prezzoeff,2)%></span></td>
										<td>
											<div class="custom-quantity-input">
												<input type="text" class="cart_quantity_input" id="quantity_<%=rs_carrello("idcar")%>" name="quantity_<%=rs_carrello("idcar")%>" value="<%=rs_carrello("quantita")%>">
												<a href="#" onclick="return false;" class="quantity-btn quantity-input-up" id="cart_quantity_up_<%=rs_carrello("idcar")%>"><i class="fa fa-angle-up"></i></a>
												<a href="#" onclick="return false;" class="quantity-btn quantity-input-down" id="cart_quantity_down_<%=rs_carrello("idcar")%>"><i class="fa fa-angle-down"></i></a>
											</div>
										</td>
										<td class="item-total-col"><span class="item-price-special"><%=simbolo_valuta&FormatNumber(subtoteff,2)%></span>
										<a href="<%=questofile%>?cart_quantity_delete=<%=rs_carrello("idcar")%>" class="close-button cart_quantity_delete"></a>
										</td>
									</tr>
									<%
									idpro=rs_carrello("idpro")
									tot_carrello=tot_carrello+subtot

									rs_carrello.MoveNext
									loop
									rs_carrello.Close
									set rs_carrello=nothing
									if Application("prezzi_con_iva")>0 then
										tot_carrello=applica_iva(tot_carrello)
									end if
							
									
									%>
									
									
								</tbody>
							  </table>
							  <div id="update_cart" style="display:none;">
	        					<div class="md-margin"></div><!-- Space -->
							          				<a href="#" class="btn btn-custom pull-right"  ><%=traduci("upcart")%></a>
							          				
							  </div>
								</form>

        					</div><!-- End .col-md-12 -->
        					
        				</div><!-- End .row -->
        				<div class="lg-margin"></div><!-- End .space -->
        				<div class="row">
        				<%if false then%>
        					<div class="col-md-8 col-sm-12 col-xs-12">
        						
        						<div class="tab-container left clearfix">
        								<ul class="nav-tabs">
										  <li class="active"><a href="#shipping" data-toggle="tab">Shipping &amp; Taxes</a></li>
										  <li><a href="#discount" data-toggle="tab">Discount Code</a></li>
										  <li><a href="#gift" data-toggle="tab">Gift voucher </a></li>
										  
										</ul>
        								<div class="tab-content clearfix">
        									<div class="tab-pane active" id="shipping">
        										
        										<form action="#" id="shipping-form">
        											<p>Enter your destination to get a shipping estimate.</p>
                                                    <div class="xs-margin"></div>
													<div class="form-group">
														<label for="select-country" class="control-label">Country&#42;</label>
														<div class="input-container">
                                                            <select name="select-country" class="form-control" id="select-country">
                                                                <option value="Italy">Italy</option>
                                                                <option value="France">France</option>
                                                                <option value="Spain">Spain</option>
                                                                <option value="Brazil">Brazil</option>
                                                            </select>
                                                        </div><!-- End .select-container -->
													</div><!-- End .form-group -->
													<div class="sm-margin"></div>
													<div class="form-group">
                                                        <label for="select-state" class="control-label">Regison&amp;State&#42;</label>
                                                        <div class="input-container">
                                                            <select name="select-state" class="form-control" id="select-state">
                                                                <option value="Italy">Italy</option>
                                                                <option value="France">France</option>
                                                                <option value="Spain">Spain</option>
                                                                <option value="Brazil">Brazil</option>
                                                            </select>
                                                        </div><!-- End .select-container -->
                                                    </div><!-- End .form-group -->
        										  <div class="sm-margin"></div>
        										<div class="form-group">
													<label for="select-country" class="control-label"  >Post Code&#42;</label>
													<div class="input-container">
                                                        <input type="text" required class="form-control" placeholder="Your fax">
                                                    </div>
												</div><!-- End .form-group -->
        										<div class="sm-margin"></div>
        										<p class="text-right">
        											<input type="submit" class="btn btn-custom-2" value="GET QUOTES">
        										</p>
        										</form>
        										
        									</div><!-- End .tab-pane -->
        									
        									<div class="tab-pane" id="discount">
        										<p>Enter your discount coupon code here.</p>
        										<form action="#">
        											<div class="input-group">
														<input type="text" required class="form-control" placeholder="Coupon code">
														
													</div><!-- End .input-group -->	
        										<input type="submit" class="btn btn-custom-2" value="APPLY COUPON">
        										</form>
        									</div><!-- End .tab-pane -->
        									
        									<div class="tab-pane" id="gift">
        										<p>Lorem ipsum dolor sit amet, consectetur adipisicing elit. Sequi dignissimos nostrum debitis optio molestiae in quam dicta labore obcaecati ullam necessitatibus animi deleniti minima dolor suscipit nobis est excepturi inventore.</p>
        									</div><!-- End .tab-pane -->
        									
        								</div><!-- End .tab-content -->
        						</div><!-- End .tab-container -->
        						
        					</div><!-- End .col-md-8 -->
        					<%end if%>
							<div class="lg-margin visible-sm visible-xs"></div><!-- space -->
        					<div class="col-md-4 col-sm-12 col-xs-12 pull-right">
        						
        						<table class="table total-table nascondi">
        							<tbody>
        								<%if false then%>
        								<tr>
        									<td class="total-table-title"><%=traduci("subtot")%>:</td>
        									<td><%=FormatNumber(tot_carrello,2)%></td>
        								</tr>
        								<tr>
        									<td class="total-table-title">Shipping:</td>
        									<td>$6.00</td>
        								</tr>
        								<%end if%>
        								<%if false then%>
        								<tr>
        									<td class="total-table-title">TAX (0%):</td>
        									<td>$0.00</td>
        								</tr>
        								<%end if%>
        							</tbody>
        							
        							<tfoot>
        								<tr>
											<td><%=traduci("tot")%>:</td>
											<td><%=simbolo_valuta&FormatNumber(tot_carrello,2)%></td>
        								</tr>
        							</tfoot>
        						</table>
								</div><!-- End .col-md-4 -->
        				</div><!-- End .row -->
						<div class="row pull-right">

								<div class="md-margin"></div><!-- End .space -->

								<div class="col-md-12 col-sm-12 col-xs-12">
        						<a href="/" class="btn btn-custom-2"><%=traduci("contshop")%></a>
								
        						<%if cdbl(tot_carrello)>=cdbl(Application("minimo_preventivo")) then%>
        						<a href="checkout.asp?preventivo=true" class="btn btn-custom nascondi"><%=traduci("preventivo")%></a>
        						<%end if%>
        						<%if tot_carrello>0 then%>
        						<a href="checkout.asp" class="btn btn-custom nascondi"><%=traduci("checkout")%></a>
        						<%end if%>
								</div><!-- End .col-md-12 -->
        				</div><!-- End .row -->
						<%if idpro>0 then%>
        				<div class="lg-margin2x"></div><!-- Space -->
        				<%
	        				
                      set rs_settore=conn.execute("select settori_prodotti.IDsettore FROM settori_prodotti WHERE (settori_prodotti.IDpro)="&idpro)
                      idsettore=rs_settore("idsettore")
                      set rs_settore=nothing
                      set rs_articoli=conn.execute("select prodotti.*, settori_prodotti.IDsettore FROM prodotti INNER JOIN settori_prodotti ON prodotti.IDpro = settori_prodotti.IDpro WHERE (settori_prodotti.IDsettore)="&idsettore&" LIMIT 0,15")
					if not rs_articoli.eof then
	        				
	        				%>
        				
        				
        				
        				<div class="similiar-items-container carousel-wrapper">
                            <header class="content-title">
                                <div class="title-bg">
                                    <h2 class="title"><%=traduci("simi")%></h2>
                                </div><!-- End .title-bg -->
                                <p class="title-desc"><%=traduci("simitxt")%></p>
                            </header>
                            
                                <div class="carousel-controls">
                                    <div id="similiar-items-slider-prev" class="carousel-btn carousel-btn-prev"></div><!-- End .carousel-prev -->
                                    <div id="similiar-items-slider-next" class="carousel-btn carousel-btn-next carousel-space"></div><!-- End .carousel-next -->
                                </div><!-- End .carousel-controls -->
                                <div class="similiar-items-slider owl-carousel">
                                    <%
									response.write box_rs_articoli(rs_articoli,0)
                                    %>
                                </div><!--purchased-items-slider -->
                            </div><!-- End .purchased-items-container -->
							<%end if 'not rs_articoli.eof
							Set rs_articoli=Nothing
							end if	'idpro>0
							%>
        				
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
			</div><!-- End .container -->
        
        </section><!-- End #content -->
    <%
	if err.number<>0 then
		add2log "ERRORE in"&questofile&":"&Err.Description,0
	end if  
	    
	%>
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
	<script src="carrello.js"></script>

    </body>
</html>
<!--#include virtual="/dettaglio_inc.asp" -->
