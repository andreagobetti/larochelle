<%
sub carrellino()
	if utente_bot then exit sub
	if session("cache_carrellino")=""  then
		dim stringa, stringa_tmp
		call carica_dizionario()
		
		sql="select Sum(carrello.quantita) FROM carrello WHERE "&where_carrello
		set rs_tot=conn.execute (sql)
		
		n_articoli=rs_tot(0)
		if not isnull(n_articoli) then
			n_articoli=clng(n_articoli)
		else
			n_articoli=0
		end if
		Set rs_tot=Nothing		
		sql="select carrello.*, prodotti.codice, prodotti.articolo, prodotti.variante1, prodotti.variante2, prodotti.prezzo, prodotti.costo, prodotti.Promozione, prodotti.Sconto, prodotti.Prodata, varianti_a.variante_a, varianti_a.prezzo_ve_va, varianti_b.variante_b, prodotti.idfor, prodotti.fileimg, prodotti.aggiungi_a_ordine FROM varianti_b RIGHT JOIN (varianti_a RIGHT JOIN (carrello INNER JOIN prodotti ON carrello.idpro = prodotti.IDpro) ON varianti_a.IDvara = carrello.idvara) ON varianti_b.IDvarb = carrello.idvarb where "&where_carrello&" order by idcar desc"
		'response.write sql
		set rs_carrello= Server.CreateObject("ADODB.Recordset")
		rs_carrello.Open sql, conn, 1, 3
		conteggio_articoli=0
		if n_articoli>0 then
			txt_carrello=traduci("txt1") 
		else 
			txt_carrello=traduci("txt2") 
		end if
		n_articolo=1
		stringa="<div class=""btn-group dropdown-cart"">"&_
					"<button type=""button"" class=""btn btn-custom dropdown-toggle"" data-toggle=""dropdown"" id=""apri_carrello"">"&_
					"<span class=""cart-menu-icon""></span>"&n_articoli&"&nbsp;"&traduci("item(s)")&_
					"</button>"&_
						"<div class=""dropdown-menu dropdown-cart-menu pull-right clearfix"" role=""menu"" id=""carrello_small"">"&_
							"<p class=""dropdown-cart-description"">"&txt_carrello&"</p>"
							if n_articoli>0 then
								stringa_tmp="<ul class=""dropdown-cart-product-list"">"
								do while not rs_carrello.eof and n_articolo<4
									prezzo_ve_va=rs_carrello("prezzo_ve_va")
									if not isnull(prezzo_ve_va) then
										prezzo_ve_va=cdbl(prezzo_ve_va)
									else
										prezzo_ve_va=0
									end if
										
									if prezzo_ve_va>0 then prezzo=rs_carrello("prezzo_ve_va") else prezzo=rs_carrello("prezzo")
									prezzo=cdbl(prezzo)
									sconto=cdbl(rs_carrello("sconto"))
									if rs_carrello("promozione")=1 and sconto>0 then
										prezzoeff=prezzo*(1-sconto/100)
									else
										prezzoeff=prezzo
									end if
									if Application("prezzi_con_iva")>0 then
										prezzoeff=applica_iva(prezzo)
									end if
									
									
									tot_carrello=tot_carrello+totale_riga(prezzoeff,0,rs_carrello("quantita"))
									
									
									if rs_carrello("promozione")=1 and sconto>0 then
										txtprezzo="<s>"&simbolo_valuta&formatnumber(prezzo,2)&"</s><br><b>" & formatnumber(prezzoeff,2)&"</b>"
									else
										txtprezzo=simbolo_valuta&formatnumber(prezzoeff,2)
									end if								
									
									
									
									
									if n_articolo<4 then
										stringa_tmp=stringa_tmp&"<li class=""item clearfix"" id=""item-car-sm_"&rs_carrello("idcar")&""">"&_
											"<figure><a href=""product.asp?idpro="&rs_carrello("idpro")&"""><img src=""send_img.asp?imgs1="&rs_carrello("fileimg")&""" alt="""&rs_carrello("articolo")&"""></a></figure>"&_
											"<div class=""dropdown-cart-details""><p class=""item-name""><a href=""product.asp?idpro="&rs_carrello("idpro")&""">"&rs_carrello("articolo")&"</a></p><p>"&rs_carrellO("quantita")&"x<span class=""item-price"">"&txtprezzo&"</span></p></div><!-- End .dropdown-cart-details --></li>"
											
									end if
									conteggio_articoli=conteggio_articoli+rs_carrellO("quantita")

									rs_carrello.MoveNext
									n_articolo=n_articolo+1
								loop
									
								
								if n_articoli-conteggio_articoli>0 then
									stringa=stringa&"<li><p>e altri "&n_articoli-conteggio_articoli&" articoli</p></li>"
								end if
								stringa_tmp=stringa_tmp&"</ul>"
								
								totale_carrello=tot_carrello
								tot_carrello=split(formatnumber(tot_carrello,2),",")
								rs_carrello.close
								set rs_carrello=Nothing
								stringa_tmp=stringa_tmp&_
									"<ul class=""dropdown-cart-total"">"&_
										"<li><span class=""dropdown-cart-total-title"">"&traduci("tot")&":</span>"&simbolo_valuta&tot_carrello(0)&"<span class=""sub-price"">,"&tot_carrello(1)&"</span></li>"&_
									"</ul><!-- .dropdown-cart-total -->"&_
									"<div class=""dropdown-cart-action"">"&_
										"<p><a href=""cart.asp"" class=""btn btn-custom-2 btn-block"">"&traduci("carrello")&"</a></p>"
	        						if cdbl(totale_carrello)>=cdbl(Application("minimo_preventivo")) then
		        						stringa_tmp=stringa_tmp&"<p><a href=""checkout.asp?preventivo=true"" class=""btn btn-custom btn-block"">"&traduci("preventiv")&"</a></p>"
        							end if								
										
										
									stringa_tmp=stringa_tmp&"<p><a href=""checkout.asp"" class=""btn btn-custom btn-block"">"&traduci("checkout")&"</a></p>"&_
										"</div><!-- End .dropdown-cart-action -->"
								stringa=stringa&stringa_tmp
									
							end if
							stringa=stringa&_
							"</div><!-- End .dropdown-cart -->"&_
							"</div><!-- End .btn-group -->"
		session("cache_carrellino")=stringa
	end if
	response.write session("cache_carrellino")
end sub
%>