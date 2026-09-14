<!--#include virtual="/setup.asp" -->
<%
const tipo_category=1
css_category=true
dim idsettore,idtag
idsettore=request("idsettore")
dim pulsanti_indietro

idsettore=request("idsettore")
page = Request("page")
ordinaPer=request("ordinaper")
ipSize=request("ipSize")
idmodello=request("idmodello")
idtag=request("idtag")

if idsettore <>"" and not isnumeric(idsettore) then
	call determina_sqlinjection(idsettore)
	idsettore=""
end if
if page <>"" and not isnumeric(page) then
	call determina_sqlinjection(page)
	page=""
end if
if ordinaPer <>"" and not isnumeric(ordinaPer) then
	call determina_sqlinjection(ordinaPer)
	ordinaPer=""
end if
if ipSize <>"" and not isnumeric(ipSize) then
	call determina_sqlinjection(ipSize)
	ipSize=""
end if
if idmodello <>"" and not isnumeric(idmodello) then
	call determina_sqlinjection(idmodello)
	idmodello=""
end if
if idtag <>"" and not isnumeric(idtag) then
	call determina_sqlinjection(idtag)
	idtag=""
end if

dim str_settori, str_elencosettori
str_elencosettori=""
if idsettore="" then idsettore=0
if idsettore=0 then
	meta_title="Catalogo"
else
	set rs_settori=conn.execute ("select settori.* from settori where idsettore="&idsettore)
	if not rs_settori.eof then
		meta_title=rs_settori("nome_settore"&lingua)
		logga "sett",idsettore

	else
		meta_title="Catalogo"
		idsettore=0
	end if
	set rs_settori=nothing

end if

%>
<!--#include virtual="/config/header_inc.asp" -->
<%

if request("riordina")<>"" then
	ipsize=2
	riordina=true
end if
%>
        <section id="content">

			<%'if idsettore>0 then %>
			<%if true then %>
        	<div id="category-breadcrumb">
        		<div class="container">
					<ul class="breadcrumb">
							<%
							pulsanti_indietro=""
							response.write breadcrumb(idsettore,"",lingua)
							str_settori=str_settori&"-"
							%>

					</ul>
        		</div>
        	</div>
        	<%else %>
			<div class="md-margin"></div><!-- .space -->
			<%end if %>

        	        												<div class="row">
                                    									<div class="container" style="background-color: #dcdc07;">
                                    										<h3>Attenzione</h3>

                                                                                Si accettano resi per sostituzioni e/o riparazioni, <strong>entro un mese dalla consegna e solo degli articoli provvisti di cartellino e nella confezione integra</strong>, con la seguente precisazione:
                                    											<ul>
                                    											<li><strong>Abbigliamento:</strong> una volta personalizzati con scritte e/o loghi, non potranno essere sostituiti salvo difetti di fabbricazione</li>
                                    											<li><strong>Calzature:</strong> solo modelli a listino (esclusi quelli su ordinazione), per difetti di fabbricazionbe, per taglia errata</li>
                                                                                </ul>
<strong>Sostituzioni</strong><br>
Per motivi amministrativi e fiscali, saranno concesse ed effettuate:<br>
dopo aver verificato l'integrità dei prodotti, 
nel caso di difetti, dopo aver esaminato il prodotto.
                                    									</div><!-- End .container -->
                                    								</div><!-- End .row -->

        		<div class="row">
        			<div class="container" style="background-color: #28fc03;">
						<h3>I BOLLINI VERDI PRESENTI SUL CATALOGO INDICANO GLI ARTICOLI CONFORMI AL NUOVO CAPITOLATO 2021 POLIZIA LOCALE REGIONE PIEMONTE</h3>
						<p>
							I codici relativi ai prodotti presenti su questo catalogo sono tutti inseriti su mepa</p>

	    			</div><!-- End .container -->
        		</div><!-- End .row -->
        		<div class="row">
        			<div class="container" style="background-color: #FF5733; margin-bottom: 10px; color: #FCDD00;">
						<h3 style="color: #FCDD00;">Gli articoli privi di bollino verde, non più previsti dal nuovo capitolato P.L. Regione Piemonte, sono in esaurimento ed  alcuni disponibili solo su ordinazione.</h3>
						 Non è possibile garantire la successiva sostituzione.

	    			</div><!-- End .container -->
        		</div><!-- End .row -->

        	<div class="container">

	        <%if tipo_category=1 then%>
	        	<div id="griglia-settori">
					        	<%
						        	if application("cache_settori"&lingua)="" then
									'if true then
							        	testo=""
							        	call Elencosettori(0,0,0,testo,false)
							        	application("cache_settori"&lingua)=testo
							        	'add2log "Rigenerata cache elenco settori",1
									end if
						        	response.write application("cache_settori")
						        	if utente_admin then
							        	response.write Get_Elencosettori_Admin()
						        	end if

					        	%>
					        	<%
						        	if idsettore>0 then response.write pulsanti_indietro



						        	%>


	        	</div>
	        <%end if%>

        		<div class="row">
        			<div class="col-md-12">

        				<div class="row">

        					<div class="col-md-9 col-sm-8 col-xs-12 main-content" id="elenco_articoli">
        						<%

        						call elenco_articoli()
        						%>

        					</div><!-- End .col-md-9 -->

        					<aside class="col-md-3 col-sm-4 col-xs-12 sidebar">

        						<div class="widget">
        							<div class="panel-group custom-accordion sm-accordion" id="category-filter">

        								<%if false then%>
        								<div class="panel">
											<div class="accordion-header">
												<div class="accordion-title"><span>Brand</span></div><!-- End .accordion-title -->
												<a class="accordion-btn opened"  data-toggle="collapse" data-target="#category-list-2"></a>
											</div><!-- End .accordion-header -->

										<div id="category-list-2" class="collapse in">
											<div class="panel-body">
											<ul class="category-filter-list jscrollpane">
												<li><a href="#">Samsung (50)</a></li>
												<li><a href="#">Apple (80)</a></li>
												<li><a href="#">HTC (20)</a></li>
												<li><a href="#">Motoroloa (20)</a></li>
												<li><a href="#">Nokia (11)</a></li>
											</ul>
											</div><!-- End .panel-body -->
										</div><!-- #collapse -->

										</div><!-- End .panel -->

        							<div class="panel">
											<div class="accordion-header">
												<div class="accordion-title"><span>Price</span></div><!-- End .accordion-title -->
												<a class="accordion-btn opened"  data-toggle="collapse" data-target="#category-list-3"></a>
											</div><!-- End .accordion-header -->

										<div id="category-list-3" class="collapse in">
											<div class="panel-body">
												<div id="price-range">

												</div><!-- End #price-range -->
												<div id="price-range-details">
													<span class="sm-separator">from</span>
													<input type="text" id="price-range-low" class="separator">
													<span class="sm-separator">to</span>
													<input type="text" id="price-range-high">
												</div>
												<div id="price-range-btns">
													<a href="#" class="btn btn-custom-2 btn-sm">Ok</a>
													<a href="#" class="btn btn-custom-2 btn-sm">Clear</a>
												</div>
											</div><!-- End .panel-body -->
										</div><!-- #collapse -->
										</div><!-- End .panel -->

                                        <%end if%>
        							</div><!-- .panel-group -->
        						</div><!-- End .widget -->
        						<%if false then%>
        						<div class="widget featured">
        							<h3>Featured</h3>

        							<div class="featured-slider flexslider sidebarslider">
        								<ul class="featured-list clearfix">
        									<li>
        										<div class="featured-product clearfix">
        											<figure>
        												<img src="images/products/thumbnails/placeholder.jpg" alt="item5">
        											</figure>
        											<h5><a href="#">Jacket Suiting Blazer</a></h5>
        											<div class="ratings-container">
														<div class="ratings">
															<div class="ratings-result" data-result="84"></div>
														</div><!-- End .ratings -->
													</div><!-- End .rating-container -->
        											<div class="featured-price">$40</div><!-- End .featured-price -->
        										</div><!-- End .featured-product -->

        										<div class="featured-product clearfix">
        											<figure>
        												<img src="images/products/thumbnails/placeholder.jpg" alt="item1">
        											</figure>
        											<h5><a href="#">Gap Graphic Cuffed</a></h5>
        											<div class="ratings-container">
														<div class="ratings">
															<div class="ratings-result" data-result="84"></div>
														</div><!-- End .ratings -->
													</div><!-- End .rating-container -->
        											<div class="featured-price">$18</div><!-- End .featured-price -->
        										</div><!-- End .featured-product -->

        										<div class="featured-product clearfix">
        											<figure>
        												<img src="images/products/thumbnails/placeholder.jpg" alt="item2">
        											</figure>
        											<h5><a href="#">Women's Lauren Dress</a></h5>
        											<div class="ratings-container">
														<div class="ratings">
															<div class="ratings-result" data-result="84"></div>
														</div><!-- End .ratings -->
													</div><!-- End .rating-container -->
        											<div class="featured-price">$30</div><!-- End .featured-price -->
        										</div><!-- End .featured-product -->
        									</li>
        									<li>
        										<div class="featured-product clearfix">
        											<figure>
                                                        <img src="images/products/thumbnails/placeholder.jpg" alt="item3">
                                                    </figure>
        											<h5><a href="#">Swiss Mobile Phone</a></h5>
        											<div class="ratings-container">
														<div class="ratings">
															<div class="ratings-result" data-result="64"></div>
														</div><!-- End .ratings -->
													</div><!-- End .rating-container -->
        											<div class="featured-price">$39</div><!-- End .featured-price -->
        										</div><!-- End .featured-product -->

        										<div class="featured-product clearfix">
        											<figure>
        												<img src="images/products/thumbnails/placeholder.jpg" alt="item4">
        											</figure>
        											<h5><a href="#">Zwinzed HeadPhones</a></h5>
        											<div class="ratings-container">
														<div class="ratings">
															<div class="ratings-result" data-result="94"></div>
														</div><!-- End .ratings -->
													</div><!-- End .rating-container -->
        											<div class="featured-price">$18.99</div><!-- End .featured-price -->
        										</div><!-- End .featured-product -->

        										<div class="featured-product clearfix">
        											<figure>
        												<img src="images/products/thumbnails/placeholder.jpg" alt="item7">
        											</figure>
        											<h5><a href="#">Kless Man Suit</a></h5>
        											<div class="ratings-container">
														<div class="ratings">
															<div class="ratings-result" data-result="74"></div>
														</div><!-- End .ratings -->
													</div><!-- End .rating-container -->
        											<div class="featured-price">$99</div><!-- End .featured-price -->
        										</div><!-- End .featured-product -->
        									</li>
        									<li>

        										<div class="featured-product clearfix">
        											<figure>
                                                        <img src="images/products/thumbnails/placeholder.jpg" alt="item4">
                                                    </figure>
        											<h5><a href="#">Gap Graphic Cuffed</a></h5>
        											<div class="ratings-container">
														<div class="ratings">
															<div class="ratings-result" data-result="84"></div>
														</div><!-- End .ratings -->
													</div><!-- End .rating-container -->
        											<div class="featured-price">$17</div><!-- End .featured-price -->
        										</div><!-- End .featured-product -->

        										<div class="featured-product clearfix">
        											<figure>
        												<img src="images/products/thumbnails/placeholder.jpg" alt="item6">
        											</figure>
        											<h5><a href="#">Women's Lauren Dress</a></h5>
        											<div class="ratings-container">
														<div class="ratings">
															<div class="ratings-result" data-result="84"></div>
														</div><!-- End .ratings -->
													</div><!-- End .rating-container -->
        											<div class="featured-price">$30</div><!-- End .featured-price -->
        										</div><!-- End .featured-product -->
        									</li>
        								</ul>
        							</div><!-- End .featured-slider -->
        						</div><!-- End .widget -->
    							<%end if
	    							if utente_admin then%>
									<a href="pag_adm_elenco_banner.asp" class="btn btn-custom">Gestisci banner</a>
									<%end if
	    							response.write Get_banner_slider()
	set conn = nothing


    							%>

        					</aside><!-- End .col-md-3 -->
        				</div><!-- End .row -->
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
			</div><!-- End .container -->

        </section><!-- End #content -->

	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
    <script src="jquery/js/jquery.scrollintoview.min.js"></script>
    <script src="jquery/js/modernizr.custom.min.js"></script>
    <script src="jquery/js/shuffle.js"></script>
	<script src="js/jquery.nouislider.min.js"></script>
	<script src="js/jquery.jscrollpane.min.js"></script>
	<script src="js/jquery.mousewheel.js"></script>
	<script src="//ajax.googleapis.com/ajax/libs/jqueryui/1.11.1/jquery-ui.min.js"></script>
	<script src="js/main.js"></script>
	<script>

	$(function(){

		 <%if tipo_category=1 then%>
		/* initialize shuffle plugin */
		var $grid = $('#griglia-settori');
		var $item;
		$grid.shuffle({
			itemSelector: '.col-md-3', // the selector for the items in the grid
			group: '<%=idsettore%>',
			initialSort: 'data-sort'

		});
		$grid.on('done.shuffle', function() {
		  console.log('Finished initializing shuffle!');
		});





		var ulbreadcrumb=$("#category-breadcrumb").find("ul");
		<%end if %>


			//onerror=handleErr;

			function handleErr(msg,url,l){
			console.log("Gestione errore");
			var txt="";
			txt="Errore: " + msg + "\n";
			txt+="URL: " + url + "\n";
			txt+="Line: " + l + "\n\n";
			$.ajaxSetup({ cache: false });
			$.ajax({
							url     : "searcher.asp",
							cache: false,
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt),
							success: function(data){
								console.log("Errore inviato");
								}
							,error:function(xhr, textStatus, error){
								console.log("xhr.statusText:"+xhr.statusText);
								console.log("xhr.responseText:"+xhr.responseText);
								console.log("textStatus:"+textStatus);
								console.log("error:"+error);
								console.log("INVIO ERRORE FALLITO");
							}
			});

			return true;
		}

		<%if utente_admin then %>
		$.ajaxSetup({ cache: false });
		<%end if%>
		//$.get("category_ajax.asp?add=si", function (result) {
		//	$("#elenco_articoli").html(result);
		//	last_id=<%=idsettore%>;

		//});

		var str_settori="<%=str_settori%>";
		var last_id=<%=idsettore%>;

			$("body").on("click",".onclick",function(e){
			e.preventDefault();

			var id = $(this).attr('id').replace("#ids_", "");
			console.log ("idsettore:"+id);

			//Breadcrumb e Shuffle
			//var groupName = $(this).attr('data-group');
			groupName=id;
			<%if tipo_category=1 then%>
			if($(this).hasClass("bread"))
			{
				$(this).parent("li").nextAll().remove()

			}
			else if($(this).hasClass("indietro"))
			{
				$(ulbreadcrumb).find("li").last().remove();
				$(ulbreadcrumb).find("li").last().addClass( "active" );
			}
			else
			{
				$(ulbreadcrumb).find("li").last().removeClass( "active" );
				$(ulbreadcrumb).append('<li><a href="category.asp?idsettore='+groupName+'" class="onclick bread" id="#ids_'+groupName+'">'+$(this).text()+'</a></li>');
			}

			// reshuffle grid
			//var item='<div class="col-md-3 col-sm-6 col-xs-12 data-groups=''["'+groupName+'"]''><a href="#" id="#ids_"  class=btn btn-custom onclick admin><i class=fa fa-star></i>Prova</a></div>';

			if (groupName!=0){
				console.log("indietro"+$("#indietro"+groupName).length);
				if ($("#indietro"+groupName).length==0){
					console.log("aggiungo indietro"+groupName)
					var $item = $('<div class="col-md-3 col-sm-6 col-xs-12" id="indietro'+groupName+'" data-groups=\'["'+groupName+'"]\' data-sort="0"><a href="category.asp?idsettore='+last_id+'"  id="#ids_'+last_id+'" class="btn btn-custom indietro onclick">&laquo; Indietro</a></div>')
					$grid.append($item);
					$grid.shuffle('appended', $item);
				}
			}
			$grid.shuffle('shuffle', groupName );
			<%end if %>


			if (id==-10){return;}
			//var id =event.target;
			if(id!=last_id){
			$.ajax({
					url     : "category_ajax.asp?add=si&idsettore=" + id,
					type    : "GET",
					success: function(data){

						$("#elenco_articoli").html(data);
						last_id=id;
						try{
							history.pushState({}, document.title , "category.asp?idsettore="+id);
						}
						catch(err){
							console.log()
						}
						//this.itemHoverAnimation();

					}
					,error:function(xhr, textStatus, error){
					      console.log("xhr.statusText:"+xhr.statusText);
					      console.log("xhr.responseText:"+xhr.responseText);
					      console.log("textStatus:"+textStatus);
					      console.log("error:"+error);
						  var txt="";
						  txt+="Errore in category.asp consultazione category_ajax<br>";
						  txt+="<br>querystring: ?idsettore="+ last_id;
						  txt+="<br>"+xhr.statusText;
						  txt+="<br>"+xhr.responseText;
						  txt+="<br>textStatus:"+textStatus;
						  txt+="<br>error:"+error;

						  $.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt),
							success: function(data){
								console.log("Errore inviato");
								}
							,error:function(xhr, textStatus, error){
								console.log("xhr.statusText:"+xhr.statusText);
								console.log("xhr.responseText:"+xhr.responseText);
								console.log("textStatus:"+textStatus);
								console.log("error:"+error);
								console.log("INVIO ERRORE FALLITO");
							}
							});
						  }
				});
			}
		});

		$("body").on("click",".ch_page",function(e){
			e.preventDefault();
			var querystring=$(this).attr('href').split("?");
			if ($(this).hasClass("myhover")){
				console.log ("selezionato, querystring prima"+querystring[1]);
				querystring[1]=removeURLParameter(querystring[1],"idtag");
				console.log ("selezionato, querystring dopo"+querystring[1]);
			}

			console.log("chpage:"+querystring[1]);
			//Aggiorno elenco
			$.ajax({
				url     : "category_ajax.asp?"+querystring[1],
				type    : "POST",
				success: function(data){

					$("#elenco_articoli").html(data);
					$("#category-breadcrumb").scrollintoview({duration: 'slow'});
					history.pushState({}, document.title , "category.asp?"+querystring[1]);

				}
				,error:function(xhr, textStatus, error){
				      console.log("xhr.statusText:"+xhr.statusText);
				      console.log("xhr.responseText:"+xhr.responseText);
				      console.log("textStatus:"+textStatus);
				      console.log("error:"+error);
					  var txt="";
					  txt+="Errore in category.asp consultazione category_ajax da .ch_page<br>";
					  txt+="<br>querystring: ?"+querystring[1];
					  txt+="<br>"+xhr.statusText;
					  txt+="<br>"+xhr.responseText;
					  txt+="<br>textStatus:"+textStatus;
					  txt+="<br>error:"+error;

					  $.ajax({
						url     : "searcher.asp",
						type    : "post",
						data	: "txt_errore="+encodeURIComponent(txt),
						success: function(data){
							console.log("Errore inviato");
							}
						,error:function(xhr, textStatus, error){
							console.log("xhr.statusText:"+xhr.statusText);
							console.log("xhr.responseText:"+xhr.responseText);
							console.log("textStatus:"+textStatus);
							console.log("error:"+error);
							console.log("INVIO ERRORE FALLITO");
						}
						});
					  }
			});



		});
		$("body").on("click","#settore-in-preferiti",function(e){
			e.preventDefault();
			var idsettore=$(this).attr("data-idset");
			console.log("idsettore"+idsettore);
			$.ajax({
				url     : "searcher.asp?",
				type    : "POST",
				dataType: 'json',
				data: {oper:'settore_in_preferiti',
				idsettore: idsettore},
				success: function(data){

					alert("Articoli aggiunti a preferiti");
				}
				,error:function(xhr, textStatus, error){
				      console.log("xhr.statusText:"+xhr.statusText);
				      console.log("xhr.responseText:"+xhr.responseText);
				      console.log("textStatus:"+textStatus);
				      console.log("error:"+error);
					  var txt="";
					  txt+="Errore in category.asp settore-in-preferiti<br>";
					  txt+="<br>querystring: ?"+querystring[1];
					  txt+="<br>"+xhr.statusText;
					  txt+="<br>"+xhr.responseText;
					  txt+="<br>textStatus:"+textStatus;
					  txt+="<br>error:"+error;

					  $.ajax({
						url     : "searcher.asp",
						type    : "post",
						data	: "txt_errore="+encodeURIComponent(txt),
						success: function(data){
							console.log("Errore inviato");
							}
						,error:function(xhr, textStatus, error){
							console.log("xhr.statusText:"+xhr.statusText);
							console.log("xhr.responseText:"+xhr.responseText);
							console.log("textStatus:"+textStatus);
							console.log("error:"+error);
							console.log("INVIO ERRORE FALLITO");
						}
						});
					  }
			});



		});


		<%if utente_admin and riordina then%>
			$( "#category-item-container" ).sortable({
				cursor: 'move',
				tolerance: 'pointer',
				handle: ".item",
				items: ".col-md-4",
				//placeholder: "red",
				update : function () {
	      			//var order = $('#category-item-container').sortable('serialize');
	      			//console.log("ordine="+order);
	  			}

			}).disableSelection();

			$("#riordina_articoli").click(function(e){
				var order = $('#category-item-container').sortable('serialize');
	      		console.log("ordine="+order);

	      		$.ajaxSetup({ cache: false });

			    $.ajax({
					url     : "ajax_function.asp?oper=riordina_articoli&idsettore="+last_id,
					type    : "post",
					dataType: 'json',
					data	: order,
					success: function(data){
						//alert(data.Message);

						$(".alert-success").hide();
						$(".alert-success").slideDown();

					}
					,error:function(xhr, textStatus, error){
					      console.log("xhr.statusText:"+xhr.statusText);
					      console.log("xhr.responseText:"+xhr.responseText);
					      console.log("textStatus:"+textStatus);
					      console.log("error:"+error);
						  }
				});
			});
		<%end if%>

	});
	function removeURLParameter(url, parameter) {
	    //prefer to use l.search if you have a location/link object

	        var prefix= encodeURIComponent(parameter)+'=';
	        var pars= url.split(/[&;]/g);

	        //reverse iteration as may be destructive
	        for (var i= pars.length; i-- > 0;) {
	            //idiom for string.startsWith
	            if (pars[i].lastIndexOf(prefix, 0) !== -1) {
	                pars.splice(i, 1);
	            }
	        }

	        url= pars.join('&');
	        return url;
	}
	</script>

    </body>
</html>
<!--#include virtual="/dettaglio_inc.asp" -->
<!--#include virtual="/category_inc.asp" -->
<%

function breadcrumb(id,tmp,lingua)
	'create recordset

	set rs_settore=conn.execute ("select * FROM settori WHERE idsettore = " & id&";")
		if id=0 then
				tmp="<li><a href='"&questofile&"' id=""#ids_0"" class=""onclick bread"">Tutte le categorie</a></li>"&tmp
		end if

	'find child with thread_parent=his parent id
	if not rs_settore.eof then
		pulsanti_indietro="<div class=""col-md-3 col-sm-6 col-xs-12 shuffle-item"" id=""indietro"&id&""" data-groups='["""&id&"""]' data-sort=""0""><a href=""category.asp?idsettore="&rs_settore("idpadre")&""" id=""#ids_"&rs_settore("idpadre")&""" class=""btn btn-custom indietro onclick"">&laquo; Indietro</a></div>"&pulsanti_indietro


		if tmp="" then txtclass="class=""active"""
		if tmp="" then
			tmp="<li><a href='category.asp?idsettore="&rs_settore("idsettore")&"' class=""active"">"&rs_settore("Nome_Settore"&lingua)&"</a></li>"&tmp
		else
			tmp="<li><a href='category.asp?idsettore="&rs_settore("idsettore")&"'>"&rs_settore("Nome_Settore"&lingua)&"</a></li>"&tmp
		end if
		tmp=breadcrumb(rs_settore("idpadre"),tmp,lingua)

	end if

	breadcrumb=tmp
	set rs_settore=nothing
End function


sub Elencosettori(layer,id,idindietro,testo,admin)
	'declaring
	dim sql_order,str,MessageSpacing,mesid,spaceing,tid
	'create recordset
	set rs_order=server.createObject("adodb.recordset")
	'find child with thread_parent=his parent id
	'sql_order="select settori.idsettore, settori.Nome_Settore, settori.idpadre, settori.nascondi_prezzi, settori.nascondi, settori.ordine, Count(prodotti.IDpro) AS ConteggioDiIDpro, Count(settori_1.idsettore) AS ConteggioDiidsettore FROM settori AS settori_1 RIGHT JOIN ((settori LEFT JOIN settori_prodotti ON settori.idsettore = settori_prodotti.IDsettore) LEFT JOIN prodotti ON settori_prodotti.IDpro = prodotti.IDpro) ON settori_1.idsettore = settori.idpadre GROUP BY settori.idsettore, settori.Nome_Settore, settori.idpadre,settori.nascondi, settori.ordine, settori.nascondi_prezzi HAVING settori.idpadre=" & id & ""


	sql_order="select settori.idsettore, settori.nome_settore,  settori.nascondi_prezzi, settori_settori.idpadre, settori.colore from settori_settori inner join settori on settori_settori.idfiglio = settori.idsettore where settori_settori.idpadre="&id

	sql_order=sql_order&" and settori.nascondi=0 "
	sql_order=sql_order&" order by settori_settori.ordine"
'response.write sql_order


	rs_order.open sql_order,conn
	nodo=true
	'testo=testo&testo_dopo
	'	testo_dopo=""
	if testo="" then 'Inizio del blocco
		'testo=testo&  "<div class=""col-md-4 item"" data-groups='[""indietro""]'><div class=""divpadding""><a href=""#"" id=""indietro"" data-group="""" ></a></div></div><!-- End .col-md-4 -->"
	end if
	if layer>0 and false then 'Inizio di ogni sottogruppo "INDIETRO"
		testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12"" data-groups='["""&id&"""]'><a href=""category.asp?idsettore="&idindietro&"""  id=""#ids_"&idindietro&""" class=""btn btn-custom indietro onclick"">&laquo; Indietro</a></div><!-- End .col-md-4 -->"
	end if

	classadmin=""
	do until rs_order.eof
		stringa="|"&rs_order("idpadre")&"-"&rs_order("idsettore")&"|"
		if instr(str_elencosettori,stringa)=0 then
			testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12"" data-groups='["""&rs_order("idpadre")&"""]'><a href=""category.asp?idsettore="&rs_order("idsettore")&""" id=""#ids_"&rs_order("idsettore")&"""  class=""btn btn-custom onclick strong "&classadmin&""""
			if rs_order("colore")<>"" then

				testo=testo&" style=""background-color: #"&rs_order("colore")&""""
			end if
			testo=testo&">"&rs_order("nome_settore")
			testo=testo&"</a></div><!-- End .col-md-4 -->"
			str_elencosettori=str_elencosettori&stringa
		end if
	    call Elencosettori(layer+1,rs_order("idsettore"),id,testo,admin)
		rs_order.movenext
	loop

	rs_order.Close
	set rs_order=Nothing
	set rs_tot=nothing
End Sub
function Get_Elencosettori_Admin()
	dim testo
	testo=""
	if application("cache_settori_admin")="" then
		''''''''''ADMIN'''''''

		'elenco settori nascosti
		sql_order="select settori.idsettore, settori.Nome_Settore, settori.idpadre, settori.nascondi_prezzi, settori.nascondi, settori.ordine FROM settori where nascondi=1"
		sql_order=sql_order&" order by settori.ordine"
		set rs_tot=conn.execute (sql_order)
		if not rs_tot.eof then
			testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12"" data-groups='[""0""]'><a href=""#"" id=""#ids_-10""  class=""btn btn-custom onclick strong admin""><i class=""fa fa-star""></i> "
			testo=testo&"settori nascosti"
			testo=testo&"</a></div><!-- End .col-md-4 -->"
		end if
		do until rs_tot.eof
			testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12"" data-groups='[""-10""]'><a href=""category.asp?idsettore="&rs_tot("idsettore")&""" id=""#ids_"&rs_tot("idsettore")&"""  class=""btn btn-custom onclick strong admin"">"
			testo=testo&rs_tot("nome_settore")
			testo=testo&"</a></div><!-- End .col-md-4 -->"
			rs_tot.movenext
		loop

		'Aggiungo categoria fuori catalogo
		sql="select Count(*) AS Conteggio FROM prodotti where visibilita>0 and codice<>''"
		set rs_tot=conn.execute (sql)
		if not rs_tot.eof then
			n=clng(rs_tot("Conteggio"))
		else
			n=0
		end if
		if n>0 then

			idsettoretmp=-2
			testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12 "" data-groups='[""0""]'><a href=""category.asp?idsettore="&idsettoretmp&""" id=""#ids_"&idsettoretmp&"""  class=""btn btn-custom onclick  admin"" ><i class=""fa fa-star""></i> "
			testo=testo&"Fuori catalogo <strong>"&n&"</strong>"
			testo=testo&"</a></div><!-- End .col-md-4 -->"
		end if
		'Aggiungo categoria non in vendita
		sql="select Count(prodotti.vendita) AS Conteggio FROM prodotti where vendita=0 and visibilita=0 and codice<>''"
		set rs_tot=conn.execute (sql)
		if not rs_tot.eof then
			n=clng(rs_tot("Conteggio"))
		else
			n=0
		end if
		if n>0 then
			n=rs_tot("Conteggio")
			idsettoretmp=-3
			testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12"" data-groups='[""0""]'><a href=""category.asp?idsettore="&idsettoretmp&""" id=""#ids_"&idsettoretmp&"""  class=""btn btn-custom onclick admin""><i class=""fa fa-star""></i> "
			testo=testo&"Non in vendita <strong>"&n&"</strong>"
			testo=testo&"</a></div><!-- End .col-md-4 -->"
		end if

		'Aggiungo categoria senza settore
		sql="select Count(prodotti.IDpro) AS Conteggio FROM prodotti LEFT JOIN settori_prodotti ON prodotti.IDpro = settori_prodotti.IDpro WHERE (settori_prodotti.ID Is Null) AND visibilita=0 and codice<>'';"
		set rs_tot=conn.execute (sql)
		if not rs_tot.eof then
			n=clng(rs_tot("Conteggio"))
		else
			n=0
		end if
		if n>0 then
			n=rs_tot("Conteggio")
			idsettoretmp=-4
			testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12"" data-groups='[""0""]'><a href=""category.asp?idsettore="&idsettoretmp&""" id=""#ids_"&idsettoretmp&"""  class=""btn btn-custom onclick admin""><i class=""fa fa-star""></i> "
			testo=testo&"Senza categoria <strong>"&n&"</strong>"
			testo=testo&"</a></div><!-- End .col-md-4 -->"
		end if

		'Aggiungo categoria nuovi prodotti
		sql="select Count(prodotti.IDpro) AS Conteggio FROM prodotti WHERE visibilita=0 and novita=1;"
		set rs_tot=conn.execute (sql)
		if not rs_tot.eof then
			n=clng(rs_tot("Conteggio"))
		else
			n=0
		end if
		if n>0 then
			n=rs_tot("Conteggio")
			idsettoretmp=-5
			testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12"" data-groups='[""0""]'><a href=""category.asp?idsettore="&idsettoretmp&""" id=""#ids_"&idsettoretmp&"""  class=""btn btn-custom onclick admin""><i class=""fa fa-star""></i> "
			testo=testo&"Nuovi <strong>"&n&"</strong>"
			testo=testo&"</a></div><!-- End .col-md-4 -->"
		end if
		'Aggiungo categoria promozione
		sql="select Count(prodotti.IDpro) AS Conteggio FROM prodotti WHERE visibilita=0 and promozione=1;"
		set rs_tot=conn.execute (sql)
		if not rs_tot.eof then
			n=clng(rs_tot("Conteggio"))
		else
			n=0
		end if
		if n>0 then
			n=rs_tot("Conteggio")
			idsettoretmp=-6
			testo=testo&  "<div class=""col-md-3 col-sm-6 col-xs-12"" data-groups='[""0""]'><a href=""category.asp?idsettore="&idsettoretmp&""" id=""#ids_"&idsettoretmp&"""  class=""btn btn-custom onclick admin""><i class=""fa fa-star""></i> "
			testo=testo&"Promozione <strong>"&n&"</strong>"
			testo=testo&"</a></div><!-- End .col-md-4 -->"
		end if
		set rs_tot=nothing
		Application.Lock
		application("cache_settori_admin")=testo
		Application.Unlock
	end if
	Get_Elencosettori_Admin=application("cache_settori_admin")
End Function
function Get_banner_slider()
	if application("cache_banner_slider")="" then
		dim stringa
		stringa=""
		set rs_brand=conn.execute ("select news.* from news where tipo='cb' order by ordine")
		if not rs_brand.eof then
			stringa=stringa&"<div class=""widget banner-slider-container"">"&_
								"<div class=""banner-slider flexslider"">"&_
									"<ul class=""banner-slider-list clearfix"">"
				do while not rs_brand.eof
					stringa=stringa&"<li><a href="""&rs_brand("testo")&"""><img src="""&rs_brand("testata")&""" alt="""&rs_brand("titolo")&"""></a></li>"
					rs_brand.movenext
				Loop
				stringa=stringa&"</ul>"&_
								"</div>"&_
							"</div><!-- End .widget -->"
		else
			stringa="<!-- Nessun banner slider -->"

	    end if
	    set rs_brand=Nothing
		Application.Lock
	    application("cache_banner_slider")=stringa
		Application.Unlock
	end if 'application("cache_banner_slider")=""
	Get_banner_slider=application("cache_banner_slider")
end function
%>
