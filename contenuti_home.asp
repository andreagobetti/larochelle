<%
function brand_slider_container()
	dim link_start,link_end,stringa
	stringa="<!-- brand-slider-container -->"&vbcrlf
	set rs_brand=conn.execute ("select brand.* from brand where in_home=1 order by ordine")
	if not rs_brand.eof then
		stringa=stringa&"<div id=""brand-slider-container"" class=""carousel-wrapper"">"&_
					"<header class=""content-title"">"&_
						"<div class=""title-bg"">"&_
							"<h2 class=""title"">"&traduci("brands_slider")&"</h2>"&_
						"</div><!-- End .title-bg -->"&_
					"</header>"&_
					"<div class=""carousel-controls"">"&_
						"<div id=""brand-slider-prev"" class=""carousel-btn carousel-btn-prev"">"&_
						"</div><!-- End .carousel-prev -->"&_
						"<div id=""brand-slider-next"" class=""carousel-btn carousel-btn-next carousel-space"">"&_
						"</div><!-- End .carousel-next -->"&_
					"</div><!-- End .carousel-controllers -->"&_
					"<div class=""row"">"&_
						"<div class=""brand-slider owl-carousel"">"
		do while not rs_brand.eof
			if rs_brand("link")<>"" then
				link_start="<a href="""&rs_brand("link") &""">"
				link_end="</a>"
			else
				link_start=""
				link_end=""
			end if
			stringa=stringa&link_start&"<img src="""&rs_brand("ImmagineBrand")&""" alt="""&rs_brand("nomeBrand")&""">"&link_end
	        rs_brand.MoveNext
        Loop
        stringa=stringa&"</div><!-- End .brand-slider -->"&_
					"</div><!-- End .row -->"&_
				"</div><!-- End #brand-slider-container -->"

	end if
	Set rs_brand=Nothing
	brand_slider_container=stringa
end function
	
function latestnews_slider()
	Dim rs_news,stringa
	stringa="<!-- latestnews-slider-container -->"&vbcrlf
	set rs_news=conn.execute ("select * from news where tipo='ln' order by ordine")
	if not rs_news.eof then
		stringa=stringa&"<div id=""latestnews-slider-container"" class=""carousel-wrapper"">"&_
					"<header class=""content-title"">"&_
                        "<div class=""title-bg"">"&_
	                        "<h2 class=""title"">"&traduci("lnews")&"</h2>"&_
	                    "</div><!-- End .title-bg -->"&_
	                "</header>"&_
	                "<div class=""carousel-controls"">"&_
	                	"<div id=""latestnews-slider-prev"" class=""carousel-btn carousel-btn-prev"">"&_
	                	"</div><!-- End .carousel-prev -->"&_
	                	"<div id=""latestnews-slider-next"" class=""carousel-btn carousel-btn-next carousel-space"">"&_
	                	"</div><!-- End .carousel-next -->"&_
	                "</div><!-- End .carousel-controllers -->"&_
	                "<div class=""sm-margin""></div><!-- space -->"&_
	                "<div class=""row"">"&_
	                	"<ul class=""latestnews-slider owl-carousel"">"
    	do while not rs_news.eof
	        stringa=stringa&"<li>"&_
						        "<h3><a href=""single-page.asp?id="&rs_news("idnews")&"&chk="&chkDataUser(rs_news("data"))&""">"&rs_news("titolo")&"</a></h3>"&_
						        "<p>"&rs_news("testo")&"</p>"&_
						        "<div class=""latestnews-meta-container"">"&_
						        	"<div class=""pull-left"">"&_
						        		"<a href=""#"">Read More...</a>"&_
						        	"</div><!-- End .pull-left -->"&_
									"<div class=""pull-right"">"&_
									formatdatetime(rs_news("data"),2)&_
									"</div><!-- End .pull-right -->"&_
								"</div><!-- End .latest-posts-meta-container -->"&_
							"</li>"
	        rs_news.moveNext
		loop
		stringa=stringa&"</ul>"&_
					"</div><!-- End .row -->"&_
				"</div><!-- End .latestnews-slider-container -->"
	end if
	set rs_news=Nothing
	latestnews_slider=stringa
end function
function materiali_tabs ()
	dim active,stringa
	stringa="<!-- products-tabs-container -->"&vbcrlf
	set rs_gruppi=conn.execute ("select tags.* from tags where tipo=1 order by nometag")
	if not rs_gruppi.eof then
		active="class=""active"""
		
		stringa=stringa&"<div class=""md-margin""></div><div class=""row"">"&_
					"<div class=""col-md-12 main-content"">"&_
						"<header class=""content-title"">"&_
							"<h2 class=""title"">"&traduci("materiali")&"</h2><p class=""title-desc"">"&traduci("materialitxt")&"</p>"&_
						"</header>"&_
						"<ul id=""products-tabs-list"" class=""tab-style-1 clearfix"">"
		do while not rs_gruppi.eof
			stringa=stringa&"<li "&active&"><a href=""#materiale"&rs_gruppi("idtag")&""" data-toggle=""tab"">"&rs_gruppi("nometag")&"</a></li>"	
			active=""
			rs_gruppi.MoveNext
		loop
		rs_gruppi.moveFirst
		active="active"
		stringa=stringa&"</ul>"&_
						"<div id=""products-tabs-content"" class=""row tab-content"">"
		do while not rs_gruppi.eof
			stringa=stringa&"<div class=""tab-pane "&active&""" id=""materiale"&rs_gruppi("idtag")&""">"
			testo=rs_gruppi("testo")
			if testo<>"" then 
				stringa=stringa&"<div  class=""col-md-12"">"&testo&"</div>"
			end if
			
			set rs_articoli=conn.execute("select prodotti.* FROM prodotti INNER JOIN prodotti_tags ON prodotti.IDpro = prodotti_tags.idpro where idtag="&rs_gruppi("idtag"))
			n_articoli=0
			stringa=stringa&box_rs_articoli(rs_articoli,4)
			
			set rs_articoli=Nothing
			stringa=stringa&"</div><!-- End .tab-pane -->"
			active=""
			rs_gruppi.MoveNext
		Loop
		stringa=stringa&"</div><!-- End #products-tabs-content -->"&_
					"</div><!-- End .col-md-12 .main-content -->"&_
				"</div><!-- End .row -->"
	end if 'not rs_gruppi.eof
	set rs_gruppi=Nothing
	materiali_tabs=stringa
end function

function products_tabs ()
	dim active,stringa
	stringa="<!-- products-tabs-container -->"&vbcrlf
	set rs_gruppi=conn.execute ("select gruppihome.* from gruppihome where visibile=1 order by ordine")
	if not rs_gruppi.eof then
		active="class=""active"""
		stringa=stringa&"<div class=""row"">"&_
					"<div class=""col-md-12 main-content"">"&_
						"<header class=""content-title"">"&_
							"<h2 class=""title"">"&traduci("ourprod")&"</h2><p class=""title-desc"">"&traduci("ourprodtxt")&"</p>"&_
						"</header>"&_
						"<ul id=""products-tabs-list"" class=""tab-style-1 clearfix"">"
		do while not rs_gruppi.eof
			stringa=stringa&"<li "&active&"><a href=""#gruppo"&rs_gruppi("idgruppo")&""" data-toggle=""tab"">"&rs_gruppi("nomegruppo")&"</a></li>"	
			active=""
			rs_gruppi.MoveNext
		loop
		rs_gruppi.moveFirst
		active="active"
		stringa=stringa&"</ul>"&_
						"<div id=""products-tabs-content"" class=""row tab-content"">"
		do while not rs_gruppi.eof
			stringa=stringa&"<div class=""tab-pane "&active&""" id=""gruppo"&rs_gruppi("idgruppo")&""">"
			set rs_articoli=conn.execute("select prodotti.* FROM prodotti INNER JOIN prodotti_gruppihome ON prodotti.IDpro = prodotti_gruppihome.idpro where idgruppo="&rs_gruppi("idgruppo"))
			n_articoli=0
			stringa=stringa&box_rs_articoli(rs_articoli,4)
			
			set rs_articoli=Nothing
			stringa=stringa&"</div><!-- End .tab-pane -->"
			active=""
			rs_gruppi.MoveNext
		Loop
		stringa=stringa&"</div><!-- End #products-tabs-content -->"&_
					"</div><!-- End .col-md-12 .main-content -->"&_
				"</div><!-- End .row -->"
	end if 'not rs_gruppi.eof
	set rs_gruppi=Nothing
	products_tabs=stringa
end function
	
function hot_items()
	dim stringa
	stringa="<!-- hot_items-container -->"&vbcrlf
	set rs_articoli=conn.execute("select prodotti.* from prodotti where prodotti.promozione=1 and visibilita=0 order by datamod desc LIMIT 0,12")
	if not rs_articoli.eof then
			stringa=stringa&"<!-- hot_items-items-slider -->"&_
				"<div class=""hot-items carousel-wrapper"">"&_
					"<header class=""content-title"">"&_
						"<div class=""title-bg"">"&_
							"<h2 class=""title"">"&traduci("OnSale")&"</h2>"&_
						"</div><!-- End .title-bg -->"&_
						"<p class=""title-desc"">"&traduci("OnSaletxt")&"</p>"&_
					"</header>"&_
					"<div class=""carousel-controls"">"&_
						"<div id=""hot-items-slider-prev"" class=""carousel-btn carousel-btn-prev"">"&_
						"</div><!-- End .carousel-prev -->"&_
						"<div id=""hot-items-slider-next"" class=""carousel-btn carousel-btn-next carousel-space"">"&_
						"</div><!-- End .carousel-next -->"&_
					"</div><!-- End .carousel-controls -->"&_
					"<div class=""hot-items-slider owl-carousel"">"
		n_articoli=0
		stringa=stringa&box_rs_articoli(rs_articoli,0)
		
		stringa=stringa&_
					"</div><!-- End .hot-items-slider -->"&_
				"</div><!-- End .hot-items .carousel-wrapper -->"
	end if
    set rs_articoli=Nothing    							
	hot_items=stringa&"<!-- END hot_items-container -->"&vbcrlf					
end function
	
	
function new_items()
	dim stringa
	stringa="<!-- new_items-container -->"&vbcrlf
	'set rs_articoli=conn.execute("select prodotti.* from prodotti where ((datamod is not null) and prodotti.novita=true and (prodotti.novdata>date() or (prodotti.novdata is null) )) order by datamod desc LIMIT 0,12")
	set rs_articoli=conn.execute("select prodotti.* from prodotti where  prodotti.novita=1 and visibilita=0 order by datamod desc LIMIT 0,12")
	if not rs_articoli.eof then
		stringa=stringa&"<div class=""new-items carousel-wrapper"">"&_
					"<header class=""content-title"">"&_
						"<div class=""title-bg"">"&_
							"<h2 class=""title"">"&traduci("nuovi")&"</h2>"&_
						"</div><!-- End .title-bg -->"&_
						"<p class=""title-desc"">"&traduci("nuovitxt")&"</p>"&_
					"</header>"&_
					"<div class=""carousel-controls"">"&_
						"<div id=""new-items-slider-prev"" class=""carousel-btn carousel-btn-prev"">"&_
						"</div><!-- End .carousel-prev -->"&_
						"<div id=""new-items-slider-next"" class=""carousel-btn carousel-btn-next carousel-space"">"&_
						"</div><!-- End .carousel-next -->"&_
					"</div><!-- End .carousel-controls -->"&_
					"<div class=""new-items-slider owl-carousel"">"
		n_articoli=0
		stringa=stringa&box_rs_articoli(rs_articoli,0)
		stringa=stringa&_
					"</div><!-- End .new-items-slider -->"&_
				"</div><!-- End .new-items .carousel-wrapper -->"
	end if
    set rs_articoli=Nothing    							
	new_items=stringa						
end function
function contenuto_libero(numero)
	dim stringa
	stringa=vbcrlf&"<!-- CONTENUTO"&numero&"-->"
	sql="select news.* from news where tipo='contenuto"&numero&"' and visualizza=1 order by ordine"
	set rs_cont=conn.execute(sql)
	
	do while not rs_cont.eof
		stringa=stringa& rs_cont("testo")
		rs_cont.MoveNext
	Loop							
	set rs_cont=Nothing
	contenuto_libero=stringa
end function
function slider_rev()
	if application("cache_slider_rev")="" then
		dim stringa
		set rs_dati_vari=conn.execute ("select dati_vari.* from dati_vari where chiave='index-slider'")
		if not rs_dati_vari.eof then
			dati_vari=split(rs_dati_vari("valore"),"||")
			stringa="<div id=""slider-rev-container"">"&_
						"<div id=""slider-rev"">"&dati_vari(0)&"</div><!-- End #slider-rev -->"&_
					"</div><!-- End #slider-rev-container -->"
		end if
		set rs_dati_vari=Nothing
		application("cache_slider_rev")=stringa
	end if
	slider_rev=application("cache_slider_rev")
end function
function marcamod() 
	dim stringa
	stringa="<div id=""marca-modello-container"" >"&_
				"<header class=""content-title"">"&_
					"<h2 class=""title"">"&traduci("marcamodtit")&"</h2>"&_
				"</header>"&_
				"<p>"&traduci("marcamodtxt")&"</p>"&_
				"<div class=""row"">"&_
					"<div class=""col-md-4 col-sm-6 col-xs-12"">"&_
						"<div class=""input-group select2-bootstrap-prepend"">"&_
							"<span class=""input-group-addon""><span class=""input-text"">"&traduci("marca")&"</span></span>"&_
							"<input type=""text""  class=""form-control input-lg"" id=""marca"" placeholder="""&traduci("marcapl")&""" value="""" >"&_
						"</div><!-- End .input-group -->"&_
					"</div><!-- End .col-md-4 -->"&_
					"<div class=""col-md-4 col-sm-6 col-xs-12"">"&_
						"<div class=""input-group select2-bootstrap-prepend"">"&_
							"<span class=""input-group-addon""><span class=""input-text"">"&traduci("serie")&"</span></span>"&_
							"<input type=""text""  class=""form-control input-lg"" id=""serie"" placeholder="""&traduci("seriepl")&""" value="""" readonly>"&_
						"</div><!-- End .input-group -->"&_
					"</div><!-- End .col-md-4 -->"&_
					"<div class=""col-md-4 col-sm-6 col-xs-12"">"&_
						"<div class=""input-group select2-bootstrap-prepend"">"&_
							"<span class=""input-group-addon""><span class=""input-text"">"&traduci("modello")&"</span></span>"&_
							"<input type=""text""  class=""form-control input-lg"" id=""modello"" placeholder="""&traduci("modellopl")&""" value="""" readonly>"&_
						"</div><!-- End .input-group -->"&_
					"</div><!-- End .col-md-4 -->"&_
				"</div><!-- End .row -->"&_
			"</div><!-- End #marca-modello-container -->"
	marcamod=stringa	
end function
%>