<!--#include virtual="/config/header_conf.asp" -->
<!--#include virtual="/head_inc.asp" -->
    <body>
    <div id="wrapper">
    	<header id="header">
			<%if utente_admin then %>
			<!--#include virtual="/sub_barra_adminsf2.asp" -->
			<%end if%>  			
    		<div id="header-top">
    			<div class="container">
    				<div class="row">
                        <div class="col-md-12">
        					<div class="header-top-left">
	        					
    							<div class="header-text-container">

									<%if session("iduser")="" then%>
        							<p class="header-link"><a href="login.asp"><%=traduci("login")%></a>&nbsp;<%=traduci("or")%>&nbsp;<a href="register-account.asp"><%=traduci("newaccount")%></a></p>
									<%else %>
									<p class="header-text"><%=traduci("welcome")%> <strong><%=session("nominativo")%></strong></p>
									<%end if%>
    							</div><!-- End .pull-right -->
    							

        						
        					</div><!-- End .header-top-left -->
							<%if session("iduser")<>"" then%>
							
							<div class="btn-group dropdown-language">
    									<button type="button" class="btn btn-custom dropdown-toggle btn-success" data-toggle="dropdown">
    										Gestione utente
    									</button>
    									<ul class="dropdown-menu pull-right" role="menu">

    										<li><a href="my-account.asp" title="I miei dati">I miei dati</a></li>
    										
    										<li><a href="index.asp?logout=si">Log out</a></li>
    									</ul>
    								</div><!-- End .btn-group -->
							
							
							<div class="header-top-right">
        						<ul id="top-links" class="clearfix">
	        						<li><a href="pdf_elenco_prodotti.asp">Scarica elenco prodotti</a></li>
        							<li><a href="products-favorites.asp" title="<%=traduci("preferiti")%>"><span class="top-icon top-icon-check"></span><span class="">Articoli preferiti</span></a></li>
									<%if utente_ha_il_permesso("A1") then %>
									<li><a href="dipendenti.asp" title="Dipendenti"><span class="top-icon top-icon-user"></span><span class="">Dipendenti</span></a></li>
									<li><a href="fogli-spettanze.asp" title="Foglio spettanze"><span class="top-icon top-icon-pencil"></span><span class="">Foglio spettanze</span></a></li>
									<%end if %>
									
									
        						</ul>
        					</div><!-- End .header-top-left -->
        					
        					
        					
        					
							<%end if%>
							
        					
        					<div class="header-top-right">
        						
        						<div class="header-top-dropdowns pull-right">
									<%if session("idadmin")<>"" then%>
                                	<div class="btn-group dropdown-language">
    									<button type="button" class="btn btn-custom dropdown-toggle btn-success" data-toggle="dropdown">
    										Amministra
    									</button>
    									<ul class="dropdown-menu pull-right" role="menu">

    										<li><a href="pag_adm_dizionario.asp" target="_blank"><span class="hide-for-xs">Dizionario</span></a></li>
    									</ul>
    								</div><!-- End .btn-group -->
									<%end if%>
    								
    							</div><!-- End .header-top-dropdowns -->

    						</div><!-- End .header-top-right -->
    						
    						
    						
    						
    					</div><!-- End .col-md-12 -->
    				</div><!-- End .row -->
    			</div><!-- End .container -->
    		</div><!-- End #header-top -->
    		
    		
    		
    		<div id="inner-header">
    			<div class="container">
    				<div class="row">
						<div class="col-md-12 col-sm-12 col-xs-12 logo-container">
							<h2 class="logo clearfix">
								<span><%=descrizione_sito%></span>
								<a href="/" title="<%=descrizione_sito%>">
								<img src="/images/LOGO-PER-SITO-0218.png" class="logo" alt="La Rochelle" title="La Rochelle">			</a>					
								
							</h2>
						</div><!-- End .col-md-5 -->

    				</div><!-- End .row -->
    			</div><!-- End .container -->
    			
    			
    			
    			<div id="main-nav-container">
    				<div class="container">
    				
    					<div class="row">
    						<div class="col-md-12 clearfix">
    							
    							<nav id="main-nav">
    								<div id="responsive-nav">
    									<div id="responsive-nav-button">
											Menu <span id="responsive-nav-button-icon"></span>
										</div><!-- responsive-nav-button -->
    								</div>
									<ul class="menu clearfix">
									<%
									if session("idadmin")<>"" then
										response.write menu_admin
									end if
						        	if application("cache_menu"&lingua)="" then
							        	testo=""
							        	call menu(0,0,testo,"")
							        	application.lock
							        	application("cache_menu"&lingua)=testo
							        	Application.unlock
							        	'add2log "Ricreata cache menu"&lingua,1
									end if							        	
						        	response.write application("cache_menu"&lingua)
									%>
									</ul>
    							</nav><!-- End #main-nav -->
    							
    							<div id="quick-access">
	    							
    								<div class="dropdown-cart-menu-container pull-right" id="carrellino">
	    							<%
		    							call carrellino()
	    							%>
									</div><!-- End .dropdown-cart-menu-container -->
									
	    							<form class="form-inline quick-search-form" role="form" action="category.asp" method="post">
										<div class="form-group">
											<input type="text" id="cerca" name="testocerca" class="form-control" placeholder="<%=traduci("searchhere")%>">
										</div><!-- End .form-inline -->
										<button type="submit" id="quick-search" class="btn btn-custom"></button>
									</form>
    							</div><!-- End #quick-access -->

    						</div><!-- End .col-md-12 -->
						</div><!-- End .row -->
	    			</div><!-- End .container -->
    				
    			</div><!-- End #main-nav-container -->
    		</div><!-- End #inner-header -->
    	</header><!-- End #header -->