<!--#include virtual="/setup.asp" -->
<!--#include virtual="/config/header_inc.asp" -->
        <section id="content">
        	<div id="breadcrumb-container">
        		<div class="container">
					<ul class="breadcrumb">
						<li><a href="/">Home</a></li>
						<li class="active"><%=traduci("tit10")%></li>
					</ul>
        		</div>
        	</div>
        	<div class="container">
			<%
			set rs_articoli=conn.execute("select prodotti.*, preferiti.iduser, preferiti.IDpre FROM prodotti INNER JOIN preferiti ON prodotti.IDpro = preferiti.idpro where iduser="& session("iduser") )
			n_articoli=0
%>
			
			
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=traduci("tit10")%></h1>
							<%if not rs_articoli.eof then%>
							<p class="title-desc"><%=traduci("tit10txt")%></p>
							<%else %>
							<p class="title-desc"><%=traduci("noartic")%></p>
							<%end if %>
						</header>
						
        				<div class="xs-margin"></div><!-- space -->

                        <%
						response.write box_rs_articoli(rs_articoli,4)
                        %>
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
        		
        		
        	<%	'not rs_articoli.eof
			set rs_articoli = Nothing	
	        	 %>	
			</div><!-- End .container -->
        </section><!-- End #content -->
        
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
	<script src="carrello.js"></script>
    </body>
</html>
<!--#include virtual="/dettaglio_inc.asp" -->
