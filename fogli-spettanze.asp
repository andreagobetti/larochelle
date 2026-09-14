<%
'Verifica chiusure 30_11_2015
%>
<!--#include virtual="/setup.asp" -->
<%

if not utente_user then response.redirect "login.asp"


set rs_fogli=conn.execute("select * from utenti_foglio_spettanze where iduser="&sessioniduser&" order by id")
if rs_fogli.eof then
	call vaiAlFoglio()
end if
pulsante_nuovo=true
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
						<li class="">Fogli spettanze</li>
					</ul>
        		</div>
        	</div>
        	<div class="container">
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title">Fogli spettanze</h1>
							<p class="title-desc">Elenco dei fogli spettanze</p>
						</header>
        				<div class="xs-margin"></div><!-- space -->
        				<div class="row">
        					
        					<div class="col-md-12 table-responsive" id="order-detail-content">
        						<table class="table cart-table" >
        						<thead>
        							<tr>
										<th class="table-title">Creato il</th>
										<th class="table-title">Completato il</th>
										<th class="table-title">Stato</th>
										<th class="table-title"></th>
        							</tr>
        						</thead>
								<tbody>
								<%
								do while not rs_fogli.EOF
									if rs_fogli("completo")="" or isnull(rs_fogli("completo")) then
										statoFoglio= "APERTO"
										pulsante_nuovo=false
									else
										statoFoglio= "COMPLETO"
									end if
								
									%>
									<tr>
										<td class="">
											<%=rs_fogli("creato")%>
										</td>
										<td class=""><%=rs_fogli("completo")%></td>
										<td class=""><%=statoFoglio %></td>
										<td class="">
											<%if statoFoglio="COMPLETO" then %>
											
											<a href="foglio-spettanze.asp?idfoglio=<%=rs_fogli("id")%>" class="btn btn-custom">VEDI</a>&nbsp; <a href="pdf-foglio-spettanze.asp?idfoglio=<%=rs_fogli("id")%>" class="btn btn-custom">Versione PDF</a> <% else %><a href="foglio-spettanze.asp?idfoglio=<%=rs_fogli("id")%>" class="btn btn-success">MODIFICA FOGLIO APERTO</a> <% end if %>
											
											
											</td>
									</tr>
									<%
									rs_fogli.MoveNext
									loop
									set rs_fogli = Nothing
									%>
								</tbody>
							  </table>
	        				<div class="xs-margin"></div><!-- space -->
	        				<%if statoFoglio ="COMPLETO" then %>
							  <div class="pull-right">
								  <a href="foglio-spettanze.asp?nuovo=nuovo" class="btn btn-success">NUOVO FOGLIO</a>
							  </div>
							<% end if %>

        					</div><!-- End .col-md-12 -->
        					
        				</div><!-- End .row -->
        				
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

    </body>
</html>
<%
	
	
	sub vaiAlFoglio()
		call connclose()
		response.redirect "foglio-spettanze.asp?nuovo=si"
		
	end sub
	
	%>