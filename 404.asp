<!--#include virtual="/setup.asp" -->
<!--#include virtual="/config/header_inc.asp" -->

<%
dim logga_evento
logga_evento=true 
codice_errore=session("codice_errore")
if codice_errore="" then codice_errore=0
if codice_errore=5 then
	'Aggiorno in ip_bannati
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "select ip_bannati.* from ip_bannati where ip='"&Request.ServerVariables("REMOTE_ADDR")&"'", conn, 3, 3
	if not rs.eof then
		rs("ultima_visita")=now()
		rs("numero_visite")=rs("numero_visite")+1
		rs.Update 
	else
		add2log "IP "&Request.ServerVariables("REMOTE_ADDR")&" non trovato in tabella",3
		call banna_ip()
	end if
	rs.Close
	set rs=Nothing

end if

%>
        <section id="content" class="no-content">
        	<div class="lg-margin"></div><!-- Space -->
        	<div class="container">
        		<div class="row">
        			<div class="col-md-12">
        				<div class="no-content-comment">
                            <h2>Errore</h2>
                            <%
                            select case codice_errore
                             case 1 
	                            response.status="401 Unauthorized"
                            %>
                            <h4>L'ordine che hai cercato di visualizzare non esiste.</h4>
                            <%case 2
	                            response.status="401 Unauthorized"	                            
                            %>
                            <h4>L'ordine che hai cercato di visualizzare non ti appartiene.<br><br>
                            L'errore &egrave; stato registrato.</h4>
                            <%case 3,4
	                            response.status="401 Unauthorized"
                            %>
                            <h4>Il contenuto che hai cercato di visualizzare<br>non esiste.</h4>
                            <%case 5
	                            logga_evento=false
                            %>
                            <h4>Your IP address<br>is banned.</h4>
                            <%case 6
	                            response.status="401 Unauthorized"
                            %>
                            <h4>Il documento non esiste.</h4>
                            <%case 7,8
							'7 pdf ddt
							'8 pdf fattura
	                            response.status="401 Unauthorized"
							%>
                            <h4>Non sei autorizzato a visualizzare questo contenuto.<br>
                            L'evento &egrave; stato registrato.</h4>
                            
                            <%case 9
	                            ' Prodotto inesistente
	                            response.status="410 Gone"
	                            logga_evento=false
	                            testo_errore="Tentato di visualizzare prodotto inesistente"
	                            oggetto_errore="idpro="&session("oggetto_errore")
                            %>
                            <h4>L'articolo non è più disponibile</h4>
                            <%case 10
	                            logga_evento=false
	                            response.status="410 Gone"
	                            testo_errore="Tentato di visualizzare prodotto non visibile"
	                            oggetto_errore=testo_log_articolo(session("oggetto_errore"))
                            %>
                            <h4>L'articolo non è più disponibile</h4>
                            <%case 11
	                            response.status="410 Gone"
	                            testo_errore="Richiamata pagina prodotto con riferimento prodotto non valido"
	                            oggetto_errore="idpro="&session("oggetto_errore")
                            %>
                            <h4>La pagina non esiste</h4>
                            <%Case Else
	                            response.status="404 Pagina inesistente"
	                            
                            %>
                            <h4>Si &egrave; verificato un errore.<br><br>
                            L'errore &egrave; stato registrato.</h4>
                            <%end select%>
                        </div><!-- End .no-content-comment -->
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
			</div><!-- End .container -->
        
        </section><!-- End #content -->
<%
	if testo_errore="" then testo_errore=codice_errore
	if oggetto_errore="" then oggetto_errore=session("oggetto_errore")
	if codice_errore>0 and logga_evento=true then
		add2log "Reindirizzamento a pagina di errore per: "&testo_errore&", oggetto errore:"&oggetto_errore&vbcrlf&"queryeform:"&queryeform()&vbcrlf&"IP:"&Request.ServerVariables("REMOTE_ADDR")&vbcrlf & allCookies()&vbcrlf&"HTTP_REFERER:"&Request.ServerVariables("HTTP_REFERER")&vbcrlf&"Pagina errore"&session("pagina_errore"),1
	end if
session("codice_errore")=""
session("oggetto_errore")=""
session("pagina_errore")=""
%>        
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
    <!-- AGGIUNTE -->


    </body>
</html>
<%
function testo_log_articolo(idpro)
	testo_log_articolo=""
	if isnumeric(idpro) and not isnull(idpro) then
        set rs=conn.execute("select prodotti.codice from prodotti where idpro="&idpro)
        testo_log_articolo="[articolo="& idpro&"]"&rs("codice")&"[/articolo]"
        Set rs = Nothing
    end if

end function	
	
%>
