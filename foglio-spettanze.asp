<%t_inizio=timer()%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->
<%
	
	
if not utente_user then response.redirect "login.asp"
iduser=session("iduser")
idfoglio=request("idfoglio")


if Request.querystring("nuovo")<>"" then
	Set rs_foglio = Server.CreateObject("ADODB.Recordset")
	rs_foglio.open "utenti_foglio_spettanze",conn,3 ,3
	rs_foglio.addnew
	rs_foglio("iduser")=iduser
	rs_foglio("creato")=date()
	rs_foglio.update
	rs_foglio.Close
	idfoglio=Get_last_id("utenti_foglio_spettanze")
	response.redirect("foglio-spettanze.asp?idforglio="&idforglio)
end if

if request.form("elimina")<>"" then
	conn.execute("delete from utenti_spettanze where idfoglio="&idfoglio)
	conn.execute("delete from utenti_foglio_spettanze where id="&idfoglio)
	response.redirect("fogli-spettanze.asp")
end if

'if request.querystring("iduser")<>"" and session("idadmin")<>"" then
'	iduser=request.querystring("iduser")
'end if	
aggiornato=false
if request.form("aggiungi")<>""  or request.form("completa")<>"" then
	call aggiorna_spettanze()
	aggiornato=true
end if
disabled=""
foglio_spettanze=true
visualizza_tabella=true
dim array_totali()
if IsNumeric(idfoglio) and idfoglio<>"" then 
	sql="select utenti_foglio_spettanze.* from utenti_foglio_spettanze where id="&idfoglio
else
	response.redirect "fogli-spettanze.asp"
end if
set rs=conn.execute (sql)
if not rs.eof then
	
	
	tipo_ordine=rs("tipo_ordine")
	cig=rs("cig")
	determinazione=rs("determinazione")
	note=rs("note")
	completo=rs("completo")
	iduser=clng(rs("iduser"))
	creato=rs("creato")
	if iduser<>session("iduser") and not utente_admin then
		call esci(3,idfoglio)	
	end if
	
	if completo="" or isnull(completo) then
		boolcompleto=false
	else
		boolcompleto=true
	end if
	if completo<>"" then
		disabled=" disabled "
	end if
	idfoglio=rs("id")
	
	if utente_admin then
		boolcompleto=true
		end if
	
else
	tipo_ordine=-1
	cig=""
	determinazione=""
	note=""
	completo=""
	
	idfoglio=0
end if
alert_display="style=""display:none;"""
if tipo_ordine=0 then
	alert_display=""
end if
ordinaper=Request.QueryString("ordina")
if ordinaper="" then
	ordinaper=Request.Cookies(questofile)("ordina")
end if
if ordinaper="codice" then
	orderby=" codice"
	Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
	Response.Cookies(questofile)("ordina")=ordinaper
elseif ordinaper="articolo" then
	orderby="  articolo"
	Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
	Response.Cookies(questofile)("ordina")=ordinaper
else
	orderby="  codice"
end if
dividi_settori=request.querystring("settori")
if dividi_settori="" then
	dividi_settori=Request.Cookies(questofile)("settori")
	end if
if dividi_settori="si" then
	orderby="nome_settore ,"&orderby
		Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
		Response.Cookies(questofile)("settori")=dividi_settori
	elseif dividi_settori="no" then
		Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
		Response.Cookies(questofile)("settori")=dividi_settori
	else
end if%>
<!--#include virtual="/config/header_inc.asp" -->
<section id="content">
	        	<div id="breadcrumb-container">
        		<div class="container">
					<ul class="breadcrumb">
						<li><a href="/">Home</a></li>
						<li class=""><a href="/fogli-spettanze.asp">Fogli spettanze</a></li>
						<li class="">Foglio spettanze del <%=creato%></li>

					</ul>
        		</div>
        	</div>

	<div class="container">
	    <form method="POST" action="<%=questofile%>" >
			<div class="row">
				<div class="col-md-12">
					<header class="content-title">
						<h1 class="title">Foglio spettanze del <%=creato%></h1>
					</header>
					<%if aggiornato and session("errore")="" then %>
					<div class="alert alert-success" role="alert"><button type="button" class="close" data-dismiss="alert">x</button> <strong>Fatto!</strong> I dati sono stati salvati. </div>
				<%end if %>
				<%if session("errore")<>"" then %>
					<div class="alert alert-danger" role="alert"> <strong>Attenzione:</strong> <%=session("errore")%>
					</div>				
				
				<%
					session("errore")=""
					end if %>	
					
	
				<%if boolcompleto then %>
					<div class="alert alert-info" role="alert"> <strong>Attenzione:</strong> Il foglio spettanze è completato, non può essere modificato. <a href="pdf-foglio-spettanze.asp?idfoglio=<%=idfoglio%>" class="btn btn-custom">SCARICA la versione PDF</a> 
					</div>
				<%end if %>
	        	</div><!-- col-md-12-->			
			</div><!-- End .row -->
       				
			<div class="xs-margin"></div><!-- space -->
			
			<input type="hidden" name="idfoglio" value="<%=idfoglio%>">
				<fieldset>
				<div class="row">
					<div class="col-md-6 col-sm-6 col-xs-12">
						<h2 class="checkout-title">Dati generali</h2>
					</div>
				</div><!-- End .row -->
					
				<div class="row">
					<div class="col-md-6 col-sm-6 col-xs-12">
						<div class="input-group">
							<span class="input-group-addon"><span class="input-text">TRATTATIVA DIRETTA MePA con PREZZO A CORPO</span></span>
							<span class="form-control input-lg no-minwidth">
							<input type="radio" name="tipo_ordine" value="0" class="tipo_ordine" <%if tipo_ordine=0 then response.write "checked"%> <%=disabled%>>
							</span>
						</div>
						<!-- End .input-group -->
					</div>
					<div class="col-md-6 col-sm-6 col-xs-12">
						<div class="input-group">
							<span class="input-group-addon"><span class="input-text">ORDINE DIRETTO FUORI MePA</span></span>
							<span class="form-control input-lg no-minwidth">
							<input type="radio" name="tipo_ordine" value="1" class="tipo_ordine" <%if tipo_ordine=1 then response.write "checked"%> <%=disabled%>>
							</span>
						</div>
						<!-- End .input-group -->
					</div>
				</div><!-- End .row -->
				<div class="alert alert-danger" role="alert" id="alert_tipo_ordine" <%=alert_display%>>
					<strong>Avvertenza:</strong> concludere la procedura di stipula entro la data da voi fissata, altrimenti dovrete eseguire nuovamente il caricamento su MePA.
				</div>
							
				<div class="row">
					<div class="col-md-6 col-sm-6 col-xs-12">
						<div class="input-group">
							<span class="input-group-addon"><span class="input-text">Cig</span></span>
							<input type="text"  class="form-control input-lg" placeholder="" name="cig" value="<%=cig%>" <%=disabled%>>
						</div><!-- End .input-group -->
					</DIV><!-- col-md6 -->
					<div class="col-md-6 col-sm-6 col-xs-12">
						<div class="input-group">
							<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text">Determinazione</span></span>
							<input type="text"  class="form-control input-lg" placeholder="" name="determinazione" value="<%=determinazione%>" <%=disabled%>>
						</div><!-- End .input-group -->
					</DIV><!-- col-md6 -->
				</div><!-- End .row -->
				<div class="row">
					<div class="col-md-12 col-sm-12 col-xs-12">
						<div class="input-group textarea-container">
							<span class="input-group-addon"><span class="input-text">Note</span></span>
							<textarea name="note" id="contact-message" class="form-control" cols="30" rows="6" placeholder="" <%=disabled%>><%=note%></textarea>
						</div>
					</div><!-- End .col-md-12 -->
				</div><!-- End .row -->
							
				</fieldset>  	
	
				<div class="row" style="margin-top:15px;">
					<div class="col-md-12 col-sm-12 col-xs-12">
		<%
			
		
			
		    	'Query con scorta tecnica
			    ' set rs_dipendenti=conn.execute ("(select dipendenti.iddip, cognome, nome,sesso from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.dimesso=0 and utenti_dipendenti.iduser="&iduser&" order by sesso desc, cognome,nome) union ( select 0 as iddip , 'Scorta tecnica' as cognome,'' as nome, 'A' as sesso ) order by sesso desc, cognome, nome")
			    if boolcompleto then
					sql="select dipendenti.iddip, cognome, nome, sesso from dipendenti inner join (select distinct iddip from utenti_spettanze where idfoglio= "&idfoglio&") as t on dipendenti.iddip = t.iddip  order by  cognome, nome"
			    else
				   sql= "select dipendenti.iddip, cognome, nome,sesso from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.dimesso=0 and utenti_dipendenti.iduser="&session("iduserDipendenti")&" order by cognome,nome"
				end if
			    set rs_dipendenti=conn.execute (sql)
			    if rs_dipendenti.eof then
				    visualizza_tabella=false
				    %>
				    <div class="alert alert-warning" role="alert"> <strong>Attenzione!</strong> Per generare il foglio delle spettanze devono essere aggiunti dei dipendenti.
					    <p>
						    <a href="dipendenti.asp" class="btn btn-custom">Aggiungi dipendenti</a>
					    </p>    
					 </div>
				    <%
				end if    
			    if boolcompleto then
					sql="select 0 as idpre, prodotti.idpro, prodotti.codice, prodotti.articolo, prodotti.prezzo from prodotti INNER JOIN (select distinct idpro from  utenti_spettanze where idfoglio="&idfoglio&") as t on prodotti.idpro=t.idpro order by "&orderby
				else
				
				if idfoglio="" then idfoglio=0
				sql="select * from prodotti where idpro in (  SELECT idpro FROM `preferiti` WHERE iduser="&iduser&" union distinct select idpro from utenti_spettanze where idfoglio="&idfoglio&") order by "&orderby
				
				
				
				
				
				'sql="select preferiti.idpre, prodotti.idpro, prodotti.codice, prodotti.articolo, prodotti.prezzo from preferiti INNER JOIN prodotti on preferiti.idpro=prodotti.idpro where preferiti.iduser="&iduser&" order by "& orderby
				end if
					'if utente_andrea then response.write sql
				set rs_prodotti=conn.execute(sql)
			    if rs_prodotti.eof then
				    visualizza_tabella=false
				    %>
				    <div class="alert alert-warning" role="alert"> <strong>Attenzione!</strong> Per generare il foglio delle spettanze devono essere aggiunti gli articoli ai preferiti.
					    <p>
						    Sfoglia il nostro catalogo prodotti e aggiungi gli articoli ai preferiti. <a href="category.asp" class="btn btn-custom">Sfoglia il catalogo</a>
					    </p>    
				     </div>
				    <%
				end if    
	
			if visualizza_tabella then
			    arrDipendenti=rs_dipendenti.getrows()
				
			    ubound_arrDipendenti=ubound(arrDipendenti,2)
				redim array_totali(ubound_arrDipendenti)
				for n = 0 to ubound_arrDipendenti
					array_totali(n)=0
				Next	
				importoTotale=0
				
				if boolcompleto=false then
	%>	
	
	<div class="col-md-12" style="border: 1px solid black; background: white; margin-bottom: 5px;">
		Clicca sul nominativo dipendente per aggiungere le quantità<br>
		<%
		for n = 0 to ubound_arrDipendenti
			if n>0 then response.write ",&nbsp;"
			response.write "<a href=""quantitadipendente.asp?foglio="&idfoglio&"&iddip="&arrDipendenti(0,n)&""">"& nominativo_breve(arrDipendenti(1,n),arrDipendenti(2,n))&"</a>"
			
		Next	

			
			%>
	</div>
	<%
		end if
		%>
	
	
				<div class="dropdown pull-right" style="margin-bottom: 10px; " >
				  <button class="btn btn-default dropdown-toggle" type="button" id="dropdownMenu1" data-toggle="dropdown" aria-haspopup="true" aria-expanded="true">
				ORDINA PER    <span class="caret"></span>
				  </button>
				  <ul class="dropdown-menu" aria-labelledby="dropdownMenu1">
					<li><a href="<%=questofile%>?ordina=codice">Codice</a></li>
					<li><a href="<%=questofile%>?ordina=articolo">Articolo</a></li>
					<%if false then %>
					<li role="separator" class="divider"></li>
					<%if dividi_settori="si" then %>
					<li><a href="<%=questofile%>?settori=no">Non dividere in categorie</a></li>
					<% else %>
					<li><a href="<%=questofile%>?settori=si">Dividi in categorie</a></li>
					<%end if %>
					<%end if %>
				  </ul>
				</div> 
				</div><!-- col-md-12 -->
				</div><!-- row -->
					<div class="col-md-12" id="tablefull">
						<a href="#" id="fullscreen" class="btn btn-info btn-xs">Modalità full-screen</A>
									<table id="example" class="display" style="width: 100%;" name="tabella">
										<thead>
											<tr>
												<th class="zindex500">Articolo</th>
												<%
												for n = 0 to ubound_arrDipendenti
												%>
												<th class="dipendenti" id="iddip<%=arrDipendenti(0,n)%>" ><span class="header_height"><a href="quantitadipendente.asp?foglio=<%=idfoglio%>&iddip=<%=arrDipendenti(0,n)%>"><%= nominativo_breve(arrDipendenti(1,n),arrDipendenti(2,n))%></a></span></th>
												<%
												Next		            
												%>
										  </tr>
									</thead>
									<tbody>
									<%
													do while not rs_prodotti.EOF
													html_riga=""
														articolo="<a href=""/product.asp?idpro="&rs_prodotti("idpro")&""" target=""_blank"">"&rs_prodotti("codice")&" "&rs_prodotti("articolo")&"</a>"
														prezzo=cdbl(rs_prodotti("prezzo"))
														quantita_tot=0
														'Loop su dipendenti
														for n = 0 to ubound_arrDipendenti
															testocella=""
															sql="select utenti_spettanze.* from utenti_spettanze where idfoglio="&idfoglio&" and idpro= "&rs_prodotti("idpro")&" and iddip="&arrDipendenti(0,n)
															set rs_spettanze=conn.execute(sql)
															if not rs_spettanze.eof then
																quantita=rs_spettanze("quantita")
																sub_totale=quantita*prezzo
																array_totali(n)=array_totali(n)+sub_totale
																importoTotale=importoTotale+sub_totale
																'id=rs_spettanze("id")
															else
																quantita=0
																'id=""
															end if
															if quantita=0 then quantita=""
																
															testocella="<input type=""text"" name=""quantita_"&rs_prodotti("idpro")&"_"&arrDipendenti(0,n)&""" class=""spett"" "&disabled&"value="""&quantita&""">"'&id
															
															
testocella=quantita
															html_riga=html_riga&"<td  class=""quantita"" id=""quantita_"&rs_prodotti("idpro")&"_"&arrDipendenti(0,n)&""">"&testocella&"</td>"
														Next	
														%>
													<tr >
													
													<td class="<%=classe_articolo_v%> articoli idpro<%=rs_prodotti("idpro")%>"><%=articolo%><%if session("vedi_prezzi")=1 or utente_admin then %><span class="pull-right prezzo"><%=simbolo_valuta&rs_prodotti("prezzo")%></span><%end if %></td>
													<%=html_riga%>
													
													
													</tr>
													<%
													rs_prodotti.MoveNext
													Loop
													
													set rs_prodotti = Nothing
									%>
									
									</tbody>
									<%if session("vedi_prezzi")=1 or utente_admin then %>
									<tfoot >
										<tr class="totalidipendenti">
											<td  >
												<span>Totale iva esclusa per dipendente</span>
											</td>
													<%
													for n = 0 to ubound_arrDipendenti
														response.write "<td class=""quantita footer"" style=""font-size: smaller;""><span class=""footer_height"">"&simbolo_valuta&formatnumber(array_totali(n),2)&"</span></td>"
													Next	
													%>
										</tr>
									</tfoot>
									<%end if %>
									</table>
							<%if session("iduser")=iduser and disabled=""  then %>
							<div class="row" style="text-align: center;">
									<input type="submit" name="aggiungi" class="btn btn-success" Value="Salva modifiche" style="margin-top:3px; margin-bottom:3px;">
							</div><!-- row -->
							<%end if %>
						<%if session("vedi_prezzi")=1 or utente_admin then %>
							<table style="width:100%; margin-top: 10px;">
											<%
												trasporto=12
												%>
								<tr class="totali" style="border-top: 1px solid black;">
									<td align="right"  width="90%"><strong>
										Trasporto: </strong>
									</td>
									<td align="right" ><strong>
									<%=simbolo_valuta&formatnumber(trasporto,2)%></strong>
									</td>
								</tr>
										<%
											importoTotale=importoTotale+trasporto
											%>
								<tr class="totali" style="border-top: 1px solid black;">
									<td align="right"  ><strong>
										Totale iva esclusa: </strong>
									</td>
									<td align="right" ><strong>
										<%=simbolo_valuta&formatnumber(importoTotale,2)%></strong>
									</td>
								</tr>
										<%
											valore_iva=RoundUp(importoTotale*iva(date())/100, 2)
											%>
								<tr class="totali">
									<td align="right" ><strong>
										Iva <%=iva(date())%>%:</strong>
									</td>
									<td align="right" ><strong>
										<%=simbolo_valuta&formatnumber(valore_iva,2)%></strong>
									</td>
								</tr>
										<%
											totale=importoTotale+valore_iva
											%>
								<tr class="totali">
									<td align="right" ><strong>
										Totale: </strong>
									</td>
									<td align="right" ><strong>
										<%=simbolo_valuta&formatnumber(totale,2)%></strong>
									</td>
								</tr>
								<tr id="avviso" style="display: none;">
									<td align="center" colspan="2" >
										I dati sono cambiati, salvare le modifiche per aggiornare i totali
									</td>
								</tr>
							</table>
							<%end if 
								
								
												if boolcompleto=false then
	%>	
	
	<div class="col-md-12" style="border: 1px solid black; background: white; margin-bottom: 5px; margin-top:10px;">
		Clicca sul nominativo dipendente per aggiungere le quantità<br>
		<%
		for n = 0 to ubound_arrDipendenti
			if n>0 then response.write ",&nbsp;"
			response.write "<a href=""quantitadipendente.asp?foglio="&idfoglio&"&iddip="&arrDipendenti(0,n)&""">"& nominativo_breve(arrDipendenti(1,n),arrDipendenti(2,n))&"</a>"
			
		Next	

			
			%>
	</div>
	<%
		end if
		%>

							<div class="row" style="margin-top:10px;">
								<%if session("iduser")=iduser and disabled="" then %>
									<div class="col-md-2 col-sm-2 col-xs-12 " style="text-align: center;">
										<input type="submit" class="btn btn-info" name="completa" value="COMPLETA e SCARICA PDF"/>
									</div><!-- col-md-2 -->
									<div class="col-md-10 col-sm-10 col-xs-12">
										<div class="alert alert-info" role="alert"> <strong>Attenzione:</strong> Premendo sul pulsante COMPLETA le modifiche al foglio spettanze verranno disabilitate. Non potranno pi&ugrave; essere effettuate modifiche ai dati della tabella.<br>Sar&agrave; possibile scaricare il foglio in formato PDF che dovrete inviarci via email.<br>
											<ul>
												<li>In caso di <b>ORDINE DIRETTO</b> dovrete allegare i dati relativi all’impegno di spesa e copia della DETERMINAZIONE</li>
												<li>In caso di <b>TRATTATIVA DIRETTA MePA (cod. LR2500)</b> potete allegare il PDF per indicare l’importo base a corpo</li>
											</ul>
												
										 </div><!-- alert -->
									</div><!-- col-md-10 -->
								<% end if %>
							</div><!-- row -->
							
							<div class="row" style="margin-top:10px;">
								<%if (session("iduser")=iduser and disabled="") or utente_admin then %>
									<div class="col-md-2 col-sm-2 col-xs-12 " style="text-align: center;">
										<input type="submit" class="btn btn-danger" name="elimina" value="Elimina"/>
									</div><!-- col-md-2 -->
									<div class="col-md-10 col-sm-10 col-xs-12">
										<div class="alert alert-danger" role="alert"> <strong>Attenzione:</strong> Questa operazione è irreversibile.												
										 </div><!-- alert -->
									</div><!-- col-md-10 -->
								<% end if %>
							</div><!-- row -->

							
							
				<%end if
					set rs = Nothing
					call connclose()
		 %>
					</div><!-- #tablefull -->
		</form>
	</div>
			
</section>
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
<script type="text/javascript" src="Jquery/Js/jquery.dirtyforms.js"></script>
<script type="text/javascript" src="Jquery/Js/jquery.dirtyforms.dialogs.bootstrap.min.js"></script>
<!-- DATATABLE-->
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/v/bs/dt-1.10.16/fc-3.2.4/fh-3.1.3/datatables.min.css"/>
<script type="text/javascript" src="https://cdn.datatables.net/v/bs/dt-1.10.16/fc-3.2.4/fh-3.1.3/datatables.min.js"></script>

<script type="text/javascript" src="Jquery/Js/screenfull.min.js"></script>
<script type="text/javascript">
      $(function() {
	      
	      
	      
 		$('#fullscreen').click(function () {
			screenfull.toggle($('#tablefull')[0]);
		});
	  screenfull.on('change', function(){
		  console.log("#tablefull.height:"+$("#tablefull").height());
		 // $("#tablefull").scrollintoview();
		  /*
		  table.destroy();
		  
		  table = $('#example').DataTable( {
			"ordering": false,
	        "info":     false,
	         "searching": false,
	        scrollY:        "500px",
	        scrollX:        true,
	        scrollCollapse: true,
	        paging:         false,
	        fixedColumns:   {
	            leftColumns: 1
	        }
	    } );
		*/
		  
	  });
 
	      
	      
  	$(".quantita").hover(function(){
		
		$('.articoli').each(function(i, obj) {
			$(obj).css('background','');			
		});
		$('.dipendenti').each(function(i, obj) {
			$(obj).css('background','');			
		});
		
		
		var string=$(this).attr('id');
		console.log("string:"+string);
		
		var array = string.split('_');
		var idpro=array[1];
		var iddip=array[2];
  	
  	
		console.log("idpro:"+idpro+" iddip:"+iddip);
		$(".idpro"+idpro).each(function(i, obj) {
			$(obj).css("background",'#dfff26');
			console.log("trovato"+idpro);
		});
		$("#iddip"+iddip).css("background",'#dfff26');
		//$("#idpro"+idpro).css("background",'#dfff26');
  	//$("#iddip"+iddip).css("background",'#dfff26');
	});
	$(".tipo_ordine").click(function(){
		
	if($(this).val()=="0"){
		$("#alert_tipo_ordine").slideDown();
	}else
		$("#alert_tipo_ordine").slideUp();
	});
	      
	var header_height = 0;
    $('.header_height').each(function() {
        if ($(this).outerWidth() > header_height){
	        header_height = $(this).outerWidth();
        } 
    });
    $('table th').height(header_height);
	var footer_height = 0;
    $('.footer_height').each(function() {
        if ($(this).outerWidth() > footer_height){
	        footer_height = $(this).outerWidth();
        } 
    });
	$('.footer').height(footer_height);
    
    console.log("Altezza celle dipendenti:"+header_height);
	      
	//	$('#pippo').DataTable();
    $('form').dirtyForms({
			    ignoreSelector: '#fullscreen',
		message: 'Hai modificato dei dati ma non hai salvato le modifiche, se lasci la pagina le modifiche andranno perse.<br>Sei sicuro di voler abbandonare la pagina?',
		dialog: {
		title: 'Avvertenza',
		proceedButtonText:'Lascia la pagina',
		stayButtonText:'Rimani'
	}});
	var hidden=false;
	$(document).bind('dirty.dirtyforms', function(event) { 
		if(!hidden){
		
		console.log("dirty");
		$(".totali").hide();
		$(".totalidipendenti").each(function(e){
			$(this).hide();
		});
		$("#avviso").show();
		hidden=true;
		}
	});
		
		var table = $('#example').DataTable( {
			"ordering": false,
	        "info":     false,
	         "searching": false,
	        //scrollY:        "500px",
	        scrollY:        "70vh",
	        scrollX:        true,
	        scrollCollapse: true,
	        paging:         false,
	        fixedColumns:   {
	            leftColumns: 1
	        }
	    } );
	    $("#example tfoot").remove();
		
	});
    
	                
	                
    </script>
    </body>
</html>
<%
function classe_articolo(quantita,tot_in_ordine,tot_pronto,tot_consegnato)
	quantita=cdbl(quantita)
	classe_articolo=""
	if tot_consegnato=quantita then
		classe_articolo="consegnatos"
	elseif tot_pronto=quantita then
		classe_articolo="prontos"
	elseif tot_in_ordine=quantita then
		classe_articolo="in_ordines"
	end if
end function
	 sub aggiorna_spettanze()
		 
		 
		 completato=false
		 
		 
		completo=false
		sql="select count(*) from utenti_spettanze where idfoglio="&idfoglio
		
		cisonoquantita=clng(conn.Execute(sql)(0))
		if cisonoquantita=0 then
			session("errore")="Non ci sono quantit&agrave; caricate"

		end if
		Set rs_spettanza = Server.CreateObject("ADODB.Recordset")

		rs_spettanza.open "select * from utenti_foglio_spettanze where id="&idfoglio,conn,3 ,3
		rs_spettanza("tipo_ordine")=cint(request.form("tipo_ordine"))
		rs_spettanza("cig")=request.form("cig")
		rs_spettanza("determinazione")=request.form("determinazione")
	    m_log_txt=m_log_txt&"<br>determinazione:"&request.form("determinazione")
		rs_spettanza("note")=request.form("note")

		if request.form("completa")<>"" and cisonoquantita>0 then
			rs_spettanza("completo")=date()		
			completo=true	
		end if
		rs_spettanza.update
		rs_spettanza.close

		set rs_spettanza = Nothing
		
		
		
		call add2log(m_log_txt&" su foglio:"&idfoglio,0)
		if completo then
			denominazione_v=get_denominazione(iduser)
			call inserisci_appunto("<a href=""pag_adm_user.asp?iduser="&iduser&"""><b>"&denominazione_v&"</b></a> ha completato il foglio spettanze "&idfoglio,array(4,1))
		end if
		
		
	end sub	'aggiungi_spettanze()
	
%>
