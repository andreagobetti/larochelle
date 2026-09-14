<!--#include virtual="/md5-2.asp"-->
<!--#include file="ClasseOrdine.asp"-->

<%
		dim subtot_carrello
		dim ordinato
		ordinato=false
		dim h
		set h = new MD5
		err=request("err")
		if request("preventivo")<>"" then preventivo=true else preventivo=false
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/mail_registrazione_inc.asp" -->
<%
new_user=true
iduser=0
if  request.form("email_login")<>"" and request.form("password_login")<>""  then 
		call login()
else
	
	invio_email_login=false
	if request.form("ordina")<>"" then
		dim strtimer
		dim timerstr
		
		
		
		
		
		
		
		timertmp=timer()
		timerstr="Inizio "&time()&"<br>"
		
		'Controllo carrello
		if utente_user  then
			where_carrello="iduser="&sessionIDuser
		elseif SessionIdGuest<>"" then
			where_carrello="idloged="&SessionIdGuest
		else
			where_carrello="iduser=-1"
			add2log "Combinazione non contemplata in controllo carrello"&vbcrlf&queryeform(),3
			
		end if
		sql="select count(*) from carrello where "&where_carrello
		count_carrello=clng(conn.execute(sql)(0))
		if count_carrello=0 then
			txt_errore=txt_errore&"Il tuo carrello &egrave; vuoto"
			call add2log("Carrello vuoto, sql:"&sql,3)
		end if
		
		'controllo piva e cf
		if utente_user then
			a_utente="[utente="&sessionIDuser&"]"&session("nominativo")&"[/utente]"
		else
			a_utente="utente non loggato"
		end if
		if trim(request.form("cf"))<>"" then
			cf=ucase(trim(request.form("cf")))
			call determina_sqlinjection(cf)
			controllo_cf="select iduser, nome, cognome, azienda, cf from utenti where cf='"& cf&"'"
			if utente_user then controllo_cf=controllo_cf& " and iduser<>"&sessionIDuser
			set rs=conn.execute(controllo_cf)
			if not rs.eof then
				txt_errore="Codice Fiscale gi&agrave; utilizzato da un altro utente."
				add2log "Bloccato ordine a "&a_utente&" per Codice Fiscale "&ucase(trim(request.form("cf")))&" gi&agrave; utilizzato da [utente="&rs("iduser")&"]"&nominativo(rs("nome"),rs("cognome"),rs("azienda"))&"[/utente]",2
			end if
			rs.close
			set rs=nothing
		timerstr=timerstr&"Ricerca cf: "&formatnumber(timer()-timertmp,3)&"<br>"

		end if
		timertmp=timer()
		if trim(request.form("piva"))<>"" then
			piva=trim(request.form("piva"))
			call determina_sqlinjection(piva)
			controllo_piva="select * from utenti where piva='"& piva&"'"
			if session("iduser")<>"" then controllo_piva=controllo_piva& " and iduser<>"&session("iduser")
			set rs=conn.execute(controllo_piva)
			if not rs.eof then
				if txt_errore<>"" then txt_errore=txt_errore&"<br>"
				txt_errore=txt_errore&"Partita Iva gi&agrave; utilizzata da un altro utente."
				add2log "Bloccato ordine a "&a_utente&" per Partita Iva "&trim(request.form("piva"))&" gi&agrave; utilizzata da [utente="&rs("iduser")&"]"&nominativo(rs("nome"),rs("cognome"),rs("azienda"))&"[/utente]",2
			end if
			rs.close
			set rs=nothing
			timerstr=timerstr&"Ricerca piva: "&formatnumber(timer()-timertmp,3)&"<br>"
		end if
		timertmp=timer()
		if trim(request.form("email"))="" and session("iduser")="" then
				if txt_errore<>"" then txt_errore=txt_errore&"<br>"
				txt_errore=txt_errore&"L'indirizzo email non è valido."
				add2log "Bloccato ordine a "&trim(request.form("cognome"))&" "&trim(request.form("nome"))&" "&trim(request.form("azienda"))&" per email non valida: "&trim(request.form("email")),2
		end if
			
	end if
	if request.form("ordina")<>"" and txt_errore="" then
		session("log_checkout")=now()&" Iniziata procedura checkout"&vbcrlf&queryeform()
	
	end if
	'if request("nuovo_utente")="si" and new_user and txt_errore="" then
	if request.form("ordina")<>"" and txt_errore="" and session("iduser")="" then
		'Controllo email univoca
		email=lcase(trim(request.form("email")))
		if email<>"" then
			if not validemail(email) then
				txt_errore="L'indirizzo email non è valido"
			end if
			call determina_sqlinjection(email)
			if txt_errore="" then
				set rs=conn.execute("select utenti.email from utenti where email='"&email&"'")
				if not rs.eof then
					txt_errore="L'indirizzo email è gi&agrave; esistente"
					add2log "Bloccata registrazione in fase di ordine per email gi&agrave; esistente:"&email,2
				end if
				rs.close
				set rs=nothing
				timerstr=timerstr&"Ricerca email: "&formatnumber(timer()-timertmp,3)&"<br>"
			end if
		end if
		timertmp=timer()
		if txt_errore="" then
			session("log_checkout")=session("log_checkout")&"<br>Inizio registrazione utente"
			'Aggiungo utente
			Set rs_utente = Server.CreateObject("ADODB.Recordset")
			rs_utente.Open "select utenti.* from utenti", conn, 1, 3
			rs_utente.addnew
			'Dati form
			rs_utente("email")=lcase(trim(request.form("email")))
			rs_utente("1email")=lcase(trim(request.form("email")))
			cognome=AllFirstUp(trim(request.form("cognome")))
			rs_utente("cognome")=cognome
			nome=AllFirstUp(trim(request.form("nome")))
			rs_utente("nome")=nome
			if Request.Form("tipo2")="P" then
				azienda=""
			else
				azienda=AllFirstUp(trim(request.form("azienda")))
			end if
			rs_utente("azienda")=azienda
			rs_utente("modalita_registrazione")=3
			if request.form("mailing")="Si" then rs_utente("mailing")=true else rs_utente("mailing")=false end if
			'Altri dati
			rs_utente("idbanca")=0	
			rs_utente("HTTP_USER_AGENT")=Request.ServerVariables("HTTP_USER_AGENT")
			rs_utente("data")=now()
			rs_utente.update
			iduser=Get_last_id("utenti")
			rs_utente.close
			conn.execute "UPDATE carrello SET carrello.iduser = "&iduser&", carrello.idloged=Null WHERE (((carrello.idloged)="&SessionIdGuest&"));",num
			logga "reg",iduser
			add2log "Registrazione [utente="&iduser&"]"&nominativo(nome,cognome,azienda)&"[/utente] in fase checkout, trasferiti "&num&" voci carrello",2
			'Tabella utenti_clienti
			rs_utente.Open "select utenti_clienti.* from utenti_clienti where iduser="&iduser, conn, 1, 3
			if rs_utente.eof then
				rs_utente.addnew
				rs_utente("iduser")=iduser
			end if
			rs_utente("tipo")=Request.Form("tipo2")
			rs_utente.Update 
			rs_utente.close
			set rs_utente=nothing
			invio_email_login=true
			session("log_checkout")=session("log_checkout")&"<br>Fine registrazione utente"
			timerstr=timerstr&"Registrazione utente: "&formatnumber(timer()-timertmp,3)&"<br>"
		end if
	end if
	timertmp=timer()
	if request.form("ordina")<>"" and txt_errore="" then
		
		if iduser=0 then iduser=sessionIDUser
		session("log_checkout")=session("log_checkout")&"<br>Inizio aggiornamento dati utente"
		'Aggiorno dati utente
		Set rs_utente = Server.CreateObject("ADODB.Recordset")
		rs_utente.Open "select * from utenti where iduser="&iduser, conn, 1, 3
		rs_utente("cognome")=AllFirstUp(trim(request.form("cognome")))
		rs_utente("nome")=AllFirstUp(trim(request.form("nome")))
		rs_utente("azienda")=AllFirstUp(trim(request.form("azienda")))
		rs_utente("indirizzo")=AllFirstUp(request.form("indirizzo"))
		rs_utente("citta")=AllFirstUp(request.form("citta"))
		rs_utente("cap")=request.form("cap")
		rs_utente("provincia")=ucase(request.form("provincia"))

		if request.form("regione")="" then
			set rs_regione=conn.execute ("select elenco_province.* from elenco_province where sigla_automobilistica='"&ucase(request.form("provincia"))&"'")
			if rs_regione.eof then 
			else
				rs_utente("regione")=rs_regione("codice_regione")
			end if
			set rs_regione=Nothing
		else
			rs_utente("regione")=cint(trim(request.form("regione")))
		end if

		if request.form("piva_editabile")="SI" then
			rs_utente("cf")=ucase(trim(request.form("cf")))
			rs_utente("piva")=ucase(trim(request.form("piva")))
		end if
		rs_utente("telefono")=trim(request.form("telefono"))
		rs_utente("cellulare")=request.form("cellulare")
		if request.form("consegna_diversa")="no" then
			rs_utente("d_azienda")=""
			rs_utente("d_indirizzo")=""
			rs_utente("d_citta")=""
			rs_utente("d_cap")=""
			rs_utente("d_provincia")=""
			rs_utente("d_regione")=0
		else
			rs_utente("d_azienda")=trim(request.form("d_azienda"))
			rs_utente("d_indirizzo")=AllFirstUp(request.form("d_indirizzo"))
			rs_utente("d_citta")=AllFirstUp(request.form("d_citta"))
			rs_utente("d_cap")=request.form("d_cap")
			rs_utente("d_provincia")=request.form("d_provincia")
			'rs_utente("d_regione")=cint(trim(request.form("d_regione")))
		end if
		rs_utente("tipopagamento")=request.form("tipopagamento")
		rs_utente("trasporto")=trim(request.form("trasporto"))
		'rs_utente("note")=trim(request.form("note-trasporto"))
		rs_utente.update
		rs_utente.close

		'Tabella utenti_clienti
		rs_utente.Open "select utenti_clienti.* from utenti_clienti where iduser="&iduser, conn, 1, 3
		if rs_utente.eof then
			rs_utente.addnew
			rs_utente("iduser")=iduser
		end if

		rs_utente("tipo")=Request.Form("tipo2")
		rs_utente.Update 
		rs_utente.close
		set rs_utente=nothing
		session("log_checkout")=session("log_checkout")&"<br>fine aggiornamento dati utente"
		session("log_checkout")=session("log_checkout")&"<br>Inizio creazione ordine"
		'add2log queryeform(),0
		timerstr=timerstr&"Trasferimento dati utente: "&formatnumber(timer()-timertmp,3)&"<br>"
		timertmp=timer()

		if request.form("ordina")="ordina" then
			preventivo=false
			'add2log "ordine iduser"&iduser,1
			set ordine= (new ClasseOrdine)(array("nuovo","ordini",iduser,iduser))
			idord=ordine.idord()
			set ordine = Nothing
		elseif request.form("ordina")="preventivo" then
			preventivo=true
			set ordine= (new ClasseOrdine)(array("nuovo","preventivi",iduser,iduser))
			idord=ordine.idord()
			set ordine = Nothing
		end if
		timerstr=timerstr&"ordine creato: "&formatnumber(timer()-timertmp,3)&" - "&time()&"<br>"
		timertmp=timer()
		ordinato=true
		session("log_checkout")=session("log_checkout")&"<br>Fine creazione ordine"
		if invio_email_login then
			res = invia_mail_registrazione(iduser,true)
			timerstr=timerstr&"Invio email registrazione: "&formatnumber(timer()-timertmp,3)&" - "&time()&"<br>"
			timertmp=timer()

		end if
		Session.Contents.remove("log_checkout")
		call add2log (timerstr,0)

	end if

end if
select2=true
set h = Nothing
%>
<!--#include virtual="/config/header_inc.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
        <section id="content">
        	<div id="breadcrumb-container">
        		<div class="container">
					<ul class="breadcrumb">
						<li><a href="/">Home</a></li>
						<%if preventivo then %>
						<li class="active"><%=traduci("tit15")%></li>
						<%else %>
						<li class="active"><%=traduci("tit10")%></li>
						<%end if%>
					</ul>
        		</div>
        	</div>
        	<div class="container">
        	<%if ordinato then %>
	        	<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<%if preventivo then%>
							<h1 class="title"><%=traduci("tit16")%></h1>
							<p class="title-desc"><%=traduci("tit16txt")%></p>
							<%else %>
							<h1 class="title"><%=traduci("tit14")%></h1>
							<p class="title-desc"><%=traduci("tit14txt")%></p>
							<%end if%>
						</header>
        			</div>
	        	</div>
        	<%else %>
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<%if preventivo then%>
							<h1 class="title"><%=traduci("tit15")%></h1>
							<p class="title-desc"><%=traduci("tit15txt")%></p>
							<%else %>
							<h1 class="title"><%=traduci("tit10")%></h1>
							<p class="title-desc"><%=traduci("tit10txt")%></p>
							<%end if%>
						</header>
						<%if txt_errore<>"" then%>
						<div class="alert alert-danger">
							<strong>Errore!</strong>&nbsp; <%=txt_errore%>
						</div>
						<%end if
						sql="select impostazioni.dati_trasporto FROM impostazioni"
						set rs=conn.execute(sql)
						dati_trasporto=split(rs("dati_trasporto"),vbcrlf)
						set rs = nothing
						piva_editabile="SI"
						pres_tipo="S"
						come_tipo="Default"
						if session("iduser")<>"" then
							pres_tipo=""
							sql="select * FROM utenti where "&where_carrello
							
							set rs_utente=conn.execute(sql)
							pres_nome=rs_utente("nome")
							pres_cognome=rs_utente("cognome")
							pres_azienda=rs_utente("azienda")
							pres_telefono=rs_utente("telefono")
							pres_cf=rs_utente("cf")
							pres_piva=rs_utente("piva")
							pres_indirizzo=rs_utente("indirizzo")
							pres_citta=rs_utente("citta")
							pres_cap=rs_utente("cap")
							pres_provincia=ucase(rs_utente("provincia"))
							pres_email=rs_utente("email")
							pres_pag_accordato=rs_utente("pag_accordato")
							
							'Recupero da Tabella utenti_clienti
							sql="select * FROM utenti_clienti where "&where_carrello
							
							set rs_utente_cliente=conn.execute(sql)
							if not rs_utente_cliente.eof then
								if rs_utente_cliente("tipo")<>"" and not isnull(rs_utente_cliente("tipo")) then
									pres_tipo=rs_utente_cliente("tipo")
									come_tipo="Da db"
								end if
							end if
							set rs_utente_cliente=nothing
							
							if pres_tipo="" then
								'Se non ancora impostato determino possibile tipo cliente
								if rs_utente("piva")="" then
									'Privato
									pres_tipo="P"
								elseif isnumeric(rs_utente("cf")) then
									'Società
									pres_tipo="S"
								else
									pres_tipo="D"
								end if
								come_tipo="Determinato"
							end if
							set rs_utente=nothing
							
							set rs_fatture=conn.execute ("select Count(fatture.IDFat) AS ConteggioDiIDFat	FROM fatture	WHERE "&where_carrello)
							if clng(rs_fatture("ConteggioDiIDFat"))=0 then
								piva_editabile="SI"
								class_disabled=""
							else
								piva_editabile="NO"
								class_disabled=" disabled"
							end if
							set rs_fatture=Nothing
						end if
						pres_tipo2=pres_tipo									
						%>

        				<form action="<%=questofile%>" method="post" id="checkout-form">
						<input type="hidden" name="piva_editabile" value="<%=piva_editabile%>">
						<input type="hidden" name="cheform" value="checkout">
						<%if preventivo then%>
        				<input type="hidden" name="ordina" value="preventivo">
						<%else %>
        				<input type="hidden" name="ordina" value="ordina">
						<%end if%>
        				<div class="panel-group custom-accordion" id="checkout">

							<div class="panel">
								<div class="accordion-header">
									<div class="accordion-title">1 Step: <span><%=traduci("step1")%></span></div><!-- End .accordion-title -->
									<a class="accordion-btn opened"  data-toggle="collapse" data-target="#checkout-option"></a>
								</div><!-- End .accordion-header -->
								
								<div id="checkout-option" class="collapse in">
								  <div class="panel-body">
								  <%if session("iduser")="" then%>
								   <div class="row">
							   	
							   	<div class="col-md-6 col-sm-6 col-xs-12">					   		
							   		<h2><%=dizionario.item("newcust")%></h2>
							   		
							   	<p><%=dizionario.item("newcusttxt")%></p>
								<div class="xs-margin"></div>
								<input type="hidden" name="nuovo_utente" value="si">
								
								<div class="nascondi">
								   	<div class="input-group" >
                                            <span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text">Email&#42;</span></span>
                                            <input type="text"  class="form-control input-lg" placeholder="<%=traduci("pl_email")%>" name="email" id="email">
                                        </div><!-- End .input-group -->
                                        <div class="input-group">
                                            <span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text"><%=traduci("repemail")%>&#42;</span></span>
                                            <input type="text"  class="form-control input-lg" placeholder="<%=traduci("repemailpl")%>" name="repeat_email" id="repeat_email">
                                        </div><!-- End .input-group -->
										</div>
									<div class="md-margin"></div>	
							   	</div><!-- End .col-md-6 -->
							   	<div class="col-md-6 col-sm-6 col-xs-12">					   		
							   		<h2><%=traduci("oldcust")%></h2>
							   		<p><%=traduci("oldcusttxt")%></p>
							   		<div class="xs-margin"></div>
							   		<%if err=1 then%>
							   		<div class="alert alert-danger">
                                    <strong>Errore!</strong>&nbsp; L'indirizzo email o la password sono errati.
                                </div>
							   		<%end if %>
							   	</div><!-- End .col-md-6 -->
						   </div><!-- End.row -->
								   
								   <%else %>
								   <div class="row">
								   	<div class="col-md-6 col-sm-6 col-xs-12">	
									   	<div class="input-group">
											<span class="input-group-addon">
											<span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("user")%></span></span>
											<span class="form-control input-lg" ><%=pres_email%></span>
										</div><!-- End .input-group -->
								   	</div><!-- End .col-md-6 -->
							   	
								   </div><!-- End.row -->
								   <%end if%>
								  </div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
							  
							  </div><!-- End .panel -->
							  <div class="panel">
								<div class="accordion-header">
									<div class="accordion-title">2 Step: <span><%=traduci("step2")%></span></div><!-- End .accordion-title -->
									<a class="accordion-btn opened"  data-toggle="collapse" data-target="#billing"></a>
								</div><!-- End .accordion-header -->
								
								<div id="billing" class="collapse in">
								  <div class="panel-body">
								  <fieldset>
								   <div class="row">
								   	<div class="col-md-6 col-sm-6 col-xs-12">
								   		
								   		<h2 class="checkout-title"><%=traduci("tit11")%></h2>
										<%if class_disabled<>"" then%>
										<%=traduci("nomodifica")%>
										<%end if%>
										<input type="hidden" name="tipo2" value="<%=pres_tipo2%>" id="tipo2">
										<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-subject"></span><span class="input-text">Tipo cliente&#42;</span></span>
											<div class="large-selectbox clearfix">
												<select id="tipo" name="tipo" >
													<option  value="S" <%if pres_tipo="S" then response.write "selected"%>>Società</option>
													<option  value="D" <%if pres_tipo="D" then response.write "selected"%>>Ditta individuale</option>
													<option  value="P" <%if pres_tipo="P" then response.write "selected"%>>Privato</option>
												</select>
											</div><!-- End .large-selectbox-->
										</div><!-- End .input-group -->									
										
								   		<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("nome")%>&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("nomepl")%>" name="nome" value="<%=pres_nome%>" id="nome">
									</div><!-- End .input-group -->
									
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("cognome")%>&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("cognomepl")%>" name="cognome" value="<%=pres_cognome%>" id="cognome">
									</div><!-- End .input-group -->
									
									 
                                    <div class="input-group " id="divazienda" <%if pres_tipo="P" or pres_tipo="D" then %>style="display: none;"<%end if%>>
										<span class="input-group-addon"><span class="input-icon input-icon-company"></span><span class="input-text"><%=traduci("azienda")%><span id="azrequired" <%if pres_tipo="D" then %>style="display: none;"<%end if%>>&#42;</span></span></span>
										<input type="text" class="form-control input-lg" placeholder="<%=traduci("aziendapl")%>" name="azienda" id="azienda" value="<%=pres_azienda%>">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-phone"></span><span class="input-text"><%=traduci("telefo")%>&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("telefopl")%>" name="telefono" value="<%=pres_telefono%>" id="telefono">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-subject"></span><span class="input-text"><%=traduci("cf")%>&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("cfplnum")%>" name="cf" value="<%=pres_cf%>" id="cf" <%=class_disabled%>>
									</div><!-- End .input-group -->
									 
									<div class="input-group" id="divpiva" <%if pres_tipo="P" then %>style="display: none;" <%end if%>>
										<span class="input-group-addon"><span class="input-icon input-icon-subject"></span><span class="input-text"><%=traduci("piva")%>&#42;</span></span>
										<input type="text" class="form-control input-lg" placeholder="<%=traduci("pivapl")%>" name="piva" value="<%=pres_piva%>" id="piva" <%=class_disabled%>>
									</div><!-- End .input-group -->
									
								   	</div><!-- End .col-md-6 -->
								   	
								   	<div class="col-md-6 col-sm-6 col-xs-12">
										<h2 class="checkout-title"><%=traduci("tit12")%></h2>
	
										<div class="input-group">
											<span class="input-group-addon"><span class="input-icon input-icon-address"></span><span class="input-text"><%=traduci("indir")%>&#42;</span></span>
											<input type="text" class="form-control input-lg" placeholder="<%=traduci("indirpl")%>" name="indirizzo" value="<%=pres_indirizzo%>" id="indirizzo">
										</div><!-- End .input-group -->
										
										<div class="input-group select2-bootstrap-prepend">
											<span class="input-group-addon"><span class="input-icon input-icon-city"></span>
											<span class="input-text"><%=traduci("citta")%>&#42;</span></span>
											<input type="text"  class="form-control input-lg" name="citta" id="citta" placeholder="<%=traduci("cittapl")%>" value="<%=pres_citta%>" ><input type="hidden" id="pres_citta" value="<%=pres_citta%>">
										</div><!-- End .input-group -->
										
										<div class="input-group">
											<span class="input-group-addon"><span class="input-icon input-icon-postcode"></span><span class="input-text"><%=traduci("cap")%>&#42;</span></span>
											<input type="text"  class="form-control input-lg" name="cap" id="cap" placeholder="<%=traduci("cappl")%>" value="<%=pres_cap%>">
										</div><!-- End .input-group -->
	
	
	
										<div class="input-group lg-margin select2-bootstrap-prepend">
											<span class="input-group-addon"><span class="input-icon input-icon-region"></span><span class="input-text"><%=traduci("prov")%>&#42;</span></span>
											
											<select name="provincia" id="provincia" class="select2 form-control" >
												<%   
												if pres_provincia="" then %>
												<option  value="" ></option>
												<%end if
												Set rs_provincia = conn.execute("select * from elenco_province order by Denominazione_provincia")
												do while not rs_provincia.eof 
												if ucase(rs_provincia("Sigla_automobilistica"))=pres_provincia then txt_selected="selectED" else txt_selected="" end if
												%>
                                                <option  value="<%=rs_provincia("Sigla_automobilistica")%>" <%=txt_selected%>><%=rs_provincia("Denominazione_provincia")%></option>
                                                <%
                                                rs_provincia.movenext
                                                loop
                                                Set rs_provincia = Nothing
                                                %>
                                            </select>
										</div><!-- End .input-group -->
										
										
										
										<%if false then%>
	                                    <div class="input-group">
											<span class="input-group-addon"><span class="input-icon input-icon-country"></span><span class="input-text"><%=traduci("country")%>&#42;</span></span>
											
										</div><!-- End .input-group -->
										<%end if%>
								   	</div><!-- End .col-md-6 -->
								   </div><!-- End .row -->
								   <div class="col-md-12">
								   	<a href="javascript: void(0)" class="btn btn-custom-2" id="btnstep2" ><%=traduci("continue")%></a>
								   </div>
								  </fieldset>
								  </div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
							  
							  </div><!-- End .panel -->
							  
							  <div class="panel"><!-- .panel 3 step--><!-- .panel 3 step--><!-- .panel 3 step--><!-- .panel 3 step--><!-- .panel 3 step-->
								<div class="accordion-header">
									<div class="accordion-title">3 Step: <span><%=traduci("step3")%></span></div><!-- End .accordion-title -->
									<a class="accordion-btn"   data-toggle="collapse" data-target="#delivery-details"></a>
								</div><!-- End .accordion-header -->
								
								<div id="delivery-details" class="collapse">
								  <div class="panel-body">
  								   <div class="row">
									   	<div class="col-md-6 col-sm-6 col-xs-12">
		   								   	<div class="input-group custom-checkbox sm-margin">
													 <input type="checkbox" checked="cheked" id="consegna_diversa" name="consegna_diversa" value="no"> <span class="checbox-container">
													 	<i class="fa fa-check"></i>
													 </span>
													 <%=traduci("txt11")%>
											</div><!-- End .input-group -->
									   	</div><!-- End .col-md-6 -->
								   </div><!-- End .row -->
									<div id="consegna" style="display:none;">
	  								   <div class="row">
										   	<div class="col-md-6 col-sm-6 col-xs-12">
			                                    <h2 class="checkout-title"><%=traduci("tit13")%></h2>
			                                    <div class="input-group">
													<span class="input-group-addon"><span class="input-icon input-icon-company"></span><span class="input-text"><%=traduci("azienda")%></span></span>
													<input type="text" class="form-control input-lg" placeholder="<%=traduci("aziendapl")%>" name="d_azienda" >
												</div><!-- End .input-group -->
			                                    <div class="input-group">
													<span class="input-group-addon"><span class="input-icon input-icon-address"></span><span class="input-text"><%=traduci("indir")%>&#42;</span></span>
													<input type="text" class="form-control input-lg" placeholder="<%=traduci("indirpl")%>" name="d_indirizzo" id="d_indirizzo">
												</div><!-- End .input-group -->
												<div class="input-group">
													<span class="input-group-addon"><span class="input-icon input-icon-city"></span><span class="input-text"><%=traduci("citta")%>&#42;</span></span>
													<input type="text"  class="form-control input-lg" placeholder="<%=traduci("cittapl")%>" name="d_citta" id="d_citta">
												</div><!-- End .input-group -->
												<div class="input-group">
													<span class="input-group-addon"><span class="input-icon input-icon-postcode"></span><span class="input-text"><%=traduci("cap")%>&#42;</span></span>
													<input type="text"  class="form-control input-lg" placeholder="<%=traduci("cappl")%>" name="d_cap" id="d_cap">
												</div><!-- End .input-group -->
		
										   	</div><!-- End .col-md-6 -->
									   </div><!-- End .row -->								   	
									</div><!-- End #consegna -->
									<div class="col-md-12">
								   	<a href="javascript: void(0)" class="btn btn-custom-2" id="btnstep3" ><%=traduci("continue")%></a>
								   </div>
								  </div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
							  
							  </div><!-- End .panel -->
							  
							 <%
								 'Genero elenco articoli perchè mi serve il totale
								 
								tot_carrello=0
								elenco_articoli=""
    							sql="select carrello.*, prodotti.codice, prodotti.articolo, prodotti.variante1, prodotti.variante2, prodotti.prezzo, prodotti.costo, prodotti.Promozione, prodotti.Sconto, prodotti.Prodata, varianti_a.variante_a, varianti_a.prezzo_ve_va, varianti_b.variante_b, prodotti.idfor, prodotti.fileimg, prodotti.aggiungi_a_ordine FROM varianti_b RIGHT JOIN (varianti_a RIGHT JOIN (carrello INNER JOIN prodotti ON carrello.idpro = prodotti.IDpro) ON varianti_a.IDvara = carrello.idvara) ON varianti_b.IDvarb = carrello.idvarb where "&where_carrello
								Set rs_carrello=conn.Execute(sql)
								do while not rs_carrello.EOF
								
								if converti_typevar14(rs_carrello("prezzo_ve_va"))>0 then prezzo=rs_carrello("prezzo_ve_va") else prezzo=rs_carrello("prezzo")
									prezzo=cdbl(prezzo)
								'prezzo=rs_carrello("prezzo")
								if rs_carrello("promozione")=1 and rs_carrello("sconto")>0 then
									prezzoeff=prezzo*(1-rs_carrello("sconto")/100)
									txtprezzo="<s>"&formatnumber(prezzo,2)&"</s><br><b>" & formatnumber(prezzoeff,2)&"</b>"
								else
									prezzoeff=prezzo
									txtprezzo=formatnumber(prezzoeff,2)
								end if
								
								subtot=prezzoeff*rs_carrello("quantita")
								
								elenco_articoli=elenco_articoli&"<tr><td class=""item-name-col"">"
								if false then %>
											<figure>
												<a href="#"><img src="images/products/compare-placeholder.jpg" alt="Lowlands Lace Blouse"></a>
											</figure>
											<%end if	
								elenco_articoli=elenco_articoli&"<header class=""item-name""><a href=""product.asp?idpro="&rs_carrello("idpro")&""">"&rs_carrello("articolo")&"</a></header>"
										

											
											'varianti 
											txt=""
											if rs_carrello("idvara")<>""  and rs_carrello("idvara")<>0 then
												txt ="<li>" & rs_carrello("variante1") & ": " &  rs_carrello("variante_a")&"</li>"
											end if
											if rs_carrello("idvarb")<>"" and rs_carrello("idvarb")<>0 then
												txt =txt & "<li>" & rs_carrello("variante2") & ": "  &  rs_carrello("variante_b")&"</li>"
											end if
											if txt<>"" then
												elenco_articoli=elenco_articoli&"<ul>"&txt&"</ul>"
											end if
										elenco_articoli=elenco_articoli&"</td><td class=""item-code"">"&rs_carrello("codice")&"</td><td class=""item-price-col""><span class=""item-price-special"">"&simbolo_valuta&FormatNumber(prezzoeff,2)&"</span></td><td class=""item-price-col""><span class=""item-price-special"">"&rs_carrello("quantita")&"</span></td><td class=""item-total-col""><span class=""item-price-special"">"&simbolo_valuta&FormatNumber(subtot,2)&"</span></td></tr>"
										
									idpro=rs_carrello("idpro")
									tot_carrello=tot_carrello+subtot
									rs_carrello.MoveNext
									loop
									set rs_carrello = Nothing
									subtot_carrello=tot_carrello
								 
								 %> 
							  
							  
							  
							  <div class="panel"><!-- .panel 4 step--><!-- .panel 4 step--><!-- .panel 4 step--><!-- .panel 4 step--><!-- .panel 4 step-->
								<div class="accordion-header">
									<div class="accordion-title">4 Step: <span><%=traduci("step4")%></span></div><!-- End .accordion-title -->
									<a class="accordion-btn"  data-toggle="collapse" data-target="#delivery-method"></a>
								</div><!-- End .accordion-header -->
								
								<div id="delivery-method" class="collapse">
								  <div class="panel-body">
									  
									  <div class="row">
										   	<div class="col-md-6 col-sm-6 col-xs-12">
										   	<p><%=traduci("txt12")%></p>
										   	<div class="input-group col-md-12 col-sm-12 col-xs-12">
											  <div class="radio">
			                                  <label>
			                                    <input type="radio" name="trasporto" id="Trasporto_1" value="1" class="radio_trasporto radio_change">
			                                    <%=tipo_trasporto(1)%>
			                                  </label>
			                                  <span class="pull-right"><%=simbolo_valuta&formatnumber(calcola_importo_trasporto(1,subtot_carrello),2)%></span>
			                                </div>
			                                <div class="radio">
			                                  <label>
			                                    <input type="radio" name="trasporto" id="Trasporto_2" value="2" class="radio_trasporto radio_change">
			                                    <%=tipo_trasporto(2)%>
			                                  </label>
			                                  <span class="pull-right"><%=simbolo_valuta&formatnumber(calcola_importo_trasporto(2,subtot_carrello),2)%></span>

			                                  
			                                </div>
			                                <div class="radio">
			                                  <label>
			                                    <input type="radio" name="trasporto" id="Trasporto_3" value="3" class="radio_trasporto radio_change">
			                                    <%=tipo_trasporto(3)%>
			                                  </label>
			                                  <span class="pull-right"><%=simbolo_valuta&formatnumber(calcola_importo_trasporto(3,subtot_carrello),2)%></span>
			                                </div>
			                                <div class="radio">
			                                  <label>
			                                    <input type="radio" name="trasporto" id="Trasporto_4" value="4" class="radio_trasporto radio_change">
			                                    <%=tipo_trasporto(4)%>
			                                  </label>
			                                  <span class="pull-right"><%=simbolo_valuta&formatnumber(calcola_importo_trasporto(4,subtot_carrello),2)%></span>
			                                </div>
										   	</div>
									   	</div><!-- End .col-md-12 -->
									   	<div class="col-md-6 col-sm-6 col-xs-12">
										   	
										   	<div class="input-group textarea-container">
			                                    <span class="input-group-addon"><span class="input-icon input-icon-message"></span><span class="input-text">Note</span></span>
			                                    <textarea name="note-trasporto" class="form-control" cols="30" rows="6" placeholder="<%=traduci("notepl")%>" style="resize: none;"></textarea>
			                                </div>
									   	</div>
									  </div><!-- End .row -->	
								
							   	<div class="col-md-12">
								   	<a href="javascript: void(0)" class="btn btn-custom-2" id="btnstep4" ><%=traduci("continue")%></a>
								   </div>
								  </div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
							  
							  </div><!-- End .panel -->
        				
							<div class="panel"><!-- .panel 5 step--><!-- .panel 5 step--><!-- .panel 5 step--><!-- .panel 5 step--><!-- .panel 5 step-->
								<div class="accordion-header">
									<div class="accordion-title">5 Step: <span><%=traduci("step5")%></span></div><!-- End .accordion-title -->
									<a class="accordion-btn"  data-toggle="collapse" data-target="#payment-method"></a>
								</div><!-- End .accordion-header -->
								
								<div id="payment-method" class="collapse">
								  <div class="panel-body">
									  <div class="row">
										   	<div class="col-md-6 col-sm-6 col-xs-12">
									  <p><%=traduci("txt13")%></p>
			                                    <div class="input-group col-md-12 col-sm-12 col-xs-12">
								  <%
								if isnull(pres_pag_accordato) then pres_pag_accordato=""
								
								response.write metodo_pagamento(-2,pres_pag_accordato)
									%>
			                                    </div><!-- End .input-group -->
								   	</div><!-- End .col-md-12 -->
										   	<div class="col-md-6 col-sm-6 col-xs-12">
										   	</div>
								  </div><!-- End .row -->	

									<div class="col-md-12">
										<a href="javascript: void(0)" class="btn btn-custom-2" id="btnstep5" ><%=traduci("continue")%></a>
									</div>
								  </div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
							  
						  	</div><!-- End .panel -->
        					<div class="panel">
								<div class="accordion-header">
									<div class="accordion-title">6 Step: <span><%=traduci("step6")%></span></div><!-- End .accordion-title -->
									<a class="accordion-btn opened"  data-toggle="collapse" data-target="#confirm"></a>
								</div><!-- End .accordion-header -->
								
								<div id="confirm" class="collapse in">
								  <div class="panel-body">
								  
								  
							  
								<div class="table-responsive">									
									<table class="table checkout-table" >
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
								<%=elenco_articoli%>
									</tbody>
									 <tfoot>
										<tr>
											<td class="checkout-table-title" colspan="4"><%=traduci("subtot")%>:</td>
											<td class="checkout-table-price"><%=simbolo_valuta&FormatNumber(tot_carrello,2)%></td>
										</tr>
										<%
											trasporto=roundup(calcola_importo_trasporto(tipotrasporto,tot_carrello),2)
											tot_carrello=tot_carrello+trasporto
										%>
										<tr>
											<td class="checkout-table-title" id="cella_testo_trasporto" colspan="4"><%=traduci("ship")%>:</td>
											<td class="checkout-table-price" id="cella_valore_trasporto"><%=simbolo_valuta&FormatNumber(trasporto)%></td>
										</tr>
										<%
											valore_iva=tot_carrello*iva(date())/100
											tot_carrello=tot_carrello+valore_iva
										%>
										<tr>
											<td class="checkout-table-title" colspan="4"><%=traduci("tax")%>&nbsp;&nbsp;<%=iva(date())%>%:</td>
											<td class="checkout-table-price" id="cella_valore_iva"><%=simbolo_valuta&FormatNumber(valore_iva,2)%></td>
										</tr>
										<%
											tot_carrello=round(tot_carrello,2)										
										%>
										<tr>
											<td class="checkout-total-title" colspan="4"><strong><%=traduci("tot")%>:</strong></td>
											<td class="checkout-total-price cart-total" id="cella_totale"><strong><%=simbolo_valuta&FormatNumber(tot_carrello,2)%></strong></td>
										</tr>
									</tfoot>
								  </table>
								
								</div><!-- End .table-reponsive -->
								<div class="lg-margin"></div><!-- space -->
								<%if preventivo then%>

										   	
								   	<div class="input-group textarea-container">
	                                    <span class="input-group-addon"><span class="input-icon input-icon-message"></span><span class="input-text"><%=traduci("notep")%></span></span>
	                                    <textarea name="note-su-ordine" class="form-control" cols="30" rows="6" placeholder="<%=traduci("noteppl")%>" style="resize: none;"></textarea>
	                                </div>
							   	<%else %>
										   	
								   	<div class="input-group textarea-container">
	                                    <span class="input-group-addon"><span class="input-icon input-icon-message"></span><span class="input-text"><%=traduci("noteo")%></span></span>
	                                    <textarea name="note-su-ordine" class="form-control" cols="30" rows="6" placeholder="<%=traduci("noteopl")%>" style="resize: none;"></textarea>
	                                </div>
									
                                   
   							   	<%end if %>						
								
								
									</div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
							  
						  	</div><!-- End .panel -->
							
							
							
							<div class="panel"><!-- .panel 5 step--><!-- .panel 5 step--><!-- .panel 5 step--><!-- .panel 5 step--><!-- .panel 5 step-->
								<div class="accordion-header">
									<div class="accordion-title"> <span>CONDIZIONI DI VENDITA</span></div><!-- End .accordion-title -->
								</div><!-- End .accordion-header -->
								
								<div id="panel_condizioni" class="collapse in">
								  <div class="panel-body">
								 <fieldset class="half-margin">
										<div class="input-desc-box">
										<span class="input-group custom-checkbox">
									 <input type="checkbox" name="condizioni"> <span class="checbox-container">
									 <i class="fa fa-check"></i>
									 </span>Dichiaro di aver letto e di accettare senza riserve le <a href="condizioni.asp">condizioni di vendita e di garanzia</a></span>
										</div><!-- End .input-desc -->
									</fieldset>
								  </div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
						  	</div><!-- End .panel -->
							
							
							
							<%if subtot_carrello>0 then%>
									<div class="">
									<%if preventivo then%>
									<input type="submit" class="btn btn-custom-2" id="conferma_ordine" value="<%=traduci("step7")%>">
									<%else%>
									
									<input type="submit" class="btn btn-custom-2" id="conferma_ordine" value="<%=traduci("step6")%>">
									<%end if%>
									</div>
								<%end if%>
        				</div><!-- End .panel-group #checkout -->
        				</form>
        				<div class="xlg-margin"></div><!-- space -->
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
        		<%end if%>
			</div><!-- End .container -->
        
        </section><!-- End #content -->
        
        
    <%
	   'Memorizzo traduzioni prima di chiudere il dizionario 
	    cfplnum=traduci("cfplnum")
	    cfplalfa=traduci("cfplalfa")
	    'Memorizzo lunghezza campi prima di chiudere la connessione
	    Set rs_utente = Server.CreateObject("ADODB.Recordset")
		rs_utente.Open "select * from utenti", conn
	    email_DefinedSize=rs_utente("email").DefinedSize
	    nome_DefinedSize=rs_utente("nome").DefinedSize
	    cognome_DefinedSize=rs_utente("cognome").DefinedSize
	    telefono_DefinedSize=rs_utente("telefono").DefinedSize
	    azienda_DefinedSize=rs_utente("azienda").DefinedSize
	    citta_DefinedSize=rs_utente("citta").DefinedSize
		rs_utente.close
		set rs_utente=nothing
	    
	 %>    
<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
    <script src="js/jquery.selectbox.min.js"></script>
	<!--#include virtual="/script_inc.asp" -->
    <%if not ordinato then %>

    <!-- AGGIUNTE -->
    <script src="jquery/js/jquery.scrollintoview.min.js"></script>
	<script src="js/Jqueryvalidate/jquery.validate.min.js"></script>
	<script src="js/Jqueryvalidate/localization/messages_it.min.js"></script>
	<script type="text/javascript" src="chk_piva_cf.js"></script>
	
	<script src="Jquery/js/select2/select2.min.js"></script>	        
	<script src="Jquery/js/select2/select2_locale_it.js"></script>
    <script src="js/jquery.selectbox.min.js"></script>
	<script>
	$(function() {
		var tipo="<%=pres_tipo%>";
		var piva_editabile="<%=piva_editabile%>";
		var txtErrori;
		var come_tipo="<%=come_tipo%>";
	
	<%
	if session("iduser")<>"" then
		testo="iduser: "&sessionIDUser
	else
		testo="idloged: "&SessionIdGuest
	end if
	%>

	$("#tipo").selectbox({
		onChange: function (val, inst) {
				console.log("tipo onChange:"+val);
				tipo=val;
				$("#tipo2").val(val);
				switch(val) {
				    case "S":
				    $("#divpiva").show();
				    $("#azrequired").show();
					
				    $("#divazienda").show();
					$("#cf").attr("placeholder", "<%=cfplnum%>");
				        break;
				    case "D":
				    $("#divpiva").show();
				    $("#azrequired").hide();
				    $("#divazienda").show();
					//Nascondo messaggi di errore
					$("#cf").attr("placeholder", "<%=cfplalfa%>");
				        break;
				    case "P":
					$("#cf").attr("placeholder", "<%=cfplalfa%>");
				    $("#divpiva").hide();
				    $("#divazienda").hide();
				    $("#piva").val("");
					//Nascondo messaggi di errore
					$("#piva-error").hide();
					$("#azienda-error").hide();
				        break;
				}
				
		}
	});

	<%
	if piva_editabile="NO" and not (  isnull(pres_tipo2) or pres_tipo2="") then %>
	$("#tipo").selectbox("disable");
	<%end if %>





		$(".radio_change").on("change",function(){
			var click_su=$(this).attr("name");
			console.log(click_su);
			//var tipo_trasporto=$(this).val();
			//var tipo_pagamento=$(".tipopagamento").val();
			var tipo_trasporto=$('.radio_trasporto:checked').val();
			var tipo_pagamento=$('.tipopagamento:checked').val();
			console.log("click su:"+click_su+"trasporto:"+tipo_trasporto+" tipopagamento:"+tipo_pagamento);			
			$.ajax({
			  dataType: "json",
			  url: "searcher.asp",
			  type: "POST",
			  cache: false,
			  data: {oper: "checkout",  <%=testo%>,tipo_trasporto: tipo_trasporto,subtot_carrello:<%=replace(subtot_carrello,",",".")%>,tipo_pagamento: tipo_pagamento, click_su: click_su},
			  success: function (data) {
	            //console.log(JSON.stringify(data));
	            $("#cella_valore_trasporto").html(data.trasporto);
	            $("#cella_valore_iva").html(data.iva);
	            $("#cella_totale").html(data.tot_carrello);
	            $("#cella_testo_trasporto").html(data.testo);
	            },
	            error:function(xhr, textStatus, error){
			      console.log("xhr.statusText:"+xhr.statusText);
			      console.log("xhr.responseText:"+xhr.responseText);
			      console.log("textStatus:"+textStatus);
			      console.log("error:"+error);
				  }
  			});
		});
		$("#consegna_diversa").click(function(){
			var show=!this.checked
			if(show){
				//$("#consegna").removeClass("hidden");
				$("#consegna").slideDown();
			}
			else{
				$("#consegna").slideUp();
			}
		});	
		$("#btnstep2").click(function(e){
			
			var txtErrori="";
			<%if session("iduser")="" then %>
			
			validator.element( "#email" );
				for (var i in validator.errorMap) {
				  console.log(i, ":", validator.errorMap[i]);
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
				
			validator.element( "#repeat_email" );
				for (var i in validator.errorMap) {
				  console.log(i, ":", validator.errorMap[i]);
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			<%end if %>
			validator.element( "#nome" );
				for (var i in validator.errorMap) {
				  console.log(i, ":", validator.errorMap[i]);
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}

			validator.element( "#cognome" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			validator.element( "#azienda" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			validator.element( "#telefono" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			validator.element( "#cf" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}

			validator.element( "#piva" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			validator.element( "#indirizzo" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			validator.element( "#citta" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			validator.element( "#cap" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			validator.element( "#provincia" );
				for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
				
			if(txtErrori==''){
				$("#delivery-details").collapse('show');
				e.preventDefault();
				$("#btnstep3").scrollintoview();
			}
			else
			{
				console.log ("txtErrori:"+txtErrori);
		           invia_dati_checkout(txtErrori,validator.numberOfInvalids());
			}
		})
		$("#btnstep3").click(function(e){
			validator.element( "#d_indirizzo" );
			validator.element( "#d_cap" );
			validator.element( "#d_citta" );
			if(validator.numberOfInvalids()==0){
				$("#delivery-method").collapse('show');
				e.preventDefault();
				$("#btnstep4").scrollintoview();
			}
		})		
		$("#btnstep4").click(function(e){
			var txtErrori="";
			console.log("Valido trasporto");
			validator.element( ".radio_trasporto" );
			
			for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			
			if(txtErrori==""){
			
				$("#payment-method").collapse('show');
				e.preventDefault();
				$("#btnstep5").scrollintoview();
			}else
			{
			console.log("trasporto non valido numberOfInvalids:"+validator.numberOfInvalids()+"  errori:"+txtErrori+" validator.errorMap:"+validator.errorMap);
			}
		})		
		$("#btnstep5").click(function(e) {
			var txtErrori="";

			validator.element( ".tipopagamento" );
			for (var i in validator.errorMap) {
				  txtErrori+=i+":"+validator.errorMap[i]+"<br>";
				}
			
			if(txtErrori==""){
				e.preventDefault();
				$("#conferma_ordine").scrollintoview({ duration: "normal" });
			}
			
		});
		
		$.validator.addMethod("codice_fiscale", function (value, element) {
			result = true;
			var tipo=$("#tipo2").val();
			if (piva_editabile=="SI"){
				if (tipo=="S")
				{
					chk=ControllaCFNumerico(value);
					txtErrori+="ControlloCF: ControllaCFNumerico:"+value;
				}
				else
				{
					chk=ControllaCF(value);
					txtErrori+="ControlloCF: ControllaCF:"+value;
				}
				if (chk != '' ) { 
					$.validator.messages.codice_fiscale = chk;
					result = false; 
				}
				if (value=="")
				{
					$.validator.messages.codice_fiscale = "Inserire il codice fiscale";
					result = false; 
				}
			}
			return result;
		}, "");
		$.validator.addMethod("partita_iva", function (value, element) {
			result = true;
			var tipo=$("#tipo2").val();
			if (tipo!="P"){
				chk=ControllaPIVA(value);
				if (chk != '' ) { 
					$.validator.messages.partita_iva = chk;
					result = false; 
				}
				if ((tipo!="P") && (value=="")){
					$.validator.messages.partita_iva = "Inserire la partita iva, se sei un privato seleziona privato nel campo tipo cliente";
					result = false; 
				}
			}
			return result;
		}, "");

		var validator=$("#checkout-form").validate({
			rules: {
				email: {required: {
			        depends: function(element){
				            //return $("#registrati").is(':checked');
							return true;
				        }},
						email: true,maxlength:<%=email_DefinedSize%>,
				        remote:{
							url: "searcher.asp?check_email=si", //make sure to return true or false with a 200 status code
							type: "post",
							dataType: 'json',
							cache:false,
							dataFilter: function(data) {
								var json = JSON.parse(data);
								if(json.isError == "true") {
									return "\"" + decodeURIComponent(escape(json.errorMessage)) + "\"";
								} else {
									return true;
								}
							},
							error:function(xhr, textStatus, error,data){
									txt="Chiamata AJAX validazione email<br>";
								  txt+="<br>xhr.statusText:"+xhr.statusText;
								  txt+="<br>xhr.responseText:"+xhr.responseText;
								  txt+="<br>textStatus:"+textStatus;
								  txt+="<br>error:"+error;
								  txt+="<br>data:"+data;
								    $.ajax({
									cache: false,
									url     : "searcher.asp?titolo=Validazione%20form%20non%20riuscita",
									type    : "post",
									data	: "txt_errore="+encodeURIComponent(txt)
									
									});
							}
						}},
						repeat_email: {required: {
							depends: function(element){
									return true;
								}
								},
								email: true,
								equalTo:{
									param:"#email",
									depends: function(element){
									//return $("#registrati").is(':checked');
									return true;
								}
				        }},

				//Step 2
				
				
				nome: {required:true,minlength:3,maxlength:<%=nome_DefinedSize%>},
				cognome: {required: true,minlength: 3,maxlength:<%=cognome_DefinedSize%>},
				azienda: {required: {
							depends: function(element){
								var tipo=$("#tipo2").val();
								if (tipo=="S"){
										return true;
									}
									else
									{
										return false;
									}
								
				            
							//return true;
				        }},maxlength:<%=azienda_DefinedSize%>},
				telefono:{required:true,digits: true,minlength: 8,maxlength:<%=telefono_DefinedSize%>},
				cf: { codice_fiscale: true },
				piva: { partita_iva: true },
				indirizzo: {required: true,minlength: 5},
				citta: {required: true,minlength: 3,maxlength:<%=citta_DefinedSize%>},
				cap: {required: true,number: true,minlength: 5, maxlength: 5},
				provincia: { required: true },
				//Step 3
				d_indirizzo: {required: {
			        depends: function(element){
			        	//console.log("consegna_diversa"+$("#consegna_diversa").is(':checked'));
				            return !$("#consegna_diversa").is(':checked');
				        }}},
				d_cap: {required: {
			        depends: function(element){
			        	//console.log("consegna_diversa"+$("#consegna_diversa").is(':checked'));
				            return !$("#consegna_diversa").is(':checked');
				        }}, number: true,minlength: 5, maxlength: 5 },
				d_citta: {required: {
			        depends: function(element){
			        	//console.log("consegna_diversa"+$("#consegna_diversa").is(':checked'));
				            return !$("#consegna_diversa").is(':checked');
				        }}, minlength: 3},
				//Step 4
				trasporto: {required:true},
				//Step 5
				tipopagamento: {required:true},
				condizioni: {required:true}
			},
			ignore: "",
			errorElement: "span",
		    errorClass: "help-block",
		    highlight: function (element, errorClass, validClass) {
		        $(element).closest('.input-group').addClass('has-error');
		    },
		    unhighlight: function (element, errorClass, validClass) {
		        $(element).closest('.input-group').removeClass('has-error');
		    },
		    errorPlacement: function (error, element) {
		        if (element.parent('.input-group').length)
		         {
		            error.insertAfter(element.parent());
		            console.log("qui1:"+element)
		          }
		          else if ( element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
			            error.insertAfter(element.closest('.input-group'));
	
		            console.log("qui2:"+element)
		        } else {
	
		           // error.insertAfter(element);
		            console.log("qui3:"+element)
	
		        }
		    },
		    
		    invalidHandler: function(e, validator){
	           if(validator.errorList.length){
		           var txt_errori='';
		           //Apro tutti i panel con errori
		           for (var i=0;i<validator.errorList.length;i++){
				        console.log("errore in:"+$(validator.errorList[i].element).attr("id"));
				        txt_errori+=$(validator.errorList[i].element).attr("name")+',';
						console.log("pannello contenitore:"+$(validator.errorList[i].element).closest( ".collapse" ).attr("id"));
						if (!$(validator.errorList[i].element).closest( ".collapse" ).hasClass('in')){
				        $(validator.errorList[i].element).closest( ".collapse" ).collapse("show");}
						
				    }
		           $(validator.errorList[0].element).closest( ".collapse" ).scrollintoview();
		           invia_dati_checkout(txt_errori,validator.numberOfInvalids());
				}
	        },
	        submitHandler: function(form){
	            //$('#form_valido').val('si');
	            $("#conferma_ordine").prop("value","Attendi");
	            //$("#conferma_ordine").prop("disabled",true);
				form.submit();
			}
		});
			
		$('.select2').select2({
	  		minimumInputLength: 1,
	  		allowClear: true
	  	});
		$('#citta').select2({
	  		minimumInputLength: 1,
	  		//allowClear: true,
	  		ajax: {
	  			quietMillis: 150,
	  			url: "searcher_comuni.asp?plugin=select2",
	  			dataType: 'json',
	  			data: function (term, page) {
	  				return {
	  					term: term
	  				};
	  			},
	  			results: function (data) {
	  				return {
	  					results: data
	  				};
	  			}
	  		},
			//Allow manually entered text in drop down.
			createSearchChoice:function(term, data) {
				if ( $(data).filter( function() {
				  return this.text.localeCompare(term)===0;
				}).length===0) {
				  return {id:term, text:term};
				}
			},
			initSelection: function (element, callback) {
				callback({ id: $("#pres_citta").val(), text: $("#pres_citta").val() });
			},
			id: function(object) {
				console.log("id function");
				return object.text;
			}
	  	});
	  	$("#citta").on('select2-selecting', function(e) {
			;$("#citta").val(e.object.text);
			$("#cap").val(e.object.cap);
			$("#provincia").val(e.object.id).trigger("change");
			$("#regione").val(e.object.regione);
		});
		function invia_dati_checkout(txt_errori,numerrori){
			
			console.log("invia_dati_checkout");
						           var txt="";
					  txt+="Validazione checkout fallita invalidHandler<br>";
					  txt+="Numero campi non validi: "+numerrori+"<br>";
					  txt+="<b>"+txt_errori+"</b><br>";
					  txt+="Nome:"+$("#nome").val()+"<br>";
					  txt+="Cognome:"+$("#cognome").val()+"<br>";
					  txt+="Tipo field: "+$("#tipo2").val()+"<br>";
					  txt+="Tipo variabile: "+tipo+"<br>";
					  txt+="CF:"+$("#cf").val()+"<br>";
					  txt+="Lunghezza cf:"+$("#cf").val().length+"<br>";
					  txt+="piva_editabile:"+piva_editabile;
					  txt+="piva:"+$("#piva").val()+"<br>";
					  txt+="Campo piva visibile: "+$("#piva").is(":visible")+"<br>";
					  txt+="Campo azienda visibile: "+$("#azienda").is(":visible")+"<br>";
					  var campiform;
					  var n;
					  campiform=$("#checkout-form").serializeArray();
					  n=campiform.length;
					  txt+="<b>Campi form:</b><br>";
					  for (var i = 0; i < n; i++) {
						  txt+=campiform[i].name+"="+campiform[i].value+"<br>";
						}
					  $.ajax({
						cache: false,
						url     : "searcher.asp?titolo=Validazione%20form%20non%20riuscita&grave=1",
						type    : "post",
						data	: "txt_errore="+encodeURIComponent(txt)
						
						});

			
			
		}
	});
	</script>
	<%end if%>

    </body>
</html>