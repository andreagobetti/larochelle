<!--#include virtual="/mail_registrazione_inc.asp" -->
<!--#include file="ClasseUtente.asp"-->

<%
		dim h
		set h = new MD5
%>
<!--#include virtual="/setup.asp" -->
<%
'PARTE PERSO PASSWORD-----------------------------------------------------------------
if request.form("email_request")<> "" then
	sqlc="select SQL_CALC_FOUND_ROWS *, utenti.iduser FROM utenti left join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id  WHERE email='"&pulisci(request.form("email_request"))&"'"
	testolog=" per email "&request.form("email_request")
end if
if request.form("pi_request")<> "" then
	sqlc="select SQL_CALC_FOUND_ROWS *, utenti.iduser FROM utenti inner join utenti_intestazioni on utenti.idintestazione=utenti_intestazioni.id  WHERE utenti_intestazioni.piva='"&pulisci(request.form("pi_request"))&"'"
	testolog=" per partita iva "&request.form("pi_request")
	testo2="<p><strong>Se l'indirizzo email non &egrave; di tua propriet&agrave; puoi effettuare una nuova registrazione</strong> <a href=""register-account.asp"" >Registrati</a></p>"
session("sql")=sqlc
end if
if request.form("cf_request")<> "" then
	sqlc="select SQL_CALC_FOUND_ROWS * FROM utenti inner join utenti_intestazioni on utenti.idintestazione=utenti_intestazioni.id  WHERE utenti_intestazioni.cf='"&pulisci(request.form("cf_request"))&"'"
	testolog=" per codice fiscale "&request.form("cf_request")
	testo2="<p><strong>Se l'indirizzo email non &egrave; di tua propriet&agrave; puoi effettuare una nuova registrazione</strong> <a href=""register-account.asp"" >Registrati</a></p>"
end if
if sqlc<>"" then
	on error resume next
	Set rs = conn.execute(sqlc)
	if err.number<>0 then
		call add2log(sqlc,0)
		response.end
		
	end if
	'rs.Open sqlc, conn, 3, 3
	lngTotalRecords=clng(conn.Execute("Select Found_Rows();")(0).Value)
	if lngTotalRecords>1 then
		'passerror="<br>I dati forniti corrispondono a pi&ugrave; di un utente, ci contatti via email.<br>"
		'add2log "Richiesta dati di accesso "&testolog&", trovate "&lngTotalRecords&" corrispondenze, dati NON inviati",3
		'call inserisci_appunto("Richiesta dati di accesso "&testolog&", trovate "&lngTotalRecords&" corrispondenze, dati NON inviati",array(4,1))
		
		step="lista_mail"
		
		
		
	elseif lngTotalRecords=1 then
		set rs_password=conn.execute("select * from password_reset where iduser="&rs("iduser"))
		ora_invio=date()-1
		if not rs_password.eof then
			ora_invio=rs_password("ora_invio")
		end if
		set rs_password = nothing
	
		if datediff("n",ora_invio,now())<6 then
			passerror="Una password di accesso &egrave; stata inviata al tuo indirizzo email meno di 5 minuti fa.<br>Controlla la posta e verifica che l'email non sia stata bloccata dal filtro anti spam.<br><br>Devi aspettare almeno 5 minuti per richiedere una nuova password.<br>"
		else
			iduser=rs("iduser")
			email_utente=rs("email")
			nominativo_v=denominazione(rs("nomE"),rs("cognome"),rs("azienda"))
			if email_utente<>"" then
				add2log "Richiesta dati di accesso [utente="&iduser&"]"&nominativo_v&"[/utente] "&testolog&", email inviata all'indirizzo "& email_utente,2
				
				res = invia_mail_registrazione(iduser,false)
				if res then 
					call inserisci_appunto("<a href=""pag_adm_user.asp?iduser="&iduser&"""><b>"&nominativo_v&"</b></a> ha richiesto i dati di accesso che sono stati inviati all'indirizzo email <b>"&email_utente&"</b>",array(4,1))
				else
					call inserisci_appunto("<a href=""pag_adm_user.asp?iduser="&iduser&"""><b>"&nominativo_v&"</b></a> ha richiesto i dati di accesso,  l'invio all'indirizzo email <b>"&email_utente&"</b> &egrave; <b>FALLITO</b>",array(4,1))
				end if
				logga "lost",request.form("email")
				session("messaggioreset")="I dati di accesso sono stati inviati all'indirizzo email <strong>"&email_utente&"</strong>, controlla la posta."&testo2
				
				response.redirect "login.asp"
			else
				add2log "Richiesta dati di accesso "&testolog&", indirizzo email vuoto",2
				passerror="<br>I dati forniti hanno consentito di risalire ai suoi dati di accesso ma non &egrave; presente un indirizzo email legato al suo account, ci contatti via email.<br>"
			end if
		end if
	else
		passerror="<br>I dati forniti non hanno consentito di risalire ai suoi dati di accesso.<br>"
		add2log "Richiesta dati di accesso "&testolog&", utente non trovato",2
		call inserisci_appunto("Richiesta dati di accesso "&testolog&", utente non trovato",array(4,1))
	end if
		
	'rs.close
	set rs = nothing
end if
if request.form("email_iduser")<>"" then
	iduser=pulisci(request.form("email_iduser"))
	
	set rs_password=conn.execute("select * from password_reset where iduser="&iduser)
	ora_invio=date()-1
	if not rs_password.eof then
		ora_invio=rs_password("ora_invio")
	end if
	set rs_password = nothing

	if datediff("n",ora_invio,now())<6 then
		passerror="Una password di accesso &egrave; stata inviata al tuo indirizzo email meno di 5 minuti fa.<br>Controlla la posta e verifica che l'email non sia stata bloccata dal filtro anti spam.<br><br>Devi aspettare almeno 5 minuti per richiedere una nuova password.<br>"
	else
		email_utente=""
		res = invia_mail_registrazione(iduser,false)
		nominativo_v=get_denominazione(iduser)
	
		if res then 
			call inserisci_appunto("<a href=""pag_adm_user.asp?iduser="&iduser&"""><b>"&nominativo_v&"</b></a> ha richiesto i dati di accesso che sono stati inviati all'indirizzo email <b>"&email_utente&"</b>",array(4,1))
		else
			call inserisci_appunto("<a href=""pag_adm_user.asp?iduser="&iduser&"""><b>"&nominativo_v&"</b></a> ha richiesto i dati di accesso,  l'invio all'indirizzo email <b>"&email_utente&"</b> &egrave; <b>FALLITO</b>",array(4,1))
		end if
		call connclose()
		session("messaggioreset")="I dati di accesso sono stati inviati all'indirizzo email <strong>"&email_utente&"</strong>, controlla la posta."&testo2
		
		response.write "qui"
		response.end
		response.redirect "login.asp"
	end if
	
	
end if

'FINE PARTE PERSO PASSWORD-----------------------------------------------------------------
meta_title="tit20"
%>
<!--#include virtual="/config/header_inc.asp" -->
        <section id="content">
        	<div id="breadcrumb-container">
        		<div class="container">
					<ul class="breadcrumb">
						<li><a href="/">Home</a></li>
						<li class="active"><%=traduci("tit20")%></li>
					</ul>
        		</div>
        	</div>
        	<div class="container">
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=traduci("tit20")%></h1>
							<%if step="" then%>
                            <div class="md-margin"><%=traduci("tit20txt")%></div><!-- space -->
			                <%end if%>
						</header>
						   <div class="row">
							   
							   
   			   	                <%if passerror<>"" then %>
	   	                 		<div class="alert alert-danger">
                                    <strong>Errore!</strong>&nbsp; <%=passerror%>
                                </div>
			   	                <%end if%> 

							   	<div class="col-md-6 col-sm-6 col-xs-12">	
									<form id="login-form" method="post" action="<%=Request.ServerVariables("Script_Name")&"?comebackto="  &sReferer%>">
							   	<%if step="lista_mail" then
							   	%>
								   	<p>Sono stati trovati <%=lngTotalRecords%> indirizzi email.<br>
								   	Seleziona a quale indirizzo email vuoi inviare i dati di accesso</p>
							   	<%
							   	set rs=conn.execute(sqlc)
							   	do while not rs.EOF
							   	%>
							   	<input type="radio" name="email_iduser" value="<%=rs("iduser")%>"> <%=rs("email")%><br>
							   	<%
							   	rs.MoveNext
							   	loop
							   	set rs = nothing
							   	 %>
                                 <div class="xs-margin"></div>
                                <button class="btn btn-custom-2">INVIA DATI</button>
							   	</form>
							   	
							   	<%end if	'if step="lista_mail" then %>
							   	<%if step="due" then %>
								<div class="alert alert-success">
								 I dati sono stati inviati all'indirizzo email <strong><%=email_utente%></strong>, controlla la posta.
	   	                 		</div>
								<%end if%>	
						 <%if step="" then%>       			
							   				   		
							   		
									<form id="login-form" method="post" action="<%=Request.ServerVariables("Script_Name")&"?comebackto="  &sReferer%>">
                                    	<p>Recupera le credenziali con l'indirizzo email</p>
                                         <div class="input-group xs-margin">
                                            <span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text">Email</span></span>
                                            <input type="text"  class="form-control input-lg" placeholder="<%=traduci("emailpl")%>" name="email_request">
                                        </div><!-- End .input-group -->
                                        <%if false then %>
                                        <p>Recupera le credenziali con la partita iva</p>
                                         <div class="input-group xs-margin">
                                            <span class="input-group-addon"><span class="input-icon input-icon-password"></span><span class="input-text"><%=traduci("pi")%></span></span>
                                            <input type="text"  class="form-control input-lg" placeholder="<%=traduci("pipl")%>" name="pi_request">
                                        </div><!-- End .input-group -->
                                        <%end if %>
                                        <%if true then %>
                                        <p>Recupera le credenziali con il codice fiscale</p>
                                        <div class="input-group xs-margin">
                                            <span class="input-group-addon"><span class="input-icon input-icon-password"></span><span class="input-text"><%=traduci("cf")%></span></span>
                                            <input type="text"  class="form-control input-lg" placeholder="<%=traduci("cfpl")%>" name="cf_request">
                                        </div><!-- End .input-group -->
                                        <%end if %>
                                    <button class="btn btn-custom-2"><%=traduci("richiedi")%></button>
                                    </form>
                                    						   <%end if%>
                                    <div class="sm-margin"></div><!-- space -->
							   	</div><!-- End .col-md-6 -->
							   	
						   </div><!-- End.row -->
								   
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
			</div><!-- End .container -->
        
        </section><!-- End #content -->
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
    </body>
</html>