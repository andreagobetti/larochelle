<!--#include virtual="/setup.asp" -->
<!--#include file="ClasseUtente.asp"-->
<!--#include file="ClasseModificheRS.asp"-->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<%
if request("registrati")<>"" then
	txt_errore=""
	'Controllo validazione form eseguita
	if request.form("form_valido")="xx" then
		txt_errore="Registrazione non valida"
		call banna_ip()
		add2log "Bloccata registrazione per validazione form bypassata.<br>"&Request.ServerVariables("REMOTE_ADDR")&" aggiunto all'elenco IP bannati"&vbcrlf&queryeform(),0
	end if
		
	
	'Controllo email valida
	if txt_errore="" then
		email=lcase(trim(request.form("email")))
		if not validemail(email) then
			txt_errore="L'indirizzo email non &egrave; valido"
		end if
		call determina_sqlinjection(email)
	end if
	
	'Controllo email univoca
	if txt_errore="" then
		set rs=conn.execute("select utenti_intestazioni.* from utenti_intestazioni inner join utenti on utenti_intestazioni.id = utenti.idintestazione where utenti.email='"&email&"'")
		if not rs.eof then
			txt_errore="L'indirizzo email &egrave; gi&agrave; utilizzato da un altro utente"
			add2log "Bloccata registrazione per email "&email&" gi&agrave; esistente utilizzata da [utente="&rs("iduser")&"]"&nominativo(rs("nome"),rs("cognome"),rs("azienda"))&"[/utente]"&vbcrlf&queryeform(),2
		end if
		rs.close
		set rs = Nothing
	end if
	
		'riverifico lato server partita e iva
		azienda=trim(request.form("azienda"))
		cf=trim(request.form("cf"))
		call determina_sqlinjection(request.form("piva"))
		if false then	'if piva<>"" then
			//call determina_sqlinjection(request.form("azienda"))
			txt="Verifica lato server azienda: "&azienda&" e cf: "&cf
			sql="select * from utenti_intestazioni where cf='"&cf&"'"
			set rs=conn.execute (sql)
			iduser=0
			if not rs.eof then
				iduser=rs("iduser")
			else
				txt=txt&" corrispondenza non trovata sql:"&sql
			end if
			if iduser>0 then 
				set rs=conn.execute("select * from utenti where iduser="&iduser)
				email=rs("email")
				txt=txt&"<br>Tentata registrazione di cliente già esistente iduser:"&iduser&" con questi dati:"&queryeform()
				txt_errore="<b>Attenzione:</b> la partita iva indicata &egrave; gi&agrave; presente nel nostro archivio</b><br>Potete utilizzare <a href=""password-reset.asp"">questa pagina</a> per recuperare i dati di accesso."
			end if
			set rs = Nothing
			call add2log(txt,2)
		end if
	
	
	
	
	if txt_errore="" then
		set utente= (new Classeutente)("add_da_user")
		iduser= utente.get_iduser()
	%>
	<!--#include virtual="/mail_registrazione_inc.asp" -->
	<%	
		res = invia_mail_registrazione(iduser,true)
		
		account_creato=true
	end if
end if
%>
<%
tipo_registrazione=request.form("tipo_registrazione")
if tipo_registrazione<>"" then
	if request.form("privato")<>"" then
		tipo_registrazione="privato"
	elseif request.form("pa")<>"" then
		tipo_registrazione="pa"
	end if
end if
css_validator=true
'meta_title=traduci("tit10")
%>
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
        	<%if account_creato then %>
        		<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=traduci("tit14")%></h1>
							<p class="title-desc"><%=traduci("tit14txt")%></p>
						</header>
				</div><!-- End .col-md-12 -->
        	<%elseif tipo_registrazione="" then%>
				<div class="row">
					<%if txt_errore<>"" then%>
						<div class="alert alert-danger">
							<strong>Errore!</strong>&nbsp; <%=txt_errore%>
						</div>
					<%end if%>

        			<div class="col-md-12 ">
						<header class="content-title">
							<h1 class="title">Registrazione</h1>
							<p class="title-desc">Scegli il modulo di registrazione</p>
						</header>
        				<div class="xs-margin"></div><!-- space -->
						<form action="<%=questofile%>" method="post" >
						<input type="hidden" name="tipo_registrazione" value="tipo">
        				<div class="row">
								<div class="col-md-6 col-sm-6 col-xs-12">
								<p>Modulo di registrazione per clienti privati</p>
									<p class="center"><input type="submit" value="Privato" class="btn btn-custom-2 btn-lg md-margin" name="privato"></p>
								</div><!-- End .col-md-6 -->
        						<div class="col-md-6 col-sm-6 col-xs-12">
								<p>Modulo di registrazione per la pubblica amministrazione, società e aziende</p>
								<p><input type="submit" value="PA o societa'" class="btn btn-custom-2 btn-lg md-margin" name="pa"></p>
        						</div><!-- End .col-md-6 -->
        				</div><!-- End .row -->
        				</form>
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
			<%elseif tipo_registrazione="privato" then
			%>
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=traduci("tit10")%></h1>
							<p class="title-desc"><%=traduci("tit10txt")%></p>
						</header>
        				<div class="xs-margin"></div><!-- space -->
						<form action="<%=questofile%>" method="post" id="register-form">
						<input type="hidden" name="tipo_form" value="<%=tipo_registrazione%>">

        				<div class="row">
							<%if txt_errore<>"" then%>
								<div class="alert alert-danger">
                                    <strong>Errore!</strong>&nbsp; <%=txt_errore%>
                                </div>
							<%end if%>
								<div class="col-md-6 col-sm-6 col-xs-12">
									<fieldset>
									<h2 class="sub-title"><%=traduci("tit11")%></h2>
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("nome")%>&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("nomepl")%>" name="nome">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("cognome")%>&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("cognomepl")%>" name="cognome">
									</div><!-- End .input-group -->
                                    <div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-company"></span><span class="input-text"><%=traduci("azienda")%></span></span>
										<input type="text" class="form-control input-lg" placeholder="<%=traduci("aziendapl")%>" name="azienda">
									</div><!-- End .input-group -->
									
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text">Email&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("emailpl")%>" name="email" id="email">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text"><%=traduci("repemail")%>&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("repemailpl")%>" name="repeat_email">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-phone"></span><span class="input-text"><%=traduci("telefo")%>&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("telefopl")%>" name="telefono">
									</div><!-- End .input-group -->
									<div class="input-group custom-checkbox">
										 <input type="checkbox" name="privacy"> <span class="checbox-container">
										 <i class="fa fa-check"></i>
										 </span>
										 <%=traduci("priv")%>
									</div><!-- End .input-group -->
									
									</fieldset>
								</div><!-- End .col-md-6 -->
        						
        						<div class="col-md-6 col-sm-6 col-xs-12">
                                    <fieldset class="half-margin">
									<h2 class="sub-title">NEWSLETTER</h2>
										<div class="input-desc-box">
										<span class="input-group custom-checkbox">
									 <input type="checkbox" name="newsletter"> <span class="checbox-container">
									 <i class="fa fa-check"></i>
									 </span><%=traduci("newsl")%></span>
										</div><!-- End .input-desc -->
									</fieldset>
									
								
        						</div><!-- End .col-md-6 -->

        				</div><!-- End .row -->

						<div class="row">
							<div class="col-md-12 col-sm-12 col-xs-12 text-center">

								
								<input type="submit" value="<%=traduci("but10")%>" class="btn btn-custom-2 btn-lg md-margin" name="registrati" id="registratibtn">
							</div><!-- End .col-md-12 -->
						</div><!-- End .row -->
						<input type="hidden" id="form_valido" name="form_valido" value="xx"> 
        				</form>
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
        		
        		
			<%elseif tipo_registrazione="pa" then
			%>
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=traduci("tit10")%></h1>
							<p class="title-desc"><%=traduci("tit10txt")%></p>
						</header>
        				<div class="xs-margin"></div><!-- space -->
						<form action="<%=questofile%>" method="post" id="register-form">
						<input type="hidden" name="tipo_form" value="<%=tipo_registrazione%>">

        				<div class="row">
							<%if txt_errore<>"" then%>
								<div class="alert alert-danger">
                                    <strong>Errore!</strong>&nbsp; <%=txt_errore%>
                                </div>
							<%end if%>
								<div class="col-md-6 col-sm-6 col-xs-12">
									<fieldset>
									<h2 class="sub-title"><%=traduci("tit11")%></h2>
                                    <div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-company"></span><span class="input-text"><%=traduci("azienda")%>&#42;</span></span>
										<input type="text" class="form-control input-lg" placeholder="<%=traduci("aziendapl")%>" name="azienda" id="azienda">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-subject"></span><span class="input-text">Codice fiscale&#42;</span></span>
										<input type="text" class="form-control input-lg" placeholder="Codice fiscale azienda" name="cf" id="piva">
									</div><!-- End .input-group -->
									
									
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("nome")%>&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("nomepl")%>" name="nome">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("cognome")%>&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("cognomepl")%>" name="cognome">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text">Ruolo o funzione</span></span>
										<input type="text"  class="form-control input-lg" placeholder="Ruolo o funzione nell'azienda" name="ruolo">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text">Email&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("emailpl")%>" name="email" id="email">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text"><%=traduci("repemail")%>&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("repemailpl")%>" name="repeat_email">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-phone"></span><span class="input-text"><%=traduci("telefo")%>&#42;</span></span>
										<input type="text" required class="form-control input-lg" placeholder="<%=traduci("telefopl")%>" name="telefono">
									</div><!-- End .input-group -->
									<div class="input-group custom-checkbox">
										 <input type="checkbox" name="privacy"> <span class="checbox-container">
										 <i class="fa fa-check"></i>
										 </span>
										 <%=traduci("priv")%>
									</div><!-- End .input-group -->
									
									</fieldset>
								</div><!-- End .col-md-6 -->
        						
        						<div class="col-md-6 col-sm-6 col-xs-12">
                                    <fieldset class="half-margin">
									<h2 class="sub-title">NEWSLETTER</h2>
										<div class="input-desc-box">
										<span class="input-group custom-checkbox">
									 <input type="checkbox" name="newsletter"> <span class="checbox-container">
									 <i class="fa fa-check"></i>
									 </span><%=traduci("newsl")%></span>
										</div><!-- End .input-desc -->
									</fieldset>
									
								
        						</div><!-- End .col-md-6 -->

        				</div><!-- End .row -->

						<div class="row">
							<div class="col-md-12 col-sm-12 col-xs-12 text-center">

								
								<input type="submit" value="<%=traduci("but10")%>" class="btn btn-custom-2 btn-lg md-margin" name="registrati" id="registratibtn">
							</div><!-- End .col-md-12 -->
						</div><!-- End .row -->
						<input type="hidden" id="form_valido" name="form_valido" value="xx"> 
        				</form>
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->

        		<%end if %>
			</div><!-- End .container -->
        
        </section><!-- End #content -->
        <%
	    'Memorizzo lunghezza campi prima di chiudere la connessione
	    Set rs_utente = Server.CreateObject("ADODB.Recordset")
		sql="select utenti.email, utenti.telefono, utenti_intestazioni.azienda, utenti_intestazioni.nome, utenti_intestazioni.cognome FROM utenti left join utenti_intestazioni on utenti.idintestazione = utenti_intestazioni.id "

	    
	    
		rs_utente.Open sql, conn
	    email_DefinedSize=rs_utente("email").DefinedSize
	    nome_DefinedSize=rs_utente("nome").DefinedSize
	    cognome_DefinedSize=rs_utente("cognome").DefinedSize
	    telefono_DefinedSize=rs_utente("telefono").DefinedSize
	    azienda_DefinedSize=rs_utente("azienda").DefinedSize
		rs_utente.close
		set rs_utente=nothing
	    
	 %>    
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
	<script type="text/javascript" src="chk_piva_cf.js"></script>

	<script src="js/Jqueryvalidate/jquery.validate.min.js"></script>
	<script src="js/Jqueryvalidate/localization/messages_it.min.js"></script>
	<script>
	$(document).ready(function() {
	$.validator.addMethod("phoneitaly", function(phone_number, element) {
	return this.optional(element) || phone_number.length > 9 &&
		phone_number.match(/^((\+|00)?39\s)?(0\d{1,3}|3\d{2})\/\d{5,8}$/);
	}, "Specificare un numero di telefono valido: +39 XXXX/YYYYYYYY");
	
	$.validator.addMethod("partita_iva", function (value, element) {
		var txt="";
		var result = true;
		if (value=='')
		{
			$.validator.messages.partita_iva = "Campo obbligatorio";
			result = false
			return false;
		}
		else
		{
			chk=ControllaPIVA(value);
			if (chk != '' ) { 
				$.validator.messages.partita_iva = chk;
				result = false; 
				return false;
			}
		}
		return true;
		
	}, "");

	
	$("#register-form").validate({
		rules: {
			<%if tipo_registrazione="pa" then %>
			azienda: {required:true,minlength:3,maxlength:<%=azienda_DefinedSize%>},
			piva: { partita_iva: true
//				,
//				remote:{
//	                url: "searcher.asp", //make sure to return true or false with a 200 status code
//	                type: "post",
//	                dataType: 'json',
//	                cache:false,
//					data: {oper: "check_azienda_piva",azienda:function(){return $("#azienda").val()} ,piva:function(){return $("#piva").val()}},
//
//	                dataFilter: function(data) {
//				        var json = JSON.parse(data);
//				        if(json.isError) {
//				            return "\"" + "<b>Attenzione:</b> questi dati sono associati all' indirizzo email <b>"+decodeURIComponent(encodeURIComponent(json.email)) + "</b><br>Utilizzare questa email e la password per accedere al sito.\"";
//				        } else {
//				            return true;
//				        }
//				
//				    },
//	                
//	                error: function(xhr, textStatus, errorThrown)
//				        {
//				            alert('ajax loading error... xhr:'+xhr+'textStatus:'+textStatus+'errorThrown:'+errorThrown);
//				            return false;
//				        }
//	            }
				
				
				
				
				 },
			<%end if %>
			
			
			nome: {required:true,minlength:3,maxlength:<%=nome_DefinedSize%>},
			cognome: {required: true,minlength: 3,maxlength:<%=cognome_DefinedSize%>},
			cf:{required:true},
			email:{
				required:true,email:true,maxlength:<%=email_DefinedSize%>
			,
			remote:{
                url: "searcher.asp?check_email=si", //make sure to return true or false with a 200 status code
                type: "post",
                dataType: 'json',
                cache:false,
                
                dataFilter: function(data) {
			        var json = JSON.parse(data);
			        if(json.isError == "true") {
			            return "\"" + '<a href=\'password-reset.asp\'>Questa email esiste gi&agrave;, recupera la password cliccando QUI</a>' + "\"";
			        } else {
			            return true;
			        }
			
			    },
                
                error: function(xhr, textStatus, errorThrown)
			        {
			            alert('ajax loading error... xhr:'+xhr+'textStatus:'+textStatus+'errorThrown:'+errorThrown);
			            return false;
			        }
            }
        },
		repeat_email:{equalTo:"#email"},			
			telefono:{required:true,digits: true,minlength: 8},			
			privacy: {required:true}
			
		},
	
		errorElement: "div",
	    errorClass: "help-block",
	    highlight: function (element, errorClass, validClass) {
	        $(element).closest('.input-group').addClass('has-error');
	    },
	    unhighlight: function (element, errorClass, validClass) {
	        $(element).closest('.input-group').removeClass('has-error');
	    },
	    errorPlacement: function (error, element) {
	        if (element.parent('.input-group').length || element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
	            error.insertAfter(element.parent());
	        } else {
	            error.insertAfter(element);
	        }
	    },
		submitHandler: function(form){
            $('#form_valido').val('si');
            $("#registratibtn").prop("value","Attendi");
            //$("#registratibtn").prop("disabled","disabled");
			form.submit();
		}


	});

});
	</script>

    </body>
</html>