<!--#include virtual="/setup.asp" -->
<!--#include file="ClasseUtente.asp"-->
<!--#include file="ClasseModificheRS.asp"-->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<%
		err=request("err")
if session("iduser") = "" then call login()
select2=true
if request.form("btn1")<>"" then
	iduser=session("iduser")
	set utente= (new Classeutente)(array("update_da_user",iduser))
	bt1errore=utente.errore()
	set utente = Nothing
	btn1=true
end if
if request.form("btn2")<>"" then
	new_mail=trim(request.form("email"))
	if new_mail<>trim(request.form("repeat_email")) then
		btn2ko1=true
		add2log "Bloccato cambio email per email non corrispondenti",3
	elseif new_mail="" then
		btn2ko2=true
		add2log "Bloccato cambio email per email vuota",3
	else
		set rs_email=conn.execute("select utenti.* from utenti where email='"&new_mail&"'")
		if not rs_email.eof then
			txt_errore="L'indirizzo email è gi&agrave; utilizzato da un altro utente"
			btn2ko3=true
			add2log "Bloccato cambio email "&email&" gi&agrave; esistente utilizzata da [utente="&rs_email("iduser")&"]"&nominativo(rs_email("nome"),rs_email("cognome"),rs_email("azienda"))&"[/utente]",2
			
			
		else
			iduser=session("iduser")
			Set rs_utente = Server.CreateObject("ADODB.Recordset")
			rs_utente.Open "select * from utenti where iduser="&iduser, conn, 1, 3
			rs_utente("email")=trim(request.form("email"))
			rs_utente.update
			rs_utente.close
			set rs_utente=nothing
			btn2=true	
			add2log "L'utente [utente="&iduser&"]"&session("nominativo")&"[/utente] ha modificato il proprio indirizzo email"&vbcrlf&queryeform(),1
			
		end if
		set rs_email=Nothing
	end if
end if
if request.form("btn3")<>"" then
	
	
	
	
	
	iduser=session("iduser")
	
	
	
	passwordInChiaro=trim(request.form("password"))
	set utente= (new Classeutente)(array("new_password",iduser,passwordInChiaro))

	btn3=true
	add2log "L'utente [utente="&iduser&"]"&session("nominativo")&"[/utente] ha modificato la propria password"&vbcrlf&"[soloio]"&"pw:"&passwordInChiaro&"[/soloio]",1
end if
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
			<%if btn1 and bt1errore="" then%>
		   		<div class="alert alert-success">
                    <strong>OK!</strong>&nbsp; <%=traduci("txt13")%>
                </div>
			<%end if%>
			<%if bt1errore<>"" then%>
		   		<div class="alert alert-danger">
                    <strong><%=traduci("error")%>!</strong>&nbsp; <%=bt1errore%>
                </div>
			<%end if%>
			
			
			
			
			<%if btn2 then%>
		   		<div class="alert alert-success">
                    <strong>OK!</strong>&nbsp; <%=traduci("txt14")%>
                </div>
			<%end if%>
			<%if btn2ko1 then%>
		   		<div class="alert alert-danger">
                    <strong><%=traduci("error")%>!</strong>&nbsp; <%=traduci("err1")%>
                </div>
			<%end if%>
			<%if btn2ko2 then%>
		   		<div class="alert alert-danger">
                    <strong><%=traduci("error")%>!</strong>&nbsp; <%=traduci("err2")%>
                </div>
			<%end if%>
			<%if btn3 then%>
		   		<div class="alert alert-success">
                    <strong>OK!</strong>&nbsp; <%=traduci("txt15")%>
                </div>
			<%end if%>
			
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=traduci("tit10")%></h1>
							<p class="title-desc"><%=traduci("tit10txt")%></p>
						</header>
						<%
								sql="SELECT * FROM utenti where "&where_carrello
								set rs_utente=conn.execute(sql)
								id_user=rs_utente("iduser")
								pres_telefono=rs_utente("telefono")
								pres_email=rs_utente("email")
								secondario=rs_utente("secondario")
								pres_ruolo=rs_utente("ruolo")
								set rs_i=conn.execute("select * from utenti_intestazioni where id="&rs_utente("idintestazione"))
								pres_nome=rs_i("nome")
								pres_cognome=rs_i("cognome")
								pres_azienda=rs_i("azienda")
								pres_cf=rs_i("cf")
								pres_piva=rs_i("piva")
								pres_indirizzo=rs_i("indirizzo")
								pres_citta=rs_i("citta")
								pres_cap=rs_i("cap")
								pres_provincia=rs_i("provincia")
								set rs_i = Nothing
								set rs_utente=nothing
								nfatture=clng(conn.execute ("SELECT Count(*) FROM fatture WHERE "&where_carrello)(0).Value)
								if nfatture=0 then
									piva_editabile="SI"
									class_disabled=""
								else
									piva_editabile="NO"
									class_disabled=" disabled"
								end if
								set rs_fatture=Nothing
								%>
        				<div class="panel-group custom-accordion" id="checkout">
        				<form action="<%=questofile%>" method="post" id="dati-form">
								<input type="hidden" name="piva_editabile" value="<%=piva_editabile%>">
								<div id="checkout-option" class="collapse in">
								  <div class="panel-body">
								   <div class="row">
								   	<div class="col-md-6 col-sm-6 col-xs-12">	
									   	<div class="input-group">
											<span class="input-group-addon">
											<span class="input-icon input-icon-user"></span><span class="input-text">ID <%=traduci("user")%></span></span>
											<span class="form-control input-lg" ><%=id_user%></span>
										</div><!-- End .input-group -->
								   	</div><!-- End .col-md-6 -->
								   	<% if secondario=1 then %>
									   	<div class="col-md-6 col-sm-6 col-xs-12">	
									   	<div class="input-group">
											<span class="input-group-addon">
											<span class="input-icon input-icon-company"></span><span class="input-text">Azienda</span></span>
											<span class="form-control input-lg" ><%=pres_azienda%></span>
										</div><!-- End .input-group -->
								   	</div><!-- End .col-md-6 -->
								   	<% end if %>
								   </div><!-- End.row -->
								  </div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
							  
								
								  <fieldset>
								   <div class="row">
								   	<div class="col-md-6 col-sm-6 col-xs-12">
								   		
								   		<h2 class="checkout-title"><%=traduci("tit11")%></h2>
								   		
								   		<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("nome")%>&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("nomepl")%>" name="nome" value="<%=pres_nome%>">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text"><%=traduci("cognome")%>&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("cognomepl")%>" name="cognome" value="<%=pres_cognome%>">
									</div><!-- End .input-group -->
								   	<% if secondario=0 then %>
                                    <div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-company"></span><span class="input-text"><%=traduci("azienda")%></span></span>
										<input type="text" class="form-control input-lg" placeholder="<%=traduci("aziendapl")%>" name="azienda" id="azienda" value="<%=pres_azienda%>">
									</div><!-- End .input-group -->
									<%end if %> 
									
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-phone"></span><span class="input-text"><%=traduci("telefo")%>&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("telefopl")%>" name="telefono" value="<%=pres_telefono%>">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-company"></span><span class="input-text">Ruolo o funzione</span></span>
										<input type="text"  class="form-control input-lg" placeholder="Il tuo ruolo o funzione" name="ruolo" value="<%=pres_ruolo%>">
									</div><!-- End .input-group -->
									
									
								   	<% if secondario=0 then %>

									<%if class_disabled<>"" then%>
									<%=traduci("nomodifica")%>
									
									<%end if%>
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-subject"></span><span class="input-text"><%=traduci("cf")%>&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("cfpl")%>" name="cf" value="<%=pres_cf%>" <%=class_disabled%>>
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-subject"></span><span class="input-text"><%=traduci("piva")%></span></span>
										<input type="text" class="form-control input-lg" placeholder="<%=traduci("pivapl")%>" name="piva" value="<%=pres_piva%>" <%=class_disabled%>>
									</div><!-- End .input-group -->
									<%end if %>
								   	</div><!-- End .col-md-6 -->
								   	<% if secondario=0 then %>
								   	<div class="col-md-6 col-sm-6 col-xs-12">
										<h2 class="checkout-title"><%=traduci("tit12")%></h2>
	
										<div class="input-group">
											<span class="input-group-addon"><span class="input-icon input-icon-address"></span><span class="input-text"><%=traduci("indir")%>&#42;</span></span>
											<input type="text" class="form-control input-lg" placeholder="<%=traduci("indirpl")%>" name="indirizzo" value="<%=pres_indirizzo%>">
										</div><!-- End .input-group -->
										
										<div class="input-group select2-bootstrap-prepend">
											<span class="input-group-addon"><span class="input-icon input-icon-city"></span>
											<span class="input-text"><%=traduci("citta")%>&#42;</span></span>
											<input type="text" class="form-control input-lg" name="citta" id="citta" placeholder="<%=traduci("cittapl")%>" value="<%=pres_citta%>" ><input type="hidden" id="pres_citta" value="<%=pres_citta%>">
										</div><!-- End .input-group -->
										
																			
										<div class="input-group">
											<span class="input-group-addon"><span class="input-icon input-icon-postcode"></span><span class="input-text"><%=traduci("cap")%>&#42;</span></span>
											<input type="text"  class="form-control input-lg" name="cap" placeholder="<%=traduci("cappl")%>" value="<%=pres_cap%>" id="cap">
										</div><!-- End .input-group -->
	
	
	
										<div class="input-group lg-margin select2-bootstrap-prepend">
											<span class="input-group-addon"><span class="input-icon input-icon-region"></span><span class="input-text"><%=traduci("prov")%>&#42;</span></span>
											
											<select name="provincia" class="select2 form-control" id="provincia">
												<%   
												Set rs_provincia = conn.execute("select * from elenco_province order by Denominazione_provincia")
												do while not rs_provincia.eof 
												if ucase(rs_provincia("Sigla_automobilistica"))=pres_provincia then txt_selected="SELECTED" else txt_selected="" end if
												%>
                                                <option  value="<%=rs_provincia("Sigla_automobilistica")%>" <%=txt_selected%>><%=rs_provincia("Denominazione_provincia")%></option>
                                                <%
                                                rs_provincia.movenext
                                                loop
                                                set rs_provincia=nothing
                                                %>
                                            </select>
										</div><!-- End .input-group -->
								   	</div><!-- End .col-md-6 -->
								   	<% end if %>
								   </div><!-- End .row -->
									   <input type="hidden" name="regione" id="regione">
									   <div class="col-md-12 text-right">
											<input type="submit" value="<%=traduci("btn1")%>" class="btn btn-custom-2" id="btn1" name="btn1">
									   </div>
								  </fieldset>
        				</form>
<!-- Email-->
	          				<form action="<%=questofile%>" method="post" id="email-form">
							  <div class="panel">
  								<div class="accordion-header">
									<div class="accordion-title"><span><%=traduci("step3")%></span></div><!-- End .accordion-title -->
									<a class="accordion-btn"   data-toggle="collapse" data-target="#email-change"></a>
								</div><!-- End .accordion-header -->
							  
								<div id="email-change" class="collapse">
								  <div class="panel-body">
								   <div class="row">
								   	<div class="col-md-6 col-sm-6 col-xs-12">	
									   	<div class="input-group">
											<span class="input-group-addon">
											<span class="input-icon input-icon-email"></span><span class="input-text">Email <%=traduci("user")%></span></span>
											<span class="form-control input-lg" ><%=pres_email%></span>
										</div><!-- End .input-group -->
								   	</div><!-- End .col-md-6 -->
							   	
								   </div><!-- End.row -->
								  <p><%=traduci("txt11")%></p>
								   <div class="row">
								   	<div class="col-md-6 col-sm-6 col-xs-12">
										<div class="input-group" >
                                            <span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text"><%=traduci("newemail")%>&#42;</span></span>
                                            <input type="text"  class="form-control input-lg" placeholder="<%=traduci("pl_email")%>" name="email" id="email">
                                        </div><!-- End .input-group -->
								   	</div>
								  
								   	<div class="col-md-6 col-sm-6 col-xs-12">
                                        <div class="input-group">
                                            <span class="input-group-addon"><span class="input-icon input-icon-email"></span><span class="input-text"><%=traduci("repemail")%>&#42;</span></span>
                                            <input type="text"  class="form-control input-lg" placeholder="<%=traduci("repemailpl")%>" name="repeat_email">
                                        </div><!-- End .input-group -->
								   	</div>
									<div class="col-md-12 text-right">
										<input type="hidden" name="btn2" value="si">
										<input type="submit" value="<%=traduci("btn2")%>" class="btn btn-custom-2" id="btn2" >
								   </div>
								   </div>
								  </div><!-- End .panel-body -->
								</div><!-- End .panel-collapse -->
								
							  
							  </div><!-- End .panel -->
	        				</form>
<!-- Email-->
	          				<form action="<%=questofile%>" method="post" id="password-form">
							  <div class="panel"><!-- .panel 4 step--><!-- .panel 4 step--><!-- .panel 4 step--><!-- .panel 4 step--><!-- .panel 4 step-->
								<div class="accordion-header">
									<div class="accordion-title"><span><%=traduci("step4")%></span></div><!-- End .accordion-title -->
									<a class="accordion-btn"  data-toggle="collapse" data-target="#delivery-method"></a>
								</div><!-- End .accordion-header -->
								
								<div id="delivery-method" class="collapse">
								  <div class="panel-body">
									  <p><%=traduci("txt12")%></p>
								   <div class="row">
								   	<div class="col-md-6 col-sm-6 col-xs-12">
										<div class="input-group">
											<span class="input-group-addon"><span class="input-icon input-icon-password"></span><span class="input-text"><%=traduci("pw1")%>&#42;</span></span>
											<input type="password"  class="form-control input-lg" placeholder="<%=traduci("pw1pl")%>" name="password" id="password">
										</div><!-- End .input-group -->
								   	</div>
								   	<div class="col-md-6 col-sm-6 col-xs-12">
										<div class="input-group">
											<span class="input-group-addon"><span class="input-icon input-icon-password"></span><span class="input-text"><%=traduci("pw2")%>&#42;</span></span>
											<input type="password"  class="form-control input-lg" placeholder="<%=traduci("pw2pl")%>" name="password_again">
										</div><!-- End .input-group -->
								   	</div><!-- End .col-md-12 -->
								   </div>
								   	<div class="col-md-12 text-right">
										<input type="hidden" name="btn3" value="si">
										<input type="submit" value="<%=traduci("btn3")%>" class="btn btn-custom-2" id="btn3">
								   </div>
								  </div><!-- End .panel-body -->
								  
								</div><!-- End .panel-collapse -->
							  
							  
							  
							  
							  </div><!-- End .panel -->
							  
        				</form>
        				</div><!-- End .panel-group #checkout -->
        				<div class="xlg-margin"></div><!-- space -->
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
			</div><!-- End .container -->
        
        </section><!-- End #content -->
        
<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
    <!-- AGGIUNTE -->
    <script src="jquery/js/jquery.scrollintoview.min.js"></script>
	<script src="js/Jqueryvalidate/jquery.validate.min.js"></script>
	<script src="js/Jqueryvalidate/localization/messages_it.min.js"></script>
	<script type="text/javascript" src="chk_piva_cf.js"></script>
	
	<script src="Jquery/js/select2/select2.min.js"></script>	        
	<script src="Jquery/js/select2/select2_locale_it.js"></script>
	<script>
	$(function() {
		
		$.validator.addMethod("codice_fiscale", function (value, element) {
			result = true;
			chk=ControllaCF(value);
			if (chk != '' ) { 
				$.validator.messages.codice_fiscale = chk;
				result = false; 
			}
			if (value=="")
			{
				$.validator.messages.codice_fiscale = "Inserire il codice fiscale";
				result = false; 
			}
			return result;
		}, "");
		$.validator.addMethod("partita_iva", function (value, element) {
			result = true;
			chk=ControllaPIVA(value);
			if (chk != '' ) { 
				$.validator.messages.partita_iva = chk;
				result = false; 
			}
			if (($("#azienda").val()!="") && (value=="")){
				$.validator.messages.partita_iva = "Inserire la partita iva, se sei un privato svuota il campo azienda";
				result = false; 
			}
			return result;
		}, "");
			$("#btn2").click(function(){
			$("#email-form").validate({
				rules: {
					email: {required: true,
					        remote:{
		                url: "searcher.asp?check_email=si", //make sure to return true or false with a 200 status code
		                type: "post",
		                dataType: 'json',
		                cache:false,
		                
		                dataFilter: function(data) {
					        var json = JSON.parse(data);
					        if(json.isError == "true") {
					            return "\"" + decodeURIComponent(encodeURIComponent(json.errorMessage)) + "\"";
					        } else {
					            return true;
					        }
					
					    },
		                
		                error: function(xhr, textStatus, errorThrown)
					        {
					            //alert('ajax loading error... xhr:'+xhr+'textStatus:'+textStatus+'errorThrown:'+errorThrown);
					            return false;
					        }
		            }},
					repeat_email: {required: true,
					        equalTo:{
						        param:"#email"
					        }
					        },
					},
					//submitHandler: function() { alert("Submitted!") },
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
			          }
			          else if ( element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
				            error.insertAfter(element.closest('.input-group'));
			        } else {
		
			           // error.insertAfter(element);
			        }
			    }});
					
					
		});
		$("#btn3").click(function(){
			$("#password-form").validate({
				rules: {
						password:{required:true,minlength: 6},
						password_again:{required:true,minlength: 6,equalTo:"#password"}
					},
					//submitHandler: function() { alert("Submitted!") },
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
			          }
			          else if ( element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
				            error.insertAfter(element.closest('.input-group'));
		
			        } else {
		
			           // error.insertAfter(element);
		
			        }
			    }});
					
					
		});
		$("#btn1").click(function(){
			console.log("citta"+$("#citta").val());
			$("#dati-form").validate({
				rules: {
					email: {required: {
				        depends: function(element){
					            return $("#registrati").is(':checked');
					        }},
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
		                
		                error: function(xhr, textStatus, errorThrown)
					        {
					            alert('ajax loading error... xhr:'+xhr+'textStatus:'+textStatus+'errorThrown:'+errorThrown);
					            return false;
					        }
		            }},
					repeat_email: {required: {
				        depends: function(element){
					            return $("#registrati").is(':checked');
					        }},
					        equalTo:{
						        param:"#email",
						        depends: function(element){
					            return $("#registrati").is(':checked');
					        }
					        }},
	
					//Step 2
					nome: {required:true,minlength:3},
					cognome: {required: true,minlength: 3},
					telefono:{required:true,digits: true,minlength: 8},
					cf: { codice_fiscale: true },
					piva: { partita_iva: true },
					indirizzo: {required: true,minlength: 5},
					citta: {required: true,minlength: 3},
					cap: {required: true,number: true,minlength: 5, maxlength: 5},
					provincia: { required: true }					
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
			          }
			          else if ( element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
				            error.insertAfter(element.closest('.input-group'));
			        } else {
		
			           // error.insertAfter(element);
			        }
			    }
			});//$("#checkout-form").validate({
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
			initSelection: function (element, callback) {
				callback({ id: $("#pres_citta").val(), text: $("#pres_citta").val() });
				//console.log("pres_citta: "+$("#pres_citta").val());
			},
			id: function(object) {
				console.log("id function");
				return object.text;
			}
	  	});
	  	$("#citta").on('select2-selecting', function(e) {
			// Access to full data
			//console.log("select2-selecting:"+e.object.id+e.object.text+e.object.provincia+e.object.regione);
			;$("#citta").val(e.object.text);
			$("#cap").val(e.object.cap);
			$("#provincia").val(e.object.id).trigger("change");
			$("#regione").val(e.object.regione);
		
		});
	  	
	});
			
	
	</script>
    </body>
</html>