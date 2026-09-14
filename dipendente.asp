<!--#include virtual="/setup.asp" -->
<!--#include file="ClasseModificheRS.asp"-->
<%
if session("iduser") = "" then call login()
select2=true
iddip=request("iddip")
if iddip<>"" then
	if not isnumeric(iddip) then
		session("codice_errore")=1
		session("oggetto_errore")=iddip
		session("pagina_errore")=questofile
		call connclose()
		response.redirect ("404.asp")
	end if
	oper="edit"
	set rs=conn.execute ("select * from utenti_dipendenti where iddip="&iddip&" and iduser="&sessioniduser)
	if rs.eof then
		bt1errore="Questo dipendente non esiste"
	end if
	set rs = nothing
end if
if request.form("btn1")<>"" then
	iduser=sessioniduser
	nuovi_dipendenti=clng(conn.execute("select count(*) from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where dipendenti.nuovo=1 and utenti_dipendenti.iduser= "&iduser)(0))

	sql="select dipendenti.* from dipendenti"
	action=request.form("action")
	if action="update" then
		sql=sql&" where iddip="&iddip
		txt_oper="Modificato "
	end if 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	if action="add" then
		txt_oper="Aggiunto "
		rs.addnew
	end if
	if action="update" then
		set modificheRS= (new ClasseModificheRS)(oper)
		modifichers.leggi(rs)		
	end if 
			
	rs("cognome")=ucase(request.form("cognome"))
	rs("nome")=ucase(request.form("nome"))
	rs("matricola")=request.form("matricola")
	'rs("Contatto")=request.form("Contatto")
	if request.form("sesso")<>"" then rs("sesso")=request.form("sesso")
	'rs("grado")=request.form("grado")
	'rs("arma")=request.form("arma")
	'rs("lato_arma")=request.form("lato_arma")&""
	'rs("nota_dipendente")=request.form("nota_dipendente")
	'lingue=request.form("lingue")
	'if request.form("lingue-altro")<>"" then lingue=lingue&","&request.form("lingue-altro")
	'response.write "lingue:"&lingue
	'rs("lingue")=lingue
	rs.update
	rs.close
	iddip=get_last_id("dipendenti")
	conn.execute ("insert into utenti_dipendenti (iddip, iduser) values ("&iddip&","&session("iduserDipendenti")&")")
	'setta spettanze in ordini
	setta_spettanze_in_ordini(iduser)
	
	denominazione_v=get_denominazione(iduser)
	
	call add2log("[utente="&iduser&"]"&denominazione_v&"[/utente] ha aggiunto dipendente "&ucase(request.form("cognome"))&" "&ucase(request.form("nome"))&" iddip:"&iddip,2)
	set rs = nothing
	if action="add" and nuovi_dipendenti=0 then		
		call inserisci_appunto("<a href=""pag_adm_user.asp?iduser="&iduser&"""><b>"&denominazione_v&"</b></a> ha iniziato ad aggiungere nuovi dipendenti",array(1,4))
	end if	
	Response.write "action:"&action
	if action="update" then
		call connclose()
		response.redirect "dipendenti.asp"
	end if
		iddip=""
		oper=""
	btn1=true
end if
%>
<!--#include virtual="/config/header_inc.asp" -->
<%
	
	if oper="edit" then
		titolo="Modifica dipendente"
		action="update"
	else
		
		titolo="Aggiungi dipendente"
		action="add"
	end if		
 %>
		
		
		
        <section id="content">
        	<div id="breadcrumb-container">
        		<div class="container">
					<ul class="breadcrumb">
						<li><a href="/">Home</a></li>
						<li class=""><a href="dipendenti.asp">Elenco dipendenti</a></li>
						<li class="active"><%=titolo%></li>
					</ul>
        		</div>
        	</div>
        	<div class="container">
			<%if btn1 and bt1errore="" then
			
			 %>
		   		<div class="alert alert-success">
                    <strong>OK!</strong>&nbsp; Il dipendente è stato aggiunto.
                </div>
				<p>Aggiungi un altro dipendente oppure torna a <a  class="btn btn-custom-2" href="dipendenti.asp">Elenco dipendenti</a> </p>
			<%end if%>
			<%if bt1errore<>"" then %>
		   		<div class="alert alert-danger">
                    <strong><%=traduci("error")%>!</strong>&nbsp; <%=bt1errore%>
                </div>
			<%end if%>
			
			
			
			
			<%if bt1errore="" then %>
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=titolo%></h1>
							<p class="title-desc">I campi contrassegnoti con &#42; sono obbligatori</p>
						</header>
						<%
						if oper="edit" then
							sql="SELECT * FROM dipendenti where iddip="&iddip
							set rs_dipendente=conn.execute(sql)
							pres_nome=rs_dipendente("nome")
							pres_cognome=rs_dipendente("cognome")
							pres_matricola=rs_dipendente("matricola")
							pres_sesso=rs_dipendente("sesso")
							
							set rs_dipendente = nothing
						end if
						%>
        				<div class="panel-group custom-accordion" id="checkout">
        				<form action="<%=questofile%>" method="post" id="dati-form">
							  <input type="hidden" name="action" value="<%=action%>">
							  <input type="hidden" name="iddip" value="<%=iddip%>">
								  <fieldset>
								   <div class="row">
								   	<div class="col-md-6 col-sm-6 col-xs-12">
								   		
								   		<h2 class="checkout-title"><%=traduci("tit11")%></h2>
								   		
								   		<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text">Cognome&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("cognomepl")%>" name="cognome" value="<%=pres_cognome%>">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-user"></span><span class="input-text">Nome&#42;</span></span>
										<input type="text"  class="form-control input-lg" placeholder="<%=traduci("nomepl")%>" name="nome" value="<%=pres_nome%>">
									</div><!-- End .input-group -->
									<div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-company"></span><span class="input-text">Sesso&#42;</span></span>
										<span class="form-control input-lg no-minwidth">
                                    <input type="radio" name="sesso" value="M" <%if pres_sesso="M" then response.write "checked"%>>Maschio &nbsp;&nbsp;&nbsp;<input type="radio" name="sesso" value="F" <%if pres_sesso="F" then response.write "checked"%>>Femmina
                                    </span>
									</div><!-- End .input-group -->

                                    <div class="input-group">
										<span class="input-group-addon"><span class="input-icon input-icon-company"></span><span class="input-text">Matricola</span></span>
										<input type="text" class="form-control input-lg" placeholder="<%=traduci("aziendapl")%>" name="matricola" id="matricola" value="<%=pres_matricola%>">
									</div><!-- End .input-group -->
									   <div class="col-md-12 text-right">
											<input type="submit" value="<%if oper="edit" then response.write "Modifica" else response.write "Aggiungi" %>" class="btn btn-custom-2" id="btn1" name="btn1">
									   </div>
								
								   	</div><!-- End .col-md-6 -->
								   	
	
								   </div><!-- End .row -->
								  </fieldset>
        				</form>
        				</div><!-- End .panel-group #checkout -->
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
				<%end if %>
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
		
			$("#dati-form").validate({
				rules: {

					nome: {required:true,minlength:3},
					cognome: {required: true,minlength: 3},
					sesso: { required: true }					
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
			
	
	</script>
    </body>
</html>