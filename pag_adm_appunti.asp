<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->
<!--#include virtual="/paginazione.asp" -->

<%
if session("idadmin") = "" then call login()

%>

<%
'------------------------------------

oper=request.form("oper")
if request.form("elimina")<>"" then
	sql="delete FROM appunti_letture WHERE idappunto =" & request.form("id") & ";"
	conn.execute(sql)
	sql="delete FROM appunti_destinatari WHERE idappunto =" & request.form("id") & ";"
	conn.execute(sql)
	sql="delete FROM appunti WHERE id =" & request.form("id") & ";"
	conn.execute(sql)
end if
select case lcase(request("oper"))
	case "fatto"
		id=request.form("id")
		iduser=request.form("iduser")
		if iduser="-1" then
			'Elimino aapunto_destinatario
			conn.execute("delete from appunti_destinatari where idappunto="&id&" and iduser="&session("iduser"))
			if clng(conn.execute("select count(*) from appunti_destinatari where idappunto="&id)(0))=0 then
				conn.execute("delete FROM appunti WHERE id =" & id & ";")
			end if
			message="Eliminato appunto gestionale"
			if request.form("nuovo")=1 then
		    	session("chache_appunti")=session("chache_appunti")-1
		    	if session("chache_appunti")=0 then session("chache_appunti")=""
		    end if


		else
			conn.execute ("update appunti set fatto=1 where id="&id)
			message="Appunto fatto"
		end if
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		JS("success")=true
		js("message")=message
		js.Flush
		set js=Nothing
		call connclose()
		response.end
	case "letto"
		id=request.form("id")
		'conn.execute ("insert into appunti_letture ( idappunto,iduser) values ( "&id&","&sessioniduser&")")
		conn.execute("UPDATE appunti_destinatari set data_lettura=now() where idappunto="&id&" and iduser="&sessioniduser)
		
		n_msg=cint(session("chache_appunti"))
		n_msg=n_msg-1
		if n_msg=0 then n_msg=""
		session("chache_appunti")=n_msg
		
		Response.ContentType = "application/json; charset=utf-8"
		Response.CodePage = 65001
		Set Js = jsObject()
		JS("success")=true
		js.Flush
		set js=Nothing
		call connclose()
		response.end
	case "annulla"
		oper="list"
	case "new"
		oper="new"
	case "aggiungi"
		oper="add"
	case "modifica"
		oper="update"
	case else
		oper="list"
		if request.form("modifica")<>"" then oper="edit"
end select
if oper="add" or oper="update" then
	'controlli
end if

if oper="add" or oper="update" then
	modifiche_rs=""
	str_arr_tags=request.form("destinatari")
	if str_arr_tags="" then
		pubblico=0
	else
		pubblico=1
	end if
		
	
	
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from appunti"
	if oper="add" then
		rs.Open sql, conn, 3, 3
		rs.addnew
		rs("data")=now()
		rs("iduser")=sessioniduser
	else
		idappunto=request.form("id")
		sql= sql & " where id=" & idappunto
		rs.Open sql, conn, 3, 3
	end if
	rs("testo")=trim(request.form("testo"))
	rs("pubblico")=pubblico
	rs("fatto")=checkbox("fatto")
	rs.update
	if idappunto="" then idappunto=Get_last_id("appunti")
	rs.close
	set rs = Nothing
	'DESTINATARI
	arr_tags = split(str_arr_tags,",")
	iubound=ubound(arr_tags)
	str_arr_db=""
	set rs_settori=conn.execute( "select appunti_destinatari.iduser from appunti_destinatari where idappunto="&idappunto)
	do while not rs_settori.eof
		if str_arr_db<>"" then str_arr_db=str_arr_db&","
		str_arr_db=str_arr_db&rs_settori("iduser")
		rs_settori.movenext
	loop
	set rs_settori=nothing
	arr_db = split(str_arr_db,",")
	
	'trovo record da eliminare
	for i = 0 to ubound(arr_db)
		found = "false"
			for j = 0 to ubound(arr_tags)
			 if (cint(arr_db(i)) = cint(arr_tags(j))) then
			 found = "true"
			 exit for
			 end if
			next
		if found ="false" then
			conn.execute "delete from appunti_destinatari where idappunto="&idappunto&" and iduser="&arr_db(i) 
			modifiche_rs= modifiche_rs&"Eliminato destinatario "&arr_db(i)&"<br>"
		end if
	next
	
	set rs_settori=Server.CreateObject("ADODB.Recordset")
	rs_settori.Open "select appunti_destinatari.* from appunti_destinatari where idappunto="&idappunto, conn, 1, 3
	'trovo record da aggiungere
	strOutput = "("
	for i = 0 to iubound
		found = "false"
			for j = 0 to ubound(arr_db)
			 if (cint(arr_tags(i)) = cint(arr_db(j))) then
				 found = "true"
				 exit for
			 end if
			next
		if found ="false" then
			striduserTo=""
			iduserTo=cint(arr_tags(i))
			 rs_settori.addnew
			 rs_settori("idappunto")=idappunto
			 rs_settori("iduser")=iduserTo
			 rs_settori.update
			 modifiche_rs= modifiche_rs& "aggiunto iduser "&arr_tags(i)&" per idappunto "&idappunto&"<br>"
			 strOutput=strOutput & arr_tags(i) & ","
			 'ARR2temp = arr_modelli(i)
		end if
	next
	rs_settori.close
	set rs_settori=nothing
		
	'DESTINATARI FINE
	
	call add2log("pubblico"&pubblico&modifiche_rs&"str_arr_tags:"&str_arr_tags,0)
	
	if pubblico=1 then
		Application.lock
		Application("time_ultimo_appunto")=now()
		Application.unlock
	end if
	if oper="add" then
		call add2log("Aggiunto appunto "&" da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin] per "&str_arr_tags&vbcrlf&modifiche_rs,0)
	end if
	oper="list"
end if
'response.write oper
%>

<!--#include virtual="/sub_head_adm.asp" -->
<style>
.destinatario_letto {
	color: green;
	
}
	
	
</style>
</head>

<body>
  <div id="wrap">
    <div id="header">
      <%=titolo_top%><%barra=0%><!--#include virtual="/sub_barra_adminsf2.asp" -->
      <div class="ui-widget-header ui-corner-all titolo_admin">
        <a href="<%=questofile%>">Appunti</a>
      </div>

        <!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font> 
        <form name="form1" id="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;">
		
        <%
if oper="list" then
%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr align="right"> 
              <td colspan="4" style="border-bottom:1px solid;"><a href="<%=questofile%>?oper=new"><strong>Nuovo</strong></a></td>
            </tr>
            <tr style="border-bottom:1px solid; background: #E5E5E5;"> 
              <td width="10" align="center" nowrap style="border-bottom:1px solid;">&nbsp;Modifica&nbsp;</td>
              <td nowrap style="border-bottom:1px solid;">Data</td>
              <td  >Testo</td>
              <td>Fatto</td>
            </tr>
            <%
	
	strsql="SELECT distinctrow a.*, d.data_lettura, admin.nominativo FROM appunti a left join (select * from appunti_destinatari where iduser="&sessioniduser&") d on a.id = d.idappunto left join admin on a.iduser = admin.iduser where a.iduser="&sessioniduser&" or d.iduser="&sessioniduser& " order by fatto, data desc"
	iPageSize=20
	
	call paginazione_start(strsql,ipagesize,"access")     

	Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
	
	if objPagingRS("fatto")=1 then 
		colore="bgcolor='#BEBEBE'"
	elseif 	objPagingRS("iduser")=-1 then
		colore="bgcolor='#00ccff'"
	
	else
		colore=""
	end if
	if isnull(objPagingRS("data_lettura")) and objPagingRS("iduser")<>sessioniduser then
		nuovo=1
	else
		nuovo=0
	end if
	
%>
            <tr <%=colore%> id="id_appunto<%=objPagingRS("id")%>" data-nuovo="<%=nuovo%>"> 
              <td align="center" style="border-bottom:1px solid;">
	              <%if objPagingRS("iduser")=sessioniduser then %> <input type="radio" name="modifica" value="<%=objPagingRS("id")%>" onClick="this.form.submit()"> <%end if %>
              </td>
              <td nowrap style="border-bottom:1px solid;"><%=objPagingRS("data")%>
			  <%if nuovo=1 then response.write "<span class=""badge""><a href=""#"" class=""letto"">Nuovo</a></span>"%>
              <%if objPagingRS("pubblico")=0 then %>
              <br><strong>Privato</strong>
              <%end if %>
              
              </td>
              <td style="border-bottom:1px solid;"><%=trim(objPagingRS("testo"))%>
              <% if objPagingRS("iduser")=-1 then 
	               response.write "<br><span style=""font-size: 75%;""><b>Gestionale</b></span>"
	               elseif objPagingRS("iduser")<>sessioniduser then
	               response.write "<br><span style=""font-size: 75%;"">"&objPagingRS("nominativo") &"</span>"

	             end if
	             %>
              </td>
              <td style="border-bottom:1px solid;">
	              <%if objPagingRS("iduser")=-1 then %>
	              <a class="fatto" href="#" data-iduser="<%=objPagingRS("iduser")%>"><span class="ui-icon ui-icon-trash" ></span></a>
	              <%else %>
	              
	              <input type="checkbox" value="<%=objPagingRS("id")%>" class="fatto" data-iduser="<%=objPagingRS("iduser")%>" <%if objPagingRS("fatto")=1 then response.write "checked disabled" %>>
	              <%end if %>
	              </td>
            </tr>
            <%
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	call paginazione_end(4,false)
	%>
    </table>

		<script>
		$(document).ready(function () {
			$( ".letto" ).click(function(e){
			e.preventDefault();
				var id=$(this).closest("tr").attr("id").replace("id_appunto","");
				console.log("id:"+id);
				var obj=$(this);
				$.ajax({
	                url: "pag_adm_appunti.asp",
	                method: "post",
	                //timeout: 5000,
	                cache: false,
	
	                dataType: 'json',
	                data	: {
		                oper: 'letto',
		                id: id
	                },
	                success: function(result) {
						if (result.success){
							decrementaContatore(obj);
							obj.closest("span").remove();
						}
	                },
	                error: function(xhr, textStatus, error) {
	                    toastr.options = {
	                        "positionClass": "toast-bottom-right",
	                        "timeOut": "0",
	                        "extendedTimeOut": "0",
	                        "closeButton": true
	                    };
	                    toastr.error('Errore in aggiornamento appunto letto','',{timeOut: 0});
	                }
	            });
		});

		$( ".fatto" ).click(function(e){
			e.preventDefault();
			  	var obj=this;
				var id=$(this).closest("tr").attr("id").replace("id_appunto","");
				var iduser=$(this).attr("data-iduser");
			  	console.log("id:"+id+" iduser:"+iduser);
  	            $.ajax({
	                url: "pag_adm_appunti.asp",
	                type: "post",
	                //timeout: 5000,
	                cache: false,
	                dataType: 'json',
	                data	: {
		                oper: 'fatto',
		                id: id,
		                iduser: iduser,
		                nuovo: $(obj).closest("tr").attr("data-nuovo")
	                },
	                success: function(result) {
		                console.log(result.message);
						if (result.success){
							if(iduser==-1){
								decrementaContatore(obj);
								$(obj).closest("tr").remove();
							}
							else{
								$(obj).attr("checked","checked");
								$(obj).attr("disabled","disabled");
							}
						}
	                },
	                error: function(xhr, textStatus, error) {
	                    toastr.options = {
	                        "positionClass": "toast-bottom-right",
	                        "timeOut": "0",
	                        "extendedTimeOut": "0",
	                        "closeButton": true
	                    };
	                    toastr.error('Errore in aggiornamento appunto','',{timeOut: 0});
	
	                }
	            });
		  	});
		  	
		  	
		  	
		  	function decrementaContatore(obj){
			  	var attuale=$(obj).closest("tr").attr("data-nuovo");
			  	console.log("Attuale:"+attuale);
			  	if (attuale=="1"){
				  	$(obj).closest("tr").attr("data-nuovo","0");
					var daleggere=$("#appuntidaleggere").html();
					if (daleggere<=1){
						$("#appuntidaleggere").remove();
					}else
					{
						$("#appuntidaleggere").html(daleggere-1);
					}
				}
		  	}
		});
	          
	          
		</script>
          <%
else
id=request.form("modifica")
if request.form("id")<>"" then txt=request.form("id")
	sql="select * from appunti "
	if request.form("modifica")<>"" then
		set rs=conn.execute(sql&"where id =" & id)
		testo= "Modifica appunto"
		oper="edit"
	else
		testo= "Nuovo appunto"
		oper="new"
	end if
		  %>
          <input type="hidden" name="id" value="<%=id%>">

          <%if error<>"" then
			  		descrizione=request.form("descrizione")
			  	elseif oper="new" then
			  		descrizione=""
				end if%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="2" class="ui-widget-header"><%response.write testo%></td>
            </tr>
            <tr>
              <td>Destinatari</td>
              <td>
	              
	     <input type='hidden' name='destinatari' id='destinatari' style="width:500px;" value=""/>
		 <script type="text/javascript">
		<%
		if oper="edit" then
			Dim Js1
			Set Js1 = jsArray()	
			
			set rsm=conn.execute("select appunti_destinatari.iduser,appunti_destinatari.data_lettura, a.nominativo FROM appunti_destinatari LEFT JOIN admin a ON appunti_destinatari.iduser = a.iduser where idappunto="&id)
			do while not rsm.eof
				Set Js1(Null) = jsObject()
				js1(null)("id")=rsm("iduser")
				Js1(null)("text")=nomebreve( rsm("nominativo"),"")
				js1(null)("lettura")=rsm("data_lettura")
				rsm.movenext
			loop
			rsm.close
			set rsm = nothing
			%>
		  var vartags = <%js1.Flush%>;
			<%
		else%>
		  var vartags = '';
		<%
		set js1=nothing
		end if
		%>
		  $(document).ready(function () {
			  
				
			   function repoFormatResult(repo) {
				   var markup="";
				   if(repo.lettura){
					   markup +='<span class="destinatario_letto">'+repo.text+'</span>';
				   }
				   else{
					   markup=repo.text;
				   }
			      return markup;
			   }

			  
		  	$('#destinatari').select2({
		  		tags: true,
		  		multiple: true,
		  		placeholder: 'Seleziona i destinatari, lascia vuoto per una nota privata',
		  		minimumInputLength: 0,
		  		allowClear: true,
		  		initSelection : function (element, callback) {
			        callback(vartags);
			    },
			    formatResult:repoFormatResult,
			    formatSelection: repoFormatResult,
		  		ajax: {
		  			quietMillis: 150,
		  			url: "ajax_function.asp?select2=destinatari_appunti",
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
		  		}
		  		
		  	});
		  	

		  	
		  	$("#destinatari").select2("data",vartags);
		  });
		</script>
         
	              

            <%if oper="edit" then
            
            
            end if %>
            
			</td>
            </tr>

            <tr>
              <td align="right" valign="top">Testo</td>
              <td>
                <%if error<>"" then
			  		var=request.form("testo")
				elseif oper="edit" then
					
					var=rs("testo")
			  	elseif oper="new" then
			  		var=""
				end if%>
				<textarea name="testo" id="testo" class="tinymce"  rows="6" style="width:100%; height: 400px;"><%=var%></textarea>
              </td>
            </tr>
            <tr>
              <td align="right" valign="top">Fatto</td>
              <td>
              <%val=""
				if oper="edit" then
					if rs("fatto") then val=" checked"
				elseif oper="new" then
					val=""
				else
					val=request.form("fatto")
					if val="si" then val=" checked"
				end if
				if val<>" checked" then val=""
				%>
              <input type="checkbox" name="fatto" value="si" title="" <%response.write val%><%if view="view" then response.write " disabled"%>></td>
            </tr>
            <tr>
	            <td align="center" colspan="2">
		            <%
					if oper="view" then
						txt=puls_new
					elseif oper="new" or error<>"" then
						txt="Aggiungi"
					elseif oper="edit" then
						txt="Modifica"
					end if
					%>
					    <input id="submit" type="submit" name="oper" value="<%=txt%>">
					    <%if oper<>"view" then%>
					    <input type="submit" name="oper" value="Annulla" onClick="annulla=true;">
					    <%end if%>
					    <%if oper="edit" and num=0 and error="" then%>
					    <input type="submit" name="elimina" value="Elimina" onClick="annulla=true;">
					    <%end if%>
	            </td>
            </tr>
          </table>
	    <script src="//cdn.jsdelivr.net/tinymce/4.1.2/tinymce.min.js" type="text/javascript"></script>
	    <script src="//cdn.jsdelivr.net/tinymce/4.1.2/jquery.tinymce.min.js" type="text/javascript"></script>
	    <script src="//cdn.jsdelivr.net/g/jquery.dirtyforms@2.0.0(jquery.dirtyforms.min.js+jquery.dirtyforms.helpers.tinymce.min.js)"></script>
          <script>
		  $(document).ready(function () {
		        $('form').dirtyForms();
		        $('textarea.tinymce').tinymce({
		            theme: 'modern',
	            	statusbar: false, 
					menubar : false,
					forced_root_block : "",
					toolbar_items_size: 'small',
					auto_focus : false,
					plugins: [ 
					"advlist autolink lists link image charmap print preview hr anchor pagebreak",
					"searchreplace wordcount visualblocks visualchars code fullscreen",
					"insertdatetime media nonbreaking save table contextmenu directionality",
					"paste textcolor textcolor"
					],
					toolbar1: " undo redo |  fontsizeselect | bold italic forecolor | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent ",
					image_advtab: true ,
					
				    font_size_style_values: "12px,13px,14px,16px,18px,20px"//,

		        });
		//$("#destinatari").select2();
		$("#privato").click(function(){
			if(this.checked){
				$("#destinatari").select2("destroy");
				$("#destinatari").hide();
				
			}
			else
			{
				console.log("mostra");
				$("#destinatari").show();
				$("#destinatari").select2();
				
			}
		});
		        
		  	});

	          
          </script>
          
        <%
	        if oper<>"new" then 
		        set rs = Nothing
		    end if
		        
end if
call connclose()
%>
        </form> 

        <!-- Colonna CORPO FINE-->
    </div>
    <div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" -->
  </div><%
  %>
</body>
</html>

<%

%>
