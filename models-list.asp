<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->
<%
if request("salva_dati")<>"" and session("idadmin")<>"" then
%>
<%
	questofile=replace(questofile,"/","")
	session("salva_dati")=queryeform()
	lingua=request("lingua")
	'Aggiorno voce titolo
   	sql="select  * FROM traduzioni where pagina='"&questofile&"' and chiave='tit10'"
	rs.Open sql, conn, 1, 3
	if rs.eof then
		rs.addNew
		rs("pagina")=questofile
		rs("chiave")="tit10"
	end if
	rs("valore"&lingua)=request("titolo")
	rs.Update
	rs.close
		
	select case lingua
	case ""
		lingua="italiano"
	case "_en"
		lingua="inglese"
		
	end Select
	
	
	add2log "Modificato testo pagina "&questofile&" lingua "&lingua&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]",2

	Response.ContentType = "application/json; charset=utf-8"
	Set Js = jsObject()
	Js("status")="success"
	Js("message")="Dati salvati"
	js.Flush
	response.end
	
	
	

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
				<header class="content-title">
							<h1 class="title editable" id="titolo"><%=traduci("tit10")%> <%if session("idadmin")<>"" then%><a href="pag_adm_marca_modello.asp" title="Modifica elemento" class=""><i class="fa fa-pencil edit-items"></i></a><%end if%></h1>
							<p class="title-desc"><%=traduci("contustxt")%></p>
					</header>
	        	<div class="panel-group custom-accordion" id="collapse">
		        	<%	
			        	set rs_modelli=server.createObject("adodb.recordset")
			        	set rs_marche=conn.execute("select marca_modello.* from marca_modello where idpadre=0 order by nome")
			        	do while not rs_marche.EOF
			        	rs_modelli.open "select marca_modello.* from marca_modello where modello=true and idnonno="&rs_marche("id")&" order by nome",conn,3,3
			        	
			        %>
                                <div class="panel">
                                    <div class="accordion-header">
                                        <div class="accordion-title"><span><%=rs_marche("nome")%></span><span class="badge"><%=rs_modelli.recordcount%></span></div><!-- End .accordion-title -->
                                        <a class="accordion-btn opened" data-toggle="collapse" data-target="#collapse-<%=rs_marche("id")%>"></a>
                                    </div><!-- End .accordion-header -->
                                    
                                    <div id="collapse-<%=rs_marche("id")%>" class="collapse" style="">
                                        <div class="panel-body marcamodello-list">
	                                        <ul>
	                                        <%
									        	do while not rs_modelli.EOF
		                                        %>
                                            <li><a href="category.asp?idmodello=<%=rs_modelli("id")%>"><%=rs_modelli("nomeconmarca")%></a></li>
                                            <%
	                                            rs_modelli.MoveNext
	                                            loop
	                                            
	                                            %>
	                                        </ul>
                                        </div><!-- End .panel-body -->
                                    </div><!-- End .collapse -->
                                </div><!-- End .panel -->
                    <%
	                    rs_modelli.close
	                    rs_marche.MoveNext
	                    loop
	                    
	                    %>           

                </div><!-- End .panel-group -->
				<%if session("idadmin")<>"" then %>
				<div class="row">
					<div class="col-md-12 col-sm-12 col-xs-12 center">
						<div class="xs-margin"></div><!-- space -->
						<button type="button" class="btn btn-success" id="invia">Salva modifiche</button>							
					</div><!-- End .col-md-12 -->
				</div><!-- End .row -->
				<%end if%>

			</div><!-- End .container -->
        
        </section><!-- End #content -->
        
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
    <!-- AGGIUNTE -->
	<%if session("idadmin")<>"" then%>
	
<script language="javascript" type="text/javascript" src="jquery/js/tinymce/tinymce.min.js"></script>
<script type="text/javascript">
	tinymce.init({
		language : 'it',
	    selector: "h1.editable",
	    inline: true,
	    toolbar: "undo redo",
	    menubar: false
	});
	$("#invia").click( function(){
		var lingua="";
		<%if session("lingua")<>"" then %>
		lingua="<%=session("lingua")%>";
		<%end if %>
		$.ajax({
			url     : "<%=questofile%>?salva_dati=si",
			type    : "post",
			dataType: 'json',
			data	: { lingua: lingua, titolo: $('#titolo').html()},
			success: function(data){
				alert(data.message);
	
			}
			,error:function(xhr, textStatus, error){
				      console.log("xhr.statusText:"+xhr.statusText);
				      console.log("xhr.responseText:"+xhr.responseText);
				      console.log("textStatus:"+textStatus);
				      console.log("error:"+error);
					  }
	
			
			
		});
	});
	</script>

	
	
	<%end if%>


    </body>
</html>
<%

%>