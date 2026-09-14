<!--#include virtual="/setup.asp" -->
<!--#include file="ClasseModificheRS.asp"-->
<%
if session("idadmin") = "" then call login()

%>

<%
'------------------------------------

oper=request.form("oper")
if request.form("elimina")<>"" then
	sql="delete FROM utenti_gradi WHERE id =" & request.form("id") & ";"
	conn.execute(sql)
end if
select case lcase(request("oper"))
	case "annulla"
		oper="view"
	case "new"
		oper="new"
	case "aggiungi"
		oper="add"
	case "modifica"
		oper="update"
	case else
		oper="view"
		if request.form("modifica")<>"" then oper="edit"
end select
if oper="add" or oper="update" then
	'controlli
end if

if oper="add" or oper="update" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from utenti_gradi"
	if oper="add" then
		rs.Open sql, conn, 3, 3
		rs.addnew
	else
		sql= sql & " where id=" & request.form("id")
		rs.Open sql, conn, 3, 3
		set modificheRS= (new ClasseModificheRS)(oper)
		modifichers.leggi(rs)		
	end if
	rs("nome_grado")=trim(request.form("nome_grado"))
	rs("ornamento")=trim(request.form("ornamento"))
	rs("colore")=trim(replace(request.form("colore"),"#",""))
	rs("colore_nominativo")=trim(replace(request.form("colore_nominativo"),"#",""))
	ordine=trim(request.form("ordine"))
	
	if ordine="" then
		set ordine_rs=conn.execute("select max(ordine) from utenti_gradi")
		if ordine_rs.eof then
			ordine=2
		else
			ordine=ordine_rs(0)
			if isnull(ordine) then
				ordine=2
			else
				ordine=ordine+2	
			end if
		end if	
	end if
	rs("ordine")=ordine
	if oper="update" then
		modifiche_rs=modificheRS.confronta(rs)		
		set modificheRS=Nothing
	end if	
	rs.update
	rs.close
	call riordina_sql("utenti_gradi","")
	call add2log("Modificato grado "&trim(request.form("nome_grado"))&" eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&vbcrlf&modifiche_rs,2)
	
	oper="view"
end if
'response.write oper
%>

<!--#include virtual="/sub_head_adm.asp" -->
<link rel="stylesheet" type="text/css" href="Jquery/js/colpick/colpick.css"/>
</head>

<body>
  <div id="wrap">
    <div id="header">
      <%=titolo_top%><%barra=0%><!--#include virtual="/sub_barra_adminsf2.asp" -->
      <div class="ui-widget-header ui-corner-all titolo_admin">
        <a href="<%=questofile%>">Elenco gradi</a>
      </div>

        <!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font> 
        <form name="form1" id="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>"  style="margin-top:0px;">
		
        <%
sql="select * from utenti_gradi  "
if oper="view" then
%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr align="right"> 
              <td colspan="5" style="border-bottom:1px solid;"><a href="<%=questofile%>?oper=new"><strong>Nuovo</strong></a></td>
            </tr>
            <tr> 
              <td width="10" align="center" nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;Modifica&nbsp;</td>
              <td nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">Descrizione</td>
              <td nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">Ornamento</td>
              <td nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">Ordine</td>
              <td nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">Colore</td>
            </tr>
            <%

	sql= sql & "ORDER BY ordine;"
	set rs=conn.execute(sql)
	do until rs.EOF
%>
            <tr> 
              <td align="center" style="border-bottom:1px solid;"> <input type="radio" name="modifica" value="<%=rs("id")%>" onClick="this.form.submit()"> 
              </td>
              <%
	              if rs("colore")<>"" then
		              stile="style=""color:#"&rs("colore")&";"""
	              else
		              stile=""
		          end if
	              
	              
	              
	          %>
              <td nowrap style="border-bottom:1px solid;"><span <%=stile%>><%=rs("nome_grado")%></span></td>
              <td nowrap style="border-bottom:1px solid;"><%=rs("ornamento")%></td>
              <td nowrap style="border-bottom:1px solid;"><%=rs("ordine")%></td>
              <td nowrap style="border-bottom:1px solid;"><%=ucase(rs("colore"))%></td>
            </tr>
            <%
	rs.movenext
	loop
	%>
          </table>
          <%
else
txt=request.form("modifica")
if request.form("id")<>"" then txt=request.form("id")
	sql=sql & " where id =" & txt
	if request.form("modifica")<>"" then
		set rs=conn.execute(sql)
		testo= "Modifica grado"
		
	else
		testo= "Nuovo grado"
	end if
		  %>
          <input type="hidden" name="id" value="<%=txt%>">
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
              <td align="right" valign="top">Nome grado</td>
              <td>  <%if error<>"" then
			  		var=request.form("nome_grado")
				elseif oper="edit" then
					var=rs("nome_grado")
			  	elseif oper="new" then
			  		var=""
				end if%> <input type="text" name="nome_grado" value="<%=var%>" size="40">
              </td>
            </tr>
            <tr> 
              <td align="right" valign="top">Ornamento</td>
              <td>  <%if error<>"" then
			  		var=request.form("ornamento")
				elseif oper="edit" then
					var=rs("ornamento")
			  	elseif oper="new" then
			  		var=""
				end if%> <input type="text" name="ornamento" value="<%=var%>" size="40">
              </td>
            </tr>
            
            <tr> 
              <td align="right" valign="top">Colore</td>
              <td>  <%if error<>"" then
			  		var=request.form("colore")
				elseif oper="edit" then
					var=rs("colore")
					if var<>"" then
						var="#"&trim(replace(var,"#",""))
						colore=var
						stile="style=""background:"&var&";"""
					end if
			  	elseif oper="new" then
			  		var=""
				end if%> <input type="text" name="colore" class="colore" value="<%=var%>" size="40" <%=stile%>>
              </td>
            </tr>
            <tr> 
              <td align="right" valign="top">Colore nominativo</td>
              <td>  <%if error<>"" then
			  		var=request.form("colore_nominativo")
				elseif oper="edit" then
					var=rs("colore_nominativo")
			  	elseif oper="new" then
			  		var=""
				end if%> <input type="text" name="colore_nominativo" class="colore" value="<%=var%>" size="40">
              </td>
            </tr>
            
            
            <tr> 
              <td align="right" valign="top">ordine</td>
              <td>  <%if error<>"" then
			  		val=request.form("ordine")
				elseif oper="edit" then
					val=rs("ordine")
			  	elseif oper="new" then
			  		val=""
				end if%> <input type="text" name="ordine"  value="<%=val%>" size="10">
              </td>
            </tr>
            
            
            <tr>
	            <td colspan="2" align="center">
		            
		                    <%
if oper="view" then
	txt=puls_new
elseif oper="new" or error<>"" then
	txt="Aggiungi"
elseif oper="edit" then
	txt="Modifica"
end if
%>
    <input type="submit" name="oper" value="<%=txt%>">
    <%if oper<>"view" then%>
    <input type="submit" name="oper" value="Annulla" formnovalidate="formnovalidate" >
    <%end if%>
    <%if oper="edit" and num=0 and error="" then%>
    <input type="submit" name="elimina" value="Elimina" formnovalidate="formnovalidate" >
    <%end if%>
	            </td>
            </tr>
          </table>
		<script type="text/javascript" src="Jquery/js/jquery.validate.min.js"></script>
		<script type="text/javascript" src="Jquery/js/jquery.validate.min_it.js"></script>
		<script type="text/javascript" src="Jquery/js/colpick/colpick.js"></script>
		
		<script>
		$('.colore').colpick({
			//layout:'hex',
			submit:true,
			colorScheme:'dark',
			color:'<%=colore%>',
			onSubmit:function(hsb,hex,rgb,el,bySetColor) {
				//$(el).css('border-right','20px solid #'+hex);
				console.log("Valore"+hex);
				$(el).val("#"+hex);
				$(el).css("background","#"+hex);
				// Fill the text box just if the color was set using the picker, and not the colpickSetColor function.
				//if(!bySetColor) $(el).val(hex);
				$(el).colpickHide();
				return true;
			}
		}).keyup(function(){
			$(this).colpickSetColor(this.value);
		});

		$(function() {
			

			$('#form1').validate({
		       ignore: "",
		       ignoreTitle: true,      
		       rules: {
			
					nome_grado:{
				required:true,minlength:3}
				}
				 ,     
       messages: {
		},
	   
	       invalidHandler: function(e, validator){
       
		        }
		   });

			
			
		});
		</script>

	    
	    
<%end if%>
        </form> 
        <%Set rs = Nothing
	        
	      call connclose()  
        %>

        <!-- Colonna CORPO FINE-->
    </div>
    <div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" -->
  </div><%
  rsClose
  %>
</body>
</html>
