<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<%
	
	'call confronta_utenti("1,2")
if session("idadmin") = "" then call login()
if request.form("elencoid")<>"" then
	call confronta_utenti(request.form("elencoid"))
	call connclose()
	response.end
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<!-- Stile: <%=Application("skins")%> -->
<!--#include virtual="/sub_head.asp" -->
<style>
	.confronta{
		color: red;
	}
	</style>
	
</head>
<body> 
<div id="wrap">
<div id="header">
<%=titolo_top%> 
<%barra=0%>	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>" class="Amministra_piccolo">Ricerca utenti duplicati</a><input type="button" id="confronta" value="confronta"></div>
<!-- Colonna CORPO INIZIO-->
 <%
	 
	 
	 
	contatore=0 
	sql="select utenti.email, Count(utenti.email) AS ConteggioDiemail FROM utenti WHERE (((utenti.Fornitore)=False)) GROUP BY utenti.email HAVING (((utenti.email)<>'') AND ((Count(utenti.email))>1));"
set objPagingRS=conn.execute(sql)
%>
    <form id="formiduser">
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="3" class="ui-widget-header">Email duplicate: questi utenti non possono fare il login</td>
          </tr>
          <tr> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;"> 
              Email</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Ripetizioni</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;"></td>
          </tr>
          <%
		  if objPagingRS.EOF then %>
			<tr> 
            <td  style="border-bottom:1px solid;">Nessuna email duplicata</td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"> </td>
          </tr>
		  
		  
		  <% end if
		  
iRecordsShown = 0
Do While  Not objPagingRS.EOF
%>
          <tr> 
            <td width="60%" style="border-bottom:1px solid;"><a href="pag_adm_user.asp?cercain=email&cerca=<%=objPagingRS("email")%>"> <%=objPagingRS("email")%></a></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=objPagingRS("ConteggioDiemail")%></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"> </td>
          </tr>
          <% 	
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	%>
        </table>
        
           <%
		
	
	sql="select utenti_intestazioni.* from utenti_intestazioni inner join (select utenti_intestazioni.cf from utenti_intestazioni inner join utenti on utenti_intestazioni.id = utenti.idintestazione group by cf having LENGTH(utenti_intestazioni.cf)>0 and Count(*)>1) t on utenti_intestazioni.cf = t.cf order by utenti_intestazioni.cf, iduser"
		
set objPagingRS=conn.execute(sql)
%>
    
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="4" class="ui-widget-header">Ricerca  duplicati per Codice Fiscale</td>
          </tr>
          <tr> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Confronta</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;"></td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Azienda</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Nominativo</td>
          </tr>
          <%
iRecordsShown = 0
Do While  Not objPagingRS.EOF
if tmp<>objPagingRS("cf") then
contatore=contatore+1
%>
          <tr> 
            <td  colspan="4" style="border-bottom:1px solid;"><input type="button" id="confronta<%=contatore%>" class="confronta" value="confronta"> Duplicati per CF <b><%=objPagingRS("cf")%></b></td>
          </tr>
<%
end if
tmp=objPagingRS("cf")
%>
          <tr style="padding : 0 1 0 1; border-bottom:1px solid;"> 
		  <td><input type="checkbox" name="confronta" value="<%=objPagingRS("iduser")%>"></td>
            <td ><input type="hidden" class="confronto<%=contatore%>" value="<%=objPagingRS("iduser")%>"><a href="pag_adm_user.asp?iduser=<%=objPagingRS("iduser")%>" target="_blank">Vedi</a></td>
            <td ><%=objPagingRS("azienda")%></td>
            <td > <%=objPagingRS("nome") & " "& objPagingRS("cognome")%></td>
          </tr>
          <% 	
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	%>
        </table>
          <%
		
		
		
sql="select utenti_intestazioni.* from utenti_intestazioni inner join (select utenti_intestazioni.piva from utenti_intestazioni inner join utenti on utenti_intestazioni.id = utenti.idintestazione group by piva having LENGTH(utenti_intestazioni.piva)>0 and Count(*)>1) t on utenti_intestazioni.piva = t.piva order by utenti_intestazioni.piva, iduser"
set objPagingRS=conn.execute(sql)
%>
    
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="4" class="ui-widget-header">Ricerca  duplicati per Partita Iva</td>
          </tr>
          <tr> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Confronta</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;"></td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Azienda</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Nominativo</td>
          </tr>
          <%
iRecordsShown = 0
tmp=""
Do While  Not objPagingRS.EOF
if tmp<>objPagingRS("piva") then
contatore=contatore+1

%>
          <tr> 
            <td  colspan="4" style="border-bottom:1px solid;"><input type="button" id="confronta<%=contatore%>" class="confronta" value="confronta">  Duplicati per PIVA <b><%=objPagingRS("piva")%></b></td>
          </tr>
<%
end if
tmp=objPagingRS("piva")
%>
          <tr style="padding : 0 1 0 1; border-bottom:1px solid;"> 
		  <td><input type="checkbox" name="confronta" value="<%=objPagingRS("iduser")%>"></td>
            <td ><input type="hidden" class="confronto<%=contatore%>" value="<%=objPagingRS("iduser")%>"><a href="pag_adm_user.asp?iduser=<%=objPagingRS("iduser")%>" target="_blank">Vedi</a></td>
            <td ><%=objPagingRS("azienda")%></td>
            <td > <%=objPagingRS("nome") & " "& objPagingRS("cognome")%></td>
          </tr>
          <% 	
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	%>
        </table>
		<%
sql="select utenti_intestazioni.* from utenti_intestazioni inner join (select utenti_intestazioni.azienda from utenti_intestazioni inner join utenti on utenti_intestazioni.id = utenti.idintestazione group by azienda having LENGTH(utenti_intestazioni.azienda)>0 and Count(*)>1) t on utenti_intestazioni.azienda = t.azienda order by utenti_intestazioni.azienda, iduser"
set objPagingRS=conn.execute(sql)
%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="4" class="ui-widget-header">Ricerca  duplicati per azienda</td>
          </tr>
          <tr> 
          <tr> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Confronta</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;"></td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Azienda</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Nominativo</td>
          </tr>
          <%
iRecordsShown = 0
Do While  Not objPagingRS.EOF
if tmp<>ucase(objPagingRS("azienda")) then
contatore=contatore+1
%>
          <tr> 
            <td  colspan="4" style="border-bottom:1px solid;"><input type="button" id="confronta<%=contatore%>" class="confronta" value="confronta"> Duplicati per <b><%=objPagingRS("azienda")%></b></td>
          </tr>
<%
end if
tmp=ucase(objPagingRS("azienda"))
call concatena_stringa(elencoid,",",objPagingRS("iduser"))
%>
          <tr style="padding : 0 1 0 1; border-bottom:1px solid;"> 
		  <td><input type="checkbox" name="confronta" value="<%=objPagingRS("iduser")%>"></td>
            <td ><input type="hidden" class="confronto<%=contatore%>" value="<%=objPagingRS("iduser")%>"> <a href="pag_adm_user.asp?iduser=<%=objPagingRS("iduser")%>" target="_blank">Vedi</a></td>
            <td ><%=objPagingRS("azienda")%></td>
            <td > <%=objPagingRS("nome") & " "& objPagingRS("cognome")%></td>
          </tr>
          <% 	
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	%>
        </table>
		</form>
       <script>
	   	$(function() {
   			   $(".confronta").click(function(e){
   			   		var contatore=$(this).attr("id").replace("confronta","");
   			   		console.log("contatore:"+contatore);
				   var elencoid = "";
					$('.confronto'+contatore).each(function () {
					
						elencoid+=elencoid==""?$(this).val():','+$(this).val();
					});
				console.log("elencoid:"+elencoid);
				var dialog=$("#dialog").dialog();
				dialog.html("Attendi...");
				$.ajax({
					url     : "<%=questofile%>",
					type    : "post",
					cache	: false,
					//dataType: 'json',
					data	: {elencoid:elencoid},
					success: function(data){
						dialog.html(data);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 600,
							height: "auto",
							title: "Confronto clienti"
						}); 
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;
						
						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});					
					}
				});	
			});

		   	
		   	
		   	
		   	
		   	
		   $("#confronta").click(function(e){
	   
				   var elencoid = "";
					$('input[type=checkbox]').each(function () {
					if(this.checked){
					
						elencoid+=elencoid==""?$(this).val():','+$(this).val();
					}
					});
	   
				var dialog=$("#dialog").dialog();
				dialog.html("Attendi...");
				$.ajax({
					url     : "<%=questofile%>",
					type    : "post",
					cache	: false,
					//dataType: 'json',
					data	: {elencoid:elencoid},
					success: function(data){
						dialog.html(data);
						dialog.dialog({
							autoOpen: true,
							modal: true,
							resizable: false,
							width: 600,
							height: "auto",
							title: "Confronto clienti"
						}); 
					}
					,error: function(xhr, textStatus, error){
						toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
						toastr.error('Errore nel caricamento della pagina');
						var txt="";
						txt+="Errore in "+location.href+"<br>";
						txt+="<br>"+xhr.statusText;
						txt+="<br>"+xhr.responseText;
						txt+="<br>textStatus:"+textStatus;
						txt+="<br>error:"+error;
						
						$.ajax({
							url     : "searcher.asp",
							type    : "post",
							data	: "txt_errore="+encodeURIComponent(txt)
						});					
					}
				});	
			});
	   });
	   
	   </script>
         		
        <!-- Div Content FINE-->
		<%
		set objPagingRS = nothing
		call connclose()
		
		%>
</div>
<div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
<%
sub confronta_utenti(stringa_iduser)
	
	
array_iduser=split(stringa_iduser,",")
if not isarray(array_iduser) then
	exit sub
end if
Dim Dict    'The Dict Object
Dim Record    'The Record Object
Dim Field    'The Field Object
Set Dict = Server.CreateObject("Scripting.Dictionary")
n_fields=0
dim nome_campo()
for n = 0 to ubound(array_iduser)
	set rs_utente=conn.execute("select utenti.*, i.nome, i.cognome, i.azienda, i.cf, i.piva, i.indirizzo, i.citta, i.provincia from utenti inner join utenti_intestazioni i on utenti.idintestazione = i.id  where utenti.iduser="&array_iduser(n)&" order by iduser")
	
	
					if n_fields=0 then
						n_fields=rs_utente.fields.count-1
						redim nome_campo(n_fields)
						'memorizzo i nomi dei campi
						for i=0 to n_fields
							nome_campo(i)=lcase(rs_utente.Fields(i).Name)
							
						next
					end if
					
					on error resume next
					for i=0 to n_fields
						
						Dict.Add array_iduser(n)&nome_campo(i),rs_utente(i).value
						if err.number<>0 then
							response.write nome_campo(i)
							response.End
						end if
							
					next
next
	response.write "<table border=""1""  cellpadding=""3"" cellspacing=""0"" class=""tabella1"">"
	'Intestazione

	call riga("ID utente","iduser",dict,array_iduser)
	etichetta="Data registrazione"
	campo="data"
		response.write "<tr><td>"&etichetta&"</td>"
		'Ciclo sui record
		for n = 0 to ubound(array_iduser)
			response.write "<td>"
			response.write formatdatetime(dict.item(array_iduser(n)&campo),2)&"</td>"
		next
		response.write "</tr>"

	etichetta="Data ultima visita"
	campo="lastvisit"
		response.write "<tr><td>"&etichetta&"</td>"
		'Ciclo sui record
		for n = 0 to ubound(array_iduser)
			response.write "<td>"
			response.write formatdatetime(dict.item(array_iduser(n)&campo),2)&"</td>"
		next
	call riga("Cognome","cognome",dict,array_iduser)
	call riga("Nome","nome",dict,array_iduser)
	call riga("Azienda","azienda",dict,array_iduser)
	call riga("Email","email",dict,array_iduser)
	call riga("Indirizzo","indirizzo",dict,array_iduser)
	call riga("Citt&agrave;","citta",dict,array_iduser)
	call riga("Cap","cap",dict,array_iduser)
	call riga("Provincia","provincia",dict,array_iduser)
	call riga("Telefono","telefono",dict,array_iduser)
	call riga("Codice fiscale","cf",dict,array_iduser)
	call riga("Partita iva","piva",dict,array_iduser)
	call riga("Email alla registrazione","1email",dict,array_iduser)
	etichetta="Modalit&agrave; registrazione"
	campo="modalita_registrazione"
		response.write "<tr><td>"&etichetta&"</td>"
		'Ciclo sui record
		for n = 0 to ubound(array_iduser)
			response.write "<td>"
			response.write modalitaRegistrazione(dict.item(array_iduser(n)&campo))&"</td>"
		next
	etichetta="Ordini"
	campo="iduser"
		response.write "<tr><td>"&etichetta&"</td>"
		'Ciclo sui record
		for n = 0 to ubound(array_iduser)
			response.write "<td>"
			
			response.write conn.execute("select count(*) from ordini where iduser="&dict.item(array_iduser(n)&campo))(0)&"</td>"
		next
		
		
		
	etichetta="Scheda cliente"
	campo="iduser"
		response.write "<tr><td>"&etichetta&"</td>"
		'Ciclo sui record
		for n = 0 to ubound(array_iduser)
			response.write "<td>"
			response.write "<a href=""pag_adm_user.asp?iduser="&dict.item(array_iduser(n)&campo)&""" target=""_blank"">Vedi</a></td>"
		next
	
response.write "</table>"
end sub

sub riga( etichetta, campo, byref dict, byref array_iduser)
		response.write "<tr><td>"&etichetta&"</td>"
		'Ciclo sui record
		for n = 0 to ubound(array_iduser)
			response.write "<td>"
			response.write dict.item(array_iduser(n)&campo)&"</td>"
		next
		response.write "</tr>"

	
end sub

%>
