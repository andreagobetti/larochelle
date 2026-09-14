<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

%>

<%
'------------------------------------

oper=request.form("oper")
if request.form("elimina")<>"" then
	sql="delete FROM ordini_risposte_predefinite WHERE idrisposta =" & request.form("id") & ";"
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
	sql="select * from ordini_risposte_predefinite"
	if oper="add" then
		rs.Open sql, conn, 3, 3
		rs.addnew
	else
		sql= sql & " where idrisposta=" & request.form("id")
		rs.Open sql, conn, 3, 3
	end if
	rs("nome_risposta")=firstup(trim(request.form("nome_risposta")))
	rs("testo_risposta")=toglivcl(trim(request.form("testo_risposta")))
	rs("tipo")=request.form("tipo")
	rs.update
	rs.close
	oper="view"
end if
'response.write oper
%>
<!DOCTYPE html>

<html>
<head>
  <title><%=application("brwstitle")%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  <!--#include virtual="/sub_head.asp" -->

  <style type="text/css">
.rosso {
    color: #F00;
  }
  </style>
  <style type="text/css">
table.c2 {border-collapse: collapse}
  td.c1 {font-weight: bold}
  </style>
  <% if oper<>"view" then%>
<script Language="JavaScript"> 
annulla=false;
function Validator(theForm) 
{
if (annulla==true)  { 
	return(true);    
  } 

  if (theForm.nome_risposta.value == "")  { 
    alert("Inserisci un nome per la risposta."); 
    theForm.nome_risposta.focus(); 
    return (false); 
  } 
  
  
 if(tinyMCE.get('testo_risposta').getContent()==""){
	 alert("Inserisci il testo dela risposta."); 
	tinyMCE.editors['testo_risposta'].focus();
	return (false); 
	 }
  return (true); 
} 

</script>

<%end if%>

</head>

<body>
  <div id="wrap">
    <div id="header">
      <%=titolo_top%><%barra=0%><!--#include virtual="/sub_barra_adminsf2.asp" -->
      <div class="ui-widget-header ui-corner-all titolo_admin">
        <a href="<%=questofile%>">Risposte predefinite per ordini</a>
      </div>

        <!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font> 
        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" <%if oper<>"view" then%> onSubmit="return Validator(this)"<%end if%> style="margin-top:0px;">
		
        <%
sql="select * from ordini_risposte_predefinite "
if oper="view" then
%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr align="right"> 
              <td colspan="3" style="border-bottom:1px solid;"><a href="<%=questofile%>?oper=new"><strong>Nuova risposta</strong></a></td>
            </tr>
            <tr> 
              <%if oper="view" then%>
              <td width="10" align="center" nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;Modifica&nbsp;</td>
              <%end if%>
              <td nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">Nome risposta </td>
              <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Testo risposta</td>
            </tr>
            <%

	sql= sql & "ORDER BY nome_risposta;"
	set rs=conn.execute(sql)
	do until rs.EOF
%>
            <tr> 
              <td align="center" style="border-bottom:1px solid;"> <input type="radio" name="modifica" value="<%=rs("idrisposta")%>" onClick="this.form.submit()"> 
              </td>
              <td nowrap style="border-bottom:1px solid;"><%=rs("nome_risposta")%></td>
              <td style="border-bottom:1px solid;"><%=rs("testo_risposta")%></td>
            </tr>
            <%
	rs.movenext
	loop
	set rs = Nothing
	%>
          </table>
          <%
else
txt=request.form("modifica")
if request.form("id")<>"" then txt=request.form("id")
	sql=sql & " where idrisposta =" & txt
	if request.form("modifica")<>"" then
		set rs=conn.execute(sql)
		testo= "Modifica lista"
		
	else
		testo= "Nuova lista"
	end if
		  %>
          <input type="hidden" name="id" value="<%=txt%>">
		  			<!--#include virtual="/tinymce4_conf2.asp" -->

          <%if error<>"" then
			  		descrizione=request.form("descrizione")
			  	elseif oper="new" then
			  		descrizione=""
				end if%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="3" class="ui-widget-header"><%response.write testo%></td>
            </tr>
            <tr> 
              <td align="right" valign="top">Nome risposta</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td>  <%if error<>"" then
			  		var=request.form("nome_risposta")
				elseif oper="edit" then
					var=rs("nome_risposta")
			  	elseif oper="new" then
			  		var=""
				end if%> <input type="text" name="nome_risposta" value="<%=var%>" size="40">
              </td>
            </tr>
            <tr>
              <td align="right" valign="top">Testo risposta</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td>
                <%if error<>"" then
			  		var=request.form("testo_risposta")
				elseif oper="edit" then
					
					var=rs("testo_risposta")
			  	elseif oper="new" then
			  		var=""
				end if%>
				<textarea name="testo_risposta" id="testo_risposta" class="mceEditor"  rows="6" style="width:100%; height: 400px;"><%=var%></textarea>
              </td>
            </tr>
            <tr>
              <td align="right" valign="top">Tipo</td>
              <td valign="top">&nbsp;</td>
              <td>
              <%if error<>"" then
			  		var=request.form("tipo")
				elseif oper="edit" then
					var=rs("tipo")
			  	elseif oper="new" then
			  		var=""
				end if%>
              <select name="tipo">
                  <option value="1" <%if var="1" then response.write " selected"%>>Conferma ordine</option>
                  <option value="2" <%if var="2" then response.write " selected"%>>Note spedizione</option>
                </select></td>
            </tr>
          </table>
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
    <input type="submit" name="oper" value="Annulla" onClick="annulla=true;">
    <%end if%>
    <%if oper="edit" and num=0 and error="" then%>
    <input type="submit" name="elimina" value="Elimina" onClick="annulla=true;">
    <%end if
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
