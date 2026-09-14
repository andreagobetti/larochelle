<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

idpro=request("idpro")
id=request("id")
%>

<%
'------------------------------------

oper=request.form("oper")
if request.form("elimina")<>"" then
	sql="delete FROM linea_prodotti WHERE id =" & request.form("id") & ";"
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
	sql="select * from linea_prodotti"
	if oper="add" then
		rs.Open sql, conn, 3, 3
		rs.addnew
	else
		sql= sql & " where id=" & request.form("id")
		rs.Open sql, conn, 3, 3
	end if
	rs("idpro")=request.form("idpro")
	rs("idlinea")=request.form("idlinea")
	rs.update
	rs.close
	oper="view"
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

<!--#include virtual="/sub_head.asp" -->
<% if oper<>"view" then%>
<script Language="JavaScript"> 
annulla=false;
function Validator(theForm) 
{
if (annulla==true)  { 
	return(true);    
  } 

  if (theForm.raggruppamento.value == "")  { 
    alert("Inserisci un nome per il raggruppamento."); 
    theForm.raggruppamento.focus(); 
    return (false); 
  } 
  return (true); 
} 

</script>

<%end if%>
</head>
<body>

        <!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font> 
        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" <%if oper<>"view" then%> onSubmit="return Validator(this)"<%end if%> style="margin-top:0px;">
          
          <%
		
		idpro=request("idpro")
		
sql="select linea.Nome_linea, linea_prodotti.IDpro, linea_prodotti.ID "
sql=sql & "FROM linea INNER JOIN linea_prodotti ON linea.idlinea = linea_prodotti.IDlinea "
sql=sql & "WHERE (((linea_prodotti.IDpro)="&idpro&")) "
sql=sql & "ORDER BY linea.Nome_linea; "
		

if oper="view" then
%><input type="hidden" name="idpro" value="<%=idpro%>">
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="2" class="ui-widget-header"><span style="float:left;">Linea</span><span style="float:right"><a href="<%=questofile%>?oper=new&idpro=<%=idpro%>"><strong>Aggiungi linea</strong></a></span></td>
            </tr>
            <tr> 
              <%if oper="view" then%>
              <td width="10" align="center" nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;Modifica&nbsp;</td>
              <%end if%>
              <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Linea</td>
            </tr>
            <%

	
	set rs=conn.execute(sql)
	do until rs.EOF
%>
            <tr> 
              <td align="center" style="border-bottom:1px solid;"> <input type="radio" name="modifica" value="<%=rs("id")%>" onClick="this.form.submit()">              </td>
              <td style="border-bottom:1px solid;"><%=rs("nome_linea")%></td>
            </tr>
            <%
	rs.movenext
	loop
	%>
          </table>
          <%
else
	idpro=request("idpro")
	if request.form("modifica")<>"" then
		id=request.form("modifica")
		sql="select * from linea_prodotti where id="& id
		set rs=conn.execute(sql)
		testo= "Modifica linea"
		
	else
		testo= "Nuova linea"
	end if
		  %>
          <input type="hidden" name="id" value="<%=id%>">
           <input type="hidden" name="idpro" value="<%=idpro%>">
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
              <td align="right" valign="top">linea</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td>     <%
set rsm=conn.execute("select * from linea  order by nome_linea")
response.write "<select name=""idlinea"" " 
if view="view" then response.write " disabled"
response.write "><option value="""">Specificare</option>"
do while not rsm.eof
response.write "<option value=""" & rsm("idlinea") &""""
if oper="edit" then
	val=rs("idlinea")
else
	val=cint(request.form("linea"))
	
end if

if val=rsm("idlinea") then response.write " selected"
response.write ">" & rsm("nome_linea") & "</option>"
rsm.movenext
loop
response.write "</select>"
rsm.close
%>                 </td>
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
          <input type="submit" name="elimina" value="Elimina">
          <%end if
end if%>
      </form>        <!-- Colonna CORPO FINE-->      </td>

   
</body>
</html>
