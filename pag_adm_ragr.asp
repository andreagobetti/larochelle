<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

%>

<%
'------------------------------------

oper=request.form("oper")
if request.form("elimina")<>"" then
	sql="delete FROM Raggruppamenti WHERE idrag =" & request.form("id") & ";"
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
	sql="select * from raggruppamenti"
	if oper="add" then
		rs.Open sql, conn, 3, 3
		rs.addnew
	else
		sql= sql & " where idrag=" & request.form("id")
		rs.Open sql, conn, 3, 3
	end if
	rs("raggruppamento")=firstup(trim(request.form("raggruppamento")))
	rs.update
	rs.close
	oper="view"
end if
//response.write oper
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
<center>
  <table height="100%" border="0" cellpadding="<%=cellp%>" cellspacing="<%=cells%>" class="tabellacentrale">     <tr align="center" valign="middle" > 
      <td height="30" colspan="3"> <!--#include virtual="/sub_top.asp" -->
 </td>
<tr> 
      <td width="<%=col_1%>" valign="top"> <!-- Colonna SINISTRA INIZIO-->
        <!--#include virtual="/sub_menu.asp" --><!-- Colonna SINISTRA FINE-->
      </td>
      <td width="<%=col_23%>" align="center" valign="top"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tabella1">
          <tr> 
            <td class=titolomenu>&nbsp;Amministra</td>
          </tr>
		  </table>
        <!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font> 
        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" <%if oper<>"view" then%> onSubmit="return Validator(this)"<%end if%> style="margin-top:0px;">
		
        <%
sql="select * from Raggruppamenti "
if oper="view" then
%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="2" class="ui-widget-header">Raggruppamenti</td>
            </tr>
            <tr align="right"> 
              <td colspan="2" style="border-bottom:1px solid;"><a href="<%=questofile%>?oper=new"><strong>Nuovo 
                raggruppamento</strong></a></td>
            </tr>
            <tr> 
              <%if oper="view" then%>
              <td width="10" align="center" nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;Modifica&nbsp;</td>
              <%end if%>
              <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Raggruppamenti</td>
            </tr>
            <%

	sql= sql & "ORDER BY Raggruppamento;"
	set rs=conn.execute(sql)
	do until rs.EOF
%>
            <tr> 
              <td align="center" style="border-bottom:1px solid;"> <input type="radio" name="modifica" value="<%=rs("idrag")%>" onClick="this.form.submit()"> 
              </td>
              <td style="border-bottom:1px solid;"><%=rs("Raggruppamento")%></td>
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
	sql=sql & " where idrag =" & txt
	if request.form("modifica")<>"" then
		set rs=conn.execute(sql)
		testo= "Modifica raggruppamento"
		
	else
		testo= "Nuovo raggruppamento"
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
              <td colspan="3" class="ui-widget-header"><%response.write testo%></td>
            </tr>
            <tr> 
              <td align="right" valign="top">Raggruppamento</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td>  <%if error<>"" then
			  		testo=request.form("Raggruppamento")
				elseif oper="edit" then
					testo=rs("Raggruppamento")
			  	elseif oper="new" then
			  		testo=""
				end if%> <input type="text" name="raggruppamento" value="<%=testo%>" size="40">
              </td>
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
        </form> 

        <!-- Colonna CORPO FINE-->
      </td>
    </tr>
    <tr> 
      <td colspan="2" height="1">&nbsp;</td>
    </tr>
  </table>
</center>
</body>
</html>
