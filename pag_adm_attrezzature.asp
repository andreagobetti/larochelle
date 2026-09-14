<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

%>

<%
'------------------------------------

oper=request.form("oper")
if request.form("elimina")<>"" then
	sql="delete FROM attrezzature WHERE id_attrezzatura =" & request.form("id_attrezzatura") & ";"
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
	sql="select * from attrezzature "
	if oper="add" then
		rs.Open sql, conn, 3, 3
		rs.addnew
	else
		sql= sql & " where id_attrezzatura=" & request.form("id_attrezzatura")
		rs.Open sql, conn, 3, 3
	end if
	
	rs("descrizione_attrezzatura")=firstup(trim(request.form("descrizione_attrezzatura")))
	rs("note_attrezzatura")=firstup(trim(request.form("note_attrezzatura")))
	rs("stato_1")=trim(request.form("stato_1"))
	rs("serial_number")=trim(request.form("serial_number"))
	rs("idord")=trim(request.form("idord"))
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

  if (theForm.descrizione_attrezzatura.value == "")  { 
    alert("Inserisci una descrizione attrezzatura."); 
    theForm.nome_linea.focus(); 
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
sql="select * from attrezzature "
if oper="view" then
%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="2" class="ui-widget-header">Attrezzature</td>
            </tr>
            <tr align="right"> 
              <td colspan="2" style="border-bottom:1px solid;"><a href="<%=questofile%>?oper=new"><strong>Aggiungi attrezzatura</strong></a></td>
            </tr>
            <tr> 
              <%if oper="view" then%>
              <td width="10" align="center" nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;Modifica&nbsp;</td>
              <%end if%>
              <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Attrezzatura</td>
            </tr>
            <%

	sql= sql & "ORDER BY descrizione_attrezzatura;"
	set rs=conn.execute(sql)
	do until rs.EOF
%>
            <tr> 
              <td align="center" style="border-bottom:1px solid;"> <input type="radio" name="modifica" value="<%=rs("id_attrezzatura")%>" onClick="this.form.submit()"> 
              </td>
              <td style="border-bottom:1px solid;"><%=rs("descrizione_attrezzatura")%></td>
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
	sql=sql & " where id_attrezzatura =" & txt
	if request.form("modifica")<>"" then
		set rs=conn.execute(sql)
		testo= "Modifica attrezzatura"
		
	else
		testo= "Nuovo attrezzatura"
	end if
		  %>
          <input type="hidden" name="id_attrezzatura" value="<%=txt%>">
          <%if error<>"" then
			  		descrizione=request.form("descrizione_attrezzatura")
			  	elseif oper="new" then
			  		descrizione=""
				end if%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="3" class="ui-widget-header"><%response.write testo%></td>
            </tr>
            <tr>
              <td align="right" valign="top">Descrizione attrezzatura</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td><%if error<>"" then
			  		testo=request.form("descrizione_attrezzatura")
				elseif oper="edit" then
					testo=rs("descrizione_attrezzatura")
			  	elseif oper="new" then
			  		testo=""
				end if%>
                  <input type="text" name="descrizione_attrezzatura" value="<%=testo%>" size="40">
              </td>
            </tr>
            <tr> 
              <td align="right" valign="top">Note</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td>  <%if error<>"" then
			  		testo=request.form("note_attrezzatura")
				elseif oper="edit" then
					testo=rs("note_attrezzatura")
			  	elseif oper="new" then
			  		testo=""
				end if%> <textarea name="note_attrezzatura" cols="40" rows="4"><%=testo%></textarea>              </td>
            </tr>
                        <tr> 
              <td align="right" valign="top">Numero di serie</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td>  <%if error<>"" then
			  		testo=request.form("serial_number")
				elseif oper="edit" then
					testo=rs("serial_number")
			  	elseif oper="new" then
			  		testo=""
				end if%>
                <input name="serial_number" type="text" value="<%=testo%>" size="40"></td>
            </tr>
                        <tr>
                          <td align="right" valign="top">Stato</td>
                          <td valign="top">&nbsp;</td>
                          <td> <%if error<>"" then
			  		val=request.form("stato_1")
				elseif oper="edit" then
					val=rs("stato_1")
			  	elseif oper="new" then
			  		val=""
				end if%>  <select name="stato_1">
                  <option value="">Selezionare</option>
                  <%
for n=1 to stato1_attrezzatura(0)
response.write "<option value='"&n&"'"
if val=n then response.write " selected "
response.write ">" & stato1_attrezzatura(n)&"</option>"
next
%>
                </select></td>
                        </tr>
                                                <tr> 
              <td align="right" valign="top">Numero ordine</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td>  <%if error<>"" then
			  		testo=request.form("idord")
				elseif oper="edit" then
					testo=rs("idord")
			  	elseif oper="new" then
			  		testo=""
				end if%>
                <input name="idord" type="text" id="idord" value="<%=testo%>" size="40"></td>
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
<%
function stato1_attrezzatura(n)
	stato1_attrezzatura=""
	if  isnumeric(n) then
		select case n
		case 0
			stato1_attrezzatura=8
		case 1
			stato1_attrezzatura="DA ORDINARE"
		case 2
			stato1_attrezzatura="ORDINATE LIBERE"
		case 3
			stato1_attrezzatura="ORDINATE VENDUTE"
		case 4
			stato1_attrezzatura="A magazzino LIBERE"
		case 5
			stato1_attrezzatura="IN USO SALA DEMO"
		case 6
			stato1_attrezzatura="FUORI IN CONTO VISIONE"
		case 7
			stato1_attrezzatura="VENDUTE DA CONSEGNARE"
		case 8
			stato1_attrezzatura="VENDUTE DA PREPARARE"
		case 9
			stato1_attrezzatura="USATO DA RICONDIZIONARE"
		case 10
			stato1_attrezzatura="USATO PRONTO DA VENDERE"
		case 11
			stato1_attrezzatura="USATO VENDUTO DA CONSEGNARE"
		case 12
			stato1_attrezzatura="VENDUTO DA RIPARARE PRESSO CLIENTE"
		case 13
			stato1_attrezzatura="VENDUTO DA RIPARARE IN SEDE"
		case 14
			stato1_attrezzatura="DA PREPARARE PER DEMO O FIERE"
		end select
	end if
end function


%>