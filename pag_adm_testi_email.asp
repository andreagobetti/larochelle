<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

'******************************************************************************************************
dim tipi(7, 1)
tipi(0, 0) = "Email di registrazione"
tipi(0, 1) = "Email di registrazione in cui viene fornito nome utente e password per l'accesso al sito<br>Posizione: Inizio email"
tipi(1, 0) = "Email invio ordine"
tipi(1, 1) = "Intestazione email invio ordine"
tipi(2, 0) = "Email di conferma ordine"
tipi(2, 1) = "Intestazione email di conferma ordine"
tipi(3, 0) = "Email di richiesta preventivo"
tipi(3, 1) = "Intestazione email di richiesta preventivo"

'******************************************************************************************************

'for each item in Request.Form
'  Response.write item & ": " & Request.Form(item) & "<br>"
'next
puls_new="Nuovo contenuto"
'------------------------------------

if request("dove")<>"" then
	if request("dove")="su" then
		dove="-15"
	else
		dove="+15"
	end if
	sql="UPDATE News SET News.ordine = [ordine]"&dove&" WHERE (((News.IDnews)="&request("id")&"));"
	conn.execute sql
	Set rs = Server.CreateObject("ADODB.Recordset")

	call riordina(request.querystring("tipo"))
end if
sqlbase="select * from news "
oper=request.form("oper")
if request.form("elimina")<>"" then
	sql="DELETE  FROM news WHERE idnews=" & request.form("id") 
	conn.execute(sql)
	sql=""
	add2log "Eliminato contenuto idnews"&request.form("id")&" da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]",3
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
	
	if oper="add" then
		rs.Open sqlbase, conn, 3, 3
		rs.addnew
		operazione="Aggiunto"
	else
		sql= sqlbase & " where idnews=" & request.form("id")
		rs.Open sql, conn, 3, 3
		operazione="Modificato"
	end if
	if trim(request.form("data"))<>"" then rs("data")=trim(request.form("data"))
	rs("titolo")=firstup(trim(request.form("titolo")))
	rs("testo")=firstup(trim(request.form("testo")))
	rs("tipo")=trim(request.form("tipo"))
	rs("stato")=trim(request.form("stato"))
	if request.form("ordine")="" then
		rs("ordine")=0
	else
		rs("ordine")=request.form("ordine")
	end if
	rs.update
	add2log operazione&" contenuto idnews"&rs("idnews")&" tipo"&rs("tipo")&" da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]",3
	rs.close
	call riordina (trim(request.form("tipo")))
	oper="view"
	'response.redirect Request.ServerVariables("Script_Name")
end if
sub riordina(tipo)
	sql="select * from news where stato='0' and tipo='" & tipo&"' order by ordine;"
	'response.write sql
	rs.open sql, conn,3,3
	'response.end
	ordine=10
	do while not rs.eof
	rs("ordine")=ordine
	rs.update
	ordine=ordine+10
	rs.movenext
	loop
	rs.close
end sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<!--#include virtual="/sub_head.asp" -->
<script>

function editor(from) { 
  popupWin = window.open('editor/editor.asp?from='+from,'editor','toolbar=no,width=700,height=525,location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=yes,scrollbars=yes');
  popupWin.focus();
}

<%if oper<>"view" then%>
annulla=false;
function Validator(theForm) 
{
if (annulla==true)  { 
	return(true);    
  } 

  if (theForm.titolo.value == "")  { 
    alert("Specificare un titolo per il box.");
	theForm.titolo.focus();
	return (false); 
}

  if (theForm.tipo.selectedIndex == 0)  { 
    alert("Seleziona la posizione nella home page.");
	theForm.tipo.focus(); 
	return (false); 
  }
  
  return (true); 
} 
<%end if%>

</script>
</head>
<body>
<div id="wrap">
<div id="header">
<%=titolo_top%> <!-- Box CORPO INIZIO--><%barra=0%>	<div id="barra_fissa">  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>" class="Amministra_piccolo">testi nelle email</a></div></div>
        <font color="#FF0000"><b><%=error%></b></font> 
		
		<form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;">
        <%
		
if oper="view" then
ordine=request.form("ordine")
cercain=request("cercain")
sql=""
%>
        

          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="4"  class="ui-widget-header">Personalizzazione testi </td>
            </tr>
            <tr> 
              <td colspan="4"  style="border-bottom:1px solid;">Ordina per: 
                <select name="ordine" style="FONT: 12px;" onChange="this.form.submit()">
                  <option value="tipo" <%if ordine="tipo" then response.write " selected"%>>Tipo</option> 
                  <option value="titolo" <%if ordine="titolo" then response.write selected%>>Titolo</option> 
                </select> <%

%>
                Filtro 
                <select name="cercain" style="FONT: 12px;" onChange="this.form.submit()">
                  <option value="0" <%if cercain="0" then response.write "selected"%>>Tutti</option> 
                  <%
			for i=0 to ubound(tipi)%> 
                  <option value="Z<%=i%>" <%if cercain="Z"&i then response.write  selected%>><%=tipi(i, 0)%></option> 
                  <%
			next
	%> 
                </select> </td>
            </tr>
            <tr bgcolor="#E5E5E5"> 
              <td align="center"  style="border-bottom:1px solid;">&nbsp;Edit&nbsp;</td>
              <td align="center" bgcolor="#E5E5E5"  style="border-bottom:1px solid;">&nbsp;Info&nbsp;</td>
              <td  style="border-bottom:1px solid;">&nbsp;Titolo</td>
              <td align="center"  style="border-bottom:1px solid;">Ordine</td>
            </tr>
	  <%for i=0 to ubound(tipi)%> 
			<tr > 
              <td colspan="4" class="titolobox"><%=i%>) <%=tipi(i, 0)%> </td>
            </tr>
			<tr> 
              <td colspan="4" class="corpobox"  style="border-bottom:1px solid;"><%=tipi(i, 1)%> </td>
            </tr>
	                  <%
			'--------sezione cerca
	sql1=""
	sql=""
	if cercain<>"" then
		select case lcase(cercain)
		case "n","a","p","t","b","d","s"
			sql1=" tipo='"&cercain&"'"
			ordine="ordine"
			case "pag"
			sql1=" tipo='p' or tipo='a'"
		case "ha","hb","hd","hs"
			sql1=" tipo='"&cercain&"'"
			ordine="ordine"
		end select
		sql=sql & " where tipo like 'E"&i&"'" 
	else
		sql=sql & " where tipo like 'E"&i&"'" 
	end if
	ordine="ordine"
	if len(ordine)=0 then ordine="datad"	'ordine di default
	if ordine="datad" then 
		sql= sql & " order by data desc"
	elseif ordine="ordine" then 
		sql= sql & " order by ordine"
	elseif ordine="titolo" then 
		sql= sql & " order by titolo"
	elseif ordine="tipo" then 
		sql= sql & " order by tipo"
	else
		sql= sql & " order by " & ordine 
	end if
	sql= sqlbase & sql
	'response.write sql
	set rs=conn.execute(sql)
	do until rs.EOF
%>
            <tr> 
              <td align="center" style="border-bottom:1px solid;"> <input type="radio" name="modifica" value="<%=rs("idnews")%>" onClick="this.form.submit()"> 
              </td>
              <td align="center" style="border-bottom:1px solid;"> <%
			txt="ID: "&RS("idnews")&vbcrlf
			
			txt=txt &"Visto: "&RS("click")&vbcrlf

						
		  %> <img src="images/question_s.gif" style="cursor:hand" alt="<%=txt%>" width="16" height="16" border="0"> 
              </td>
              <td style="border-bottom:1px solid;"><%=rs("titolo")%></td>
              <td align="center" valign="middle" style="border-bottom:1px solid;"><a href="<%=questofile%>?dove=giu&id=<%=rs("idnews")%>&tipo=<%=rs("tipo")%>" title=" Sposta giu' "><img src="images/fr_giu.gif" width="7" height="4" border="0"></a><%=rs("ordine")%><a href="<%=questofile%>?dove=su&id=<%=rs("idnews")%>&tipo=<%=rs("tipo")%>" title=" Sposta su "><img src="images/fr_su.gif" width="7" height="4" border="0"></a></td>
            </tr>
            <%
	rs.movenext
	loop
	%>
            <tr> 
              <td colspan="4" align="right" style="border-bottom: black 1px solid;"><a href="<%=questofile%>?oper=new&tipo=E<%=i%>"><strong>Aggiungi</strong></a></td>
            </tr>
			<tr bgcolor="#E5E5E5"> 
              <td colspan="4" height="10" style="border-left: #E5E5E5 0px solid;"></td>
            </tr>
	<%next
	%>
          </table>
          <%
else
	txt=request.form("modifica")
	if request.form("id")<>"" then txt=request.form("id")
		'response.write "OPER:" & oper
		sql=sqlbase & " where idnews =" & txt
%>        
          <input type="hidden" name="id" value="<%=txt%>">
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="3" class="ui-widget-header"><%
			  tmp=ucase(request.QueryString("tipo"))
			  if tmp<>"" then
			  	tmp=replace (tmp,"E","")
				tmp=cint(tmp)
				end if
		if oper<>"new" then
			set rs=conn.execute(sql)
			tmp=ucase(rs("tipo"))
			if tmp<>"" then
				tmp=trim(replace (tmp,"E",""))
				tmp=cint(tmp)
			end if

			response.write "Modifica testo per: <span class='corpobox'>" & tipi(tmp, 0) &"</span>"
			
		else
			tmp=ucase(request.QueryString("tipo"))
			if tmp<>"" then
				tmp=replace (tmp,"E","")
				tmp=cint(tmp)
			end if
			Response.write "Inserimento testo per: <span class='corpobox'>" & tipi(tmp, 0) &"</span>"
		end if
		  %>
                <input type="hidden" name="cercain" value="<%=request("cercain")%>"></td>
            </tr>
            <tr> 
              <td align="right" valign="top">Titolo</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Titolo" width="16" height="16"></td>
              <td> <%if error<>"" then
			  		testo=request.form("titolo")
				elseif oper="edit" then
					testo=rs("titolo")
			  	elseif oper="new" then
			  		testo=""
				end if%> <input type="text" name="titolo" value="<%=fixquotes(testo)%>" size="40">
              </td>
            </tr>
            <tr> 
              <td align="right" valign="top">Testo</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="" width="16" height="16"></td>
              <td><%if error<>"" then
			  		testo=request.form("testo")
				elseif oper="edit" then
					testo=rs("testo")
			  	elseif oper="new" then
			  		testo=""
				end if
%>
 <textarea name="testo" class="mceEditor" rows="35" style="width:100%"><%=testo%></textarea></td>
            </tr>
              <tr> 
                <td valign="top" class=rigasotto>Tipo</td> 
                <td valign="top" class=rigasotto><img src="images/question_s.gif" style="cursor:hand" alt="Tipo di contenuto, vedi l'help per maggiori informazioni" width="16" height="16"></td> 
                <td class=rigasotto>
				<%if error<>"" then
			  		testo=request.form("tipo")
				elseif oper="edit" then
					testo=lcase(rs("tipo"))
			  	elseif oper="new" then
			  		testo=request.querystring("tipo")
				end if%> 
                  <select name="tipo"> 
                    <option value="" >Seleziona</option> 
                    <%
			for i=0 to ubound(tipi)%> 
                    <option value="E<%=i%>" <%if testo="e"&i then response.write " selected"%>><%=tipi(i, 0)%></option> 
                    <%
			next

	%> 
                  </select> </td> 
              </tr> 
            <tr> 
              <td align="right" valign="top">Ordine:</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="" width="16" height="16"></td>
              <td> <%if error<>"" then
			  		testo=request.form("ordine")
				elseif oper="edit" then
					testo=rs("ordine")
			  	elseif oper="new" then
			  		testo=""
				end if%> <input type="text" name="ordine" value="<%=testo%>" size="4"> 
              </td>
            </tr>
              
          </table>
		  <!--#include file="tinymce4_conf4.asp"-->
        <%
if oper="view" then
	txt=puls_new
elseif oper="new" or error<>"" then
	txt="Aggiungi"
elseif oper="edit" then
	txt="Modifica"
end if
%>
    <input type="submit" name="oper" value="<%=txt%>" onClick="return Validator(this.form)"">
    <%if oper<>"view" then%>
    <input type="submit" name="oper" value="Annulla">
    <%end if%>
    <%if oper="edit" and num=0 and error="" then%>
    <input type="submit" name="elimina" value="Elimina">
    <%end if
end if%>
        </form> 
        <!-- Box CORPO FINE-->
</div>
<div id="footer"></div><!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
<%
rsClose
%>
