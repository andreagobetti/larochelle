<%
'-----------------------------------------------------
'-----------------Pagina: cat_adm08.asp
'---------- Applicazione: CATALOGO
'-----------------------------------------------------
'INSERIMENTO MODIFICA CONTENUTI
'-----------------------------------------------------
%>
<%
if session("idadmin") = "" then call login()
%>
<!--#include file="setup.asp" -->
<%
'******************************************************************************************************
dim tipi(7, 1)
tipi(0, 0) = "PAGINA Identificati: BOX inizio pagina"
tipi(0, 1) = "(T0) Vengono richiesti i dati per l'identificazione utente (prima del form) "
tipi(1, 0) = "TESTO Invito alla registrazione"
tipi(1, 1) = "(T1) In calce all'elenco articoli e ai dettagli"
tipi(2, 0) = "PAGINA Ordina: BOX Note su articoli acquistati"
tipi(2, 1) = "(T2) Descrizione del campo note acquisto prima della conferma ordine"
tipi(3, 0) = "PAGINA Ordine inviato"
tipi(3, 1) = "(T3)Pagina dopo invio ordine"
tipi(4, 0) = "PAGINA Registrazione effettuata"
tipi(4, 1) = "(T4) Pagina dopo registrazione effettuata"
tipi(5, 0) = "BOX Tornato (in homepage)"
tipi(5, 1) = "(T5) Box nella home page visualizzato solo agli utenti identificati sotto le categorie"
tipi(6, 0) = "PAGINA Tutela del consumatore"
tipi(6, 1) = "Pagina tutela del consumatore"
tipi(7, 0) = "PAGINA Condizioni di vendita"
tipi(7, 1) = "Pagina condizioni di vendita"
'******************************************************************************************************

'for each item in Request.Form
'  Response.write item & ": " & Request.Form(item) & "<br>"
'next
puls_new="Nuovo elemento"
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
sql=""
sqlbase="select * from news "
oper=request.form("oper")
if request.form("elimina")<>"" then
	sql="DELETE  idnews FROM news WHERE idnews=" & request.form("id") 
	conn.execute(sql)
	sql=""
end if
select case lcase(request.form("oper"))
	case "annulla"
		oper="view"
	case lcase(puls_new)
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
	else
		sql= sqlbase & " where idnews=" & request.form("id")
		rs.Open sql, conn, 3, 3
	end if
	if trim(request.form("data"))<>"" then rs("data")=trim(request.form("data"))
	rs("titolo")=firstup(trim(request.form("titolo")))
	rs("testata")=firstup(trim(request.form("testata")))
	rs("tipo")=trim(request.form("tipo"))
	rs("stato")=trim(request.form("stato"))
	txt=trim(request.form("nordine"))
	if txt="" then txt=0
	rs("ordine")=txt
	rs.update
	rs.close
	call riordina (trim(request.form("tipo")))
	rsclose
	
	response.redirect Request.ServerVariables("Script_Name")
end if
sub riordina(tipo)
	sql="select * from news where stato='0' and tipo='" & tipo&"' order by ordine;"
	'response.write sql
	rs.open sql, conn,3,3
	'response.end
	riordine=10
	do while not rs.eof
	rs("ordine")=riordine
	rs.update
	riordine=riordine+10
	rs.movenext
	loop
	rs.close
end sub
'response.write oper
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<!--#include file="sub_head.asp" -->
<script>
function leggi_ordine()
{
        var selezione  =  window.document.form1.tipo;
        var index  =  selezione.selectedIndex;
        var sel_testo = selezione.options[index].value; 
		
  window.open('popordine.asp?tipo='+sel_testo,'popordine','toolbar=no,width=400,height=525,location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=yes,scrollbars=yes');

		
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
<center> 
<table border="0" height="100%" cellpadding="<%=cellp%>" cellspacing="<%=cells%>"> 
<tr align="center" valign="middle" > 
  <td height="30" colspan="3"> <!--#include file="sub_top.asp" --> </td> 
<tr> 
  <td width="<%=col_1%>" valign="top"> <!-- Colonna SINISTRA INIZIO--> 
    <!--#include file="sub_menu.asp" --> 
    <!-- Colonna SINISTRA FINE--> </td> 
  <td width="<%=col_23%>" align="center" valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tabella1"> 
      <tr> 
        <td class=titolomenu>&nbsp;Amministra</td> 
      </tr> 
    </table> 
    <!-- Colonna CORPO INIZIO--> 
        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" <%if oper<>"view" then%> onsubmit="return Validator(this)"<%end if%> style="margin-top:0px;">
	  <%
if oper="view" then
ordine=request.form("ordine")
cercain=request("cercain")
sql=""



%> 
          <table border="1" cellpadding="2" cellspacing="0" class="tabella1" > 
            <tr> 
              <td colspan="5" align="center"><table  border="1" cellpadding="2" cellspacing="0" style="font-size:10px;"> 
                  <%
						dim i
			for i=0 to ubound(tipi)%> 
                  <tr> 
                    <td><%=tipi(i, 0)%></td> 
                    <td><%=tipi(i, 1)%></td> 
                  </tr> 
                  <%next%> 
                </table> 
                Ordina per:
                <select name="ordine" style="FONT-size: 12px;" onChange="this.form.submit()"> 
                  <option value="tipo" <%if ordine="tipo" then response.write " selected"%>>Tipo</option> 
                  <option value="titolo" <%if ordine="titolo" then response.write selected%>>Titolo</option> 
                </select> 
                <%
if cercain<>"" then
	sql1=""
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
	sql=sql & " where " & sql1
else
	sql=sql & " where tipo like 'Z%'" 
end if
'ordine di default
if len(ordine)=0 then ordine="datad"

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

%> 
                Filtro
                <select name="cercain" style="FONT-size: 12px;" onChange="this.form.submit()"> 
                  <option value="0" <%if cercain="0" then response.write "selected"%>>Tutti</option> 
                  <%
			for i=0 to ubound(tipi)%> 
                  <option value="Z<%=i%>" <%if cercain="Z"&i then response.write  selected%>><%=tipi(i, 0)%></option> 
                  <%
			next
	%> 
                </select></td> 
            </tr> 
            <tr> 
              <%if oper="view" then%> 
              <td align="center" bgcolor="#E5E5E5">Modifica</td> 
              <td align="center" bgcolor="#E5E5E5">Ordine</td> 
              <%end if%> 
              <td bgcolor="#E5E5E5">&nbsp;Titolo</td> 
              <td align="center" bgcolor="#E5E5E5"> Tipo</td> 
              <td align="center" bgcolor="#E5E5E5">Stato</td> 
            </tr> 
            <%

	sql= sqlbase & sql
	response.write sql
	set rs=conn.execute(sql)
	do until rs.EOF
%> 
            <tr> 
              <td align="center"> <input type="radio" name="modifica" value="<%=rs("idnews")%>" onclick="this.form.submit()"> </td> 
              <td align="center" valign="middle"><a href="<%=questofile%>?dove=giu&id=<%=rs("idnews")%>&tipo=<%=rs("tipo")%>" title=" Sposta giu' "><img src="fr_giu.gif" width="7" height="4" border="0"></a><%=rs("ordine")%><a href="<%=questofile%>?dove=su&id=<%=rs("idnews")%>&tipo=<%=rs("tipo")%>" title=" Sposta su "><img src="fr_su.gif" width="7" height="4" border="0"></a></td> 
              <br> 
              <td><%=rs("titolo")%></td> 
              <td align="center"> <%
			response.write tipi(cint(mid(rs("tipo"),2)),0)
			%></td> 
              <td align="center"><%=rs("stato")%></td> 
            </tr> 
            <%
	rs.movenext
	loop
	%> 
          </table> 











	
        

          <%
else
id=request.form("modifica")
if request.form("id")<>"" then txt=request.form("id")
	'response.write "OPER:" & oper
	sql=sqlbase & " where idnews =" & id
	if oper<>"new" then
	
		set rs=conn.execute(sql)
		txt= "Modifica contenuto: <b>" & rs("titolo") & "</b>"
		
	else
		txt= "Nuovo contenuto"
	end if
		  %> 
		  <font color="#FF0000"><b><%=error%></b></font> 
            <input type="hidden" name="id" value="<%=id%>"> 
            <bR> 
            <table border="0" cellpadding="2" cellspacing="0" class=tabellaadmin> 
              <tr> 
                <td colspan="3" valign="top" class=rigasotto><b><%=txt%></b></td> 
              </tr> 
              <tr> 
                <td valign="top" class=rigasotto>Titolo box </td> 
                <td valign="top" class=rigasotto><img src="question_s.gif" style="cursor:hand" alt="Visualizzato come titolo del riquadro" width="16" height="16"></td> 
                <td class=rigasotto> <%if error<>"" then
			  		testo=request.form("titolo")
				elseif oper="edit" then
					testo=rs("titolo")
			  	elseif oper="new" then
			  		testo=""
				end if%> 
                  <input type="text" name="titolo" value="<%=testo%>" size="40"> </td> 
              </tr> 
              <tr> 
                <td valign="top" class=rigasotto>Testo</td> 
                <td valign="top" class=rigasotto><img src="question_s.gif" style="cursor:hand" alt="Visualizzato come corpo del riquadro" width="16" height="16"></td> 
                <td class=rigasotto> <a href="javascript:editor('form3.testata');">Editor</a><br> 
                  <%if error<>"" then
			  		testo=request.form("testata")
				elseif oper="edit" then
					testo=rs("testata")
			  	elseif oper="new" then
			  		testo=""
				end if
%> 
                  <textarea name="testata" cols="40" rows="10"><%=testo%></textarea></td> 
              </tr> 
              <tr> 
                <td valign="top" class=rigasotto>Tipo</td> 
                <td valign="top" class=rigasotto><img src="question_s.gif" style="cursor:hand" alt="Tipo di contenuto, vedi l'help per maggiori informazioni" width="16" height="16"></td> 
                <td class=rigasotto> <%if error<>"" then
			  		testo=request.form("tipo")
				elseif oper="edit" then
					testo=rs("tipo")
			  	elseif oper="new" then
			  		testo=session(questofile&"cercain")
				end if%> 
                  <select name="tipo"> 
                    <option value="" >Seleziona</option> 
                    <%
			for i=0 to ubound(tipi)%> 
                    <option value="Z<%=i%>" <%if testo="Z"&i then response.write " selected"%>><%=tipi(i, 0)%></option> 
                    <%
			next

	%> 
                  </select> </td> 
              </tr> 
              <tr> 
                <td valign="top" class=rigasotto>Ordine:</td> 
                <td valign="top" class=rigasotto><img src="question_s.gif" style="cursor:hand" alt="Ordine in cui vengono visualizzati i contenuti dello stesso tipo" width="16" height="16"></td> 
                <td class=rigasotto> <%if error<>"" then
			  		testo=request.form("nordine")
				elseif oper="edit" then
					testo=rs("ordine")
			  	elseif oper="new" then
			  		testo=""
				end if%> 
                  <input type="text" name="nordine" value="<%=testo%>" size="4">                </td> 
              </tr> 
              <tr> 
                <td valign="top" class=rigasotto>Visibile</td> 
                <td valign="top" class=rigasotto><img src="question_s.gif" style="cursor:hand" alt="Indica se il contenuto viene visualizzato" width="16" height="16"></td> 
                <td class=rigasotto>
				 <%if error<>"" then
			  		testo=request.form("stato")
				elseif oper="edit" then
					testo=rs("stato")
			  	elseif oper="new" then
			  		testo=""
				end if%>  <select name="stato"> 
                    <option value="0" <%if testo=0 then response.write " selected"%>>Si</option> 
                    <option value="1" <%if testo=1 then response.write " selected"%>>No</option> 
                  </select> </td> 
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
    <input type="submit" name="oper" value="Annulla" onclick="annulla=true;">
    <%end if%>
    <%if oper="edit" and num=0 and error="" then%>
    <input type="submit" name="elimina" value="Elimina">
    <%end if
end if%>
        </form> 
          <!-- Colonna CORPO FINE--> </td> 
      </tr> 
      <tr> 
        <td colspan="2" height="1">&nbsp;</td> 
      </tr> 
    </table> 
    </center> 
</body>
</html>
