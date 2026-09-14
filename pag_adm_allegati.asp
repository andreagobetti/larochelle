<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

%>

<%
'------------------------------------

oper=request.form("oper")
if request.form("elimina")<>"" then

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
		if request.form("elimina")<>"" then oper="edit"
		if request.form("aggiungi")<>"" then oper="edit"
end select
if oper="add" or oper="update" then
	'controlli
end if

if oper="add" or oper="update" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from articoli_consigliati"
	if oper="add" then
		rs.Open sql, conn, 3, 3
		rs.addnew
	else
		sql= sql & " where id_pro_consigliati=" & request.form("id_pro_consigliati")
		rs.Open sql, conn, 3, 3
	end if
	rs("nome_ragruppamento")=firstup(trim(request.form("nome_ragruppamento")))
	rs("descrizione_ragruppamento")=firstup(trim(request.form("descrizione_ragruppamento")))

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
<!-- Stile: <%=Application("skins")%> -->

<!--#include virtual="/sub_head.asp" -->
    	<style>
	#articolo { list-style-type: none; margin: 0; padding: 0; width:100%; text-align:center;  }
	#articolo div { margin: 3px 3px 3px 0; padding: 1px; float: left; width: 125px; height: 125px; font-size: 1.5em; text-align: center; font-weight:bold; position: relative;}
	.articolo {  font-size: 0.6em; text-align: center; font-weight: normal; }
	.bottone {  font-size: 9px; }
	.posizionato{
		position: absolute;
		right: 0;/*distanza tra il margine destro di #posizionato e il padding di #contenitore*/
		bottom: 0;/*distanza tra il margine inferiore di #posizionato e il padding di #contenitore*/
	}


	</style>

</head>
<body> 
<div id="wrap">
      <div id="header"> <%=titolo_top%>
      	  <%barra=0%>
          <div id="barra_fissa">
	   <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Allegati</a></div></div>
        <!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font> 
        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" <%if oper<>"view" then%> onSubmit="return Validator(this)"<%end if%> style="margin-top:0px;">
		
        <%
sql="select * from files "
if oper="view" then
%>
<div id="articolo">

            <%

	sql= sql & "ORDER BY uploaddate ;"
	set rs=conn.execute(sql)
	do until rs.EOF
%>
            <div class="ui-widget ui-widget-content ui-state-default ui-corner-all" style="font-size:10px;"> <input type="radio" name="modifica" value="<%=rs("idfiles")%>" onClick="this.form.submit()"> <br /><img src="/images/file_ext/<%=trova_estensione_file(rs("filename"))%>.png"  /><br />
             <%=rs("filename")%><br /><br /><%=allegato_a(rs("cosa"))%>
            </div>
            <%
	rs.movenext
	loop
	rs.close
	%>
          </div>
          <%
else
id_pro_consigliati=request.form("modifica")
if request.form("id")<>"" then id_pro_consigliati=request.form("id")
if request.form("id_pro_consigliati")<>"" then id_pro_consigliati=request.form("id_pro_consigliati")
	sql=sql & " where id_pro_consigliati =" & id_pro_consigliati
	if oper="edit" then
		set rs=conn.execute(sql)
		testo= "Modifica raggruppamento"
		
	else
		testo= "Nuovo raggruppamento"
	end if
		  %>
          <input type="hidden" name="id_pro_consigliati" value="<%=id_pro_consigliati%>">
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
              <td align="right" valign="top">Nome correlazione</td>
              <td valign="top"><img src="question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td><%if error<>"" then
			  		testo=request.form("nome_ragruppamento")
				elseif oper="edit" then
					testo=rs("nome_ragruppamento")
			  	elseif oper="new" then
			  		testo=""
				end if%>
                  <input type="text" name="nome_ragruppamento" value="<%=testo%>" size="40">
              </td>
            </tr>
            <tr> 
              <td align="right" valign="top">Descrizione</td>
              <td valign="top"><img src="question_s.gif" style="cursor:hand" alt="Testo visualizzato nel menu" width="16" height="16"></td>
              <td>  <%if error<>"" then
			  		testo=request.form("descrizione_ragruppamento")
				elseif oper="edit" then
					testo=rs("descrizione_ragruppamento")
			  	elseif oper="new" then
			  		testo=""
				end if%> <textarea name="descrizione_ragruppamento" cols="40" rows="4"><%=testo%></textarea>              </td>
            </tr>
            <tr>
              <td colspan="3" align="center"> <%
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
    <%end if%></td>
            </tr>
          </table>
       <%

end if%>
        </form> 
 <!-- Div Content FINE-->
</div>
<div id="footer"></div><!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
<%
rsClose

function trova_estensione_file(fileName) 

    if InStr(fileName, ".") > 0 then 
        trova_estensione_file = Right(fileName, Len(fileName) - InStrRev(fileName, ".")) 
		trova_estensione_file= left(trova_estensione_file,3)
    else 
        trova_estensione_file = "" 
    end if 

end function 

%>
