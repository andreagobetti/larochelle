<!--#include virtual="/setup.asp" -->
<!--#include virtual="/chk_piva_cf.asp" -->

<%
if session("idadmin") = "" then call login()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<!-- Stile: <%=Application("skins")%> -->

<!--#include virtual="/sub_head.asp" -->
</head>
<body> 
<div id="wrap">
<div id="header">
<%=titolo_top%> 
<%barra=0%>	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>" class="Amministra_piccolo">Anomalie su codice fiscale e partita iva</a></div>
<!-- Colonna CORPO INIZIO-->
		<%
		sql="select utenti.nome,utenti.cognome,utenti.azienda,utenti.iduser,utenti.cf,utenti.piva FROM utenti INNER JOIN ordini ON utenti.iduser = ordini.iduser GROUP BY utenti.nome,utenti.cognome,utenti.azienda,utenti.iduser,utenti.cf,utenti.piva;"
		sql=sql & " " 

set objPagingRS=conn.execute(sql)


%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">

          <tr> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;"> 
              <p>&nbsp;Nominativo</p></td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Errore&nbsp;&nbsp;</td>
          </tr>
          <%
iRecordsShown = 0
Do While  Not objPagingRS.EOF
				chkpiva=ControllaPIVA(trim(objPagingRS("piva")))
				chkcf=controllacf(trim(objPagingRS("cf")))
				if objPagingRS("cf")=objPagingRS("piva") then
					if chkcf="" or chkpiva="" then
						chkcf=""
						chkpiva=""
					end if
				
				end if

if chkcf<>"" or chkpiva<>"" then
%>
          <tr> 
            <td width="60%" style="border-bottom:1px solid;"><a href="pag_adm_user.asp?iduser=<%=objPagingRS("iduser")%>"><%=objPagingRS("nome")%> &nbsp;<%=objPagingRS("cognome")%>&nbsp; <%=objPagingRS("azienda")%></a></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"> <%
							if chkpiva<>"" then response.write objPagingRS("piva")&"<br><span style='color: red;'>"&chkpiva&"</span>"
					if chkcf<>"" then
					if chkpiva<>"" then response.write "<br>"
						response.write objPagingRS("cf")&"<br><span style='color: red;'>"&chkcf&"</span>"
					end if
			
		%></td>
          </tr>
          <% 	
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		end if
		objPagingRS.MoveNext
	Loop
	%>
        </table>
       <!-- Div Content FINE-->
</div>
<div id="footer"></div><!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
<%
rsClose
%>