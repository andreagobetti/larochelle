<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

mese=request.form("mese")
anno=request.form("anno")
if anno="" then
	mese="0"
	anno=year(date())
else
	anno=request.form("anno")
end if
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
         <%barra=0%>
	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Top 10 articoli</a></div>
        <!-- Div Content INIZIO-->
        <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" >

               <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="4" class="ui-widget-header">Filtro: </td>
          </tr>
          <tr > 
            <td colspan="4" style="border-bottom:1px solid;">
               Anno
              <input type="text" name="anno" value="<%=anno%>">
              <input type="submit" name="Submit" value="Aggiorna" style="FONT: 12px;"></td>
          </tr></table></form>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="2" class="ui-widget-header">Top 10 articoli per mese</td>
          </tr>
          <tr>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Mese</td> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;"> 
              <p>Articoli</p></td>
          </tr>
		<%for n=1 to 12
		
%>
          <tr>
            <td style="border-bottom:1px solid;"><%=nome_mese(n)%></td> 
            <td style="border-bottom:1px solid;"><%
			sql="select prodotti.IDpro, prodotti.codice, prodotti.articolo, sum(ordini_dett.quantita) AS somma, ordini.data  FROM (ordini INNER JOIN ordini_dett ON ordini.idord = ordini_dett.idord) INNER JOIN prodotti ON ordini_dett.idpro = prodotti.IDpro  GROUP BY prodotti.IDpro, prodotti.codice, prodotti.articolo, Month(ordini.Data), ordini.anno "
			sql=sql&"HAVING Month(ordini.data)="&n&" AND ordini.anno="&anno&" "
			sql=sql&"ORDER BY  sum(ordini_dett.quantita) DESC; "
			set rs=conn.execute(sql)
			conteggio=1
			do while not rs.eof
			response.write rs("somma")&" "&rs("codice")&" "&rs("articolo")&"<br>"
			conteggio=conteggio+1
			if conteggio>12 then
				exit do
			end if
			rs.movenext
			loop
			%>
   </td>
          </tr>
          <% 	
next	
set rs = Nothing
call connclose()
%>
         
        </table>
        
        <!-- Div Content FINE-->
</div>
<div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>

