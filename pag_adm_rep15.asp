<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/paginazione.asp" -->
<!--#include virtual="/pag_adm_fatture_for_inc.asp" -->
<%
if session("idadmin") = "" then call login()
mese=request("mese")
tutte=request("tutte")
if mese="" then
	mese="14"
end if
anno=request("anno")
if anno="" then
	anno=year(date())
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<!--#include virtual="/sub_head.asp" -->
<script>
function aggiusta_amese()
{
	if (document.getElementById('AMESE').selectedIndex<document.getElementById('DAMESE').selectedIndex){
	document.getElementById('AMESE').selectedIndex = document.getElementById('DAMESE').selectedIndex;
	}
}
function aggiusta_damese()
{
	if (document.getElementById('AMESE').selectedIndex<document.getElementById('DAMESE').selectedIndex){
	document.getElementById('DAMESE').selectedIndex = document.getElementById('AMESE').selectedIndex;
	}
}
</script>
</head>
<body> 
<div id="wrap">
<div id="header">
<%=titolo_top%>
        <!-- Div Content INIZIO-->      <%barra=0%>
	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Scadenze fatture da fornitore</a></div> 
                <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin:0px 0px 0px 0px;" >
               <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr > 
            <td colspan="4" style="border-bottom:1px solid;">Periodo
              <select name="MESE" onChange="this.form.submit()">
				<option value="14" <%if 14=cint(mese) then response.write " selected" %>>Tutte</option>
				<option value="13" <%if 13=cint(mese) then response.write " selected" %>>Prossime</option>
								<%
				for n=1 to nome_mese(0)
				response.write "<option value='"&n&"'"
				if n=cint(mese) then response.write " selected"
				response.write  ">" & nome_mese(n)&"</option>"
				next
				%>
              </select>
               Anno
              <input name="anno" type="text" value="<%=anno%>" size="4">
              <input type="submit" name="Submit" value="Aggiorna" style="FONT: 12px;"> 
              <a href="pag_adm_rep06_excel.asp?mese=<%=mese%>&anno=<%=anno%>">Esporta excel</a>
              </td>
          </tr></table></form>
        <%if true then%>
		<%
			
sql="select fatture_for.*, scadenze_for.*, i.Cognome, i.Nome, i.Azienda FROM (scadenze_for INNER JOIN fatture_for ON scadenze_for.idfat = fatture_for.idfat) INNER JOIN utenti ON fatture_for.idfor = utenti.iduser inner join utenti_intestazioni i on utenti.idintestazione = i.id "
if tutte="" then
	sql=sql & "where pagato=0 "
end if
			
			
			
if mese=13 then
	sql=sql & "and scadenze_for.scadenza>=CURDATE() "
elseif mese=14 then
	'sql=sql & "where  scadenze.causale=0 "
else
	sql=sql & "and month(scadenze_for.scadenza)="&mese &" and year(scadenze_for.scadenza)="&anno &"  "
end if
sql=sql&" order by scadenze_for.scadenza"
call paginazione_start(sql,10000,"access")
%>
        <table width="100%" border="1" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
		<%if mese=13 then%>
            <td colspan="8" class="ui-widget-header">Prossime scadenze: <%=lngtotalrecords%> elementi</td>
	<%
elseif mese=14 then%>
            <td colspan="8" class="ui-widget-header">Tutte le scadenze: <%=lngtotalrecords%> elementi</td>
<%else %>
            <td colspan="8" class="ui-widget-header">Scadenze mese di <%=nome_mese(mese)%>&nbsp;<%=anno%></td>
<%end if %>
          </tr>
          <tr>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;</td> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Data</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Fatt.</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">ID Cli.</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Nominativo</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Tot. fattura</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Conto</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Pagato</td>
          </tr>
          <%
iRecordsShown = 0
totale=0
solo_una_volta=true
Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
	importo=cdbl(objPagingRS("importo"))
	if objPagingRS("pagato") then
		colore="bgcolor='#99FF99'"
	else
		colore="bgcolor='#FFCCCC'"
	end if
	%>
		  <%if objPagingRS("scadenza")>now() and solo_una_volta  then%>
		  <tr bgcolor="#FFCCCC">
            <td colspan="8" style="border-bottom:1px solid;"><b>Totale dare scaduto: <%=formatcurrency(totale,2)%></b></td>
          </tr>
	  <%
	  solo_una_volta=false
	  end if%>
	
			          <tr <%=colore%>>
			            <td style="border-bottom:1px solid;"><input name="" type="checkbox" value=""></td> 
			            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=formatDateTime(objPagingRS("scadenza"), vbShortDate)%></td>
			            <td style="border-bottom:1px solid;"><a href="pag_adm_fatture_for.asp?idfat=<%=objPagingRS("idFAT")%>"><%=objPagingRS("NFAT")%></a></td>
			            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=objPagingRS("idfor")%></td>
			            <td nowrap style="border-bottom:1px solid;"><%=objPagingRS("azienda")%></td>
			            <td align="right" class="corpobox" style="border-bottom:1px solid;"><%=formatcurrency(objPagingRS("importo"),2)%></td>
			            <td style="border-bottom:1px solid;" class="corpobox"><%=conto(objPagingRS("conto"))%></td>
			            <td style="border-bottom:1px solid;" class="corpobox"><%if objPagingRS("pagato") then%>SI<%else%>NO<%end if%></td>
			          </tr>
		  
		  
		  
		  
          <% 	
		  if importo<>"" then totale=totale+importo
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	set objPagingRS = Nothing
	call paginazione_close()
	call connclose()
	%>
                  <tr>
            <td colspan="8" style="border-bottom:1px solid;" align="right"><b>Totale: <%=formatcurrency(totale,2)%></b></td>
          </tr>
</table><%end if%>
        <!-- FINE BLOCCO CENTRALE -->
        <!-- Colonna CORPO FINE-->
       <!-- Div Content FINE-->
</div>
<div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
