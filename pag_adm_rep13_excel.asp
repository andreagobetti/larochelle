<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
if session("idadmin") = "" then call login()

tdebug=false
mese=request("mese")
if mese="" then
	mese="14"
	anno=year(date())
else
	anno=request.form("anno")
end if

if not tdebug then
	Response.ContentType = "application/vnd.ms-excel"
else
	Response.ContentType = "text/html"
end if
Set metodipagamento = New ClasseMetodipagamento 'Create an Instance of the Class

if not tdebug then Response.AddHeader "content-disposition", "attachment; filename=fatture_"&nome_mese(mese)&"_"&anno&".xls"
	
sql="select fatture.*, scadenze.riba, scadenze.data as scadenze_data, scadenze.causale, scadenze.importo, ordini.Cognome, ordini.Nome, ordini.Azienda, scadenze.causale FROM (fatture INNER JOIN scadenze ON fatture.IDFat = scadenze.idfat) INNER JOIN ordini ON fatture.idord = ordini.idord  "
if mese=13 then
	sql=sql & "where scadenze.data>=date()  "
elseif mese=14 then
	sql=sql & "where  true "
else
	sql=sql & "where month(scadenze.data)="&mese &" and year(scadenze.data)="&anno &"  "
end if
sql=sql&" and scadenze.riba=true "
sql=sql&" order by scadenze.data"

	set rs=conn.execute(sql)
	%>
<table width="100%" border="1" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="8" class="ui-widget-header">scadenze mese di <%=nome_mese(mese)%> <%=anno%>: <%=rs.recordcount%> elementi</td>
          </tr>
          <tr>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;</td>
                        <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Scadenza</td>
 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Fatt.</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Data</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">ID Cli.</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Nominativo</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Tot. fattura</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Modalit&agrave; pagamento</td>
          </tr>
          <%

Do While Not rs.EOF
%>
          <tr>
            <td style="border-bottom:1px solid;"><input name="" type="checkbox" value=""></td> 
			<td style="border-bottom:1px solid;" class="corpobox"><%=rs("scadenze_data")%></td>

            <td style="border-bottom:1px solid;"><%=rs("NFAT")%></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=formatDateTime(rs("data"), vbShortDate)%></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=rs("iduser")%></td>
            <td nowrap style="border-bottom:1px solid;"><%=rs("azienda")%></td>
            <td align="right" class="corpobox" style="border-bottom:1px solid;"><%=formatcurrency(rs("totale_fattura"),2)%></td>
            <td style="border-bottom:1px solid;" class="corpobox"><%=metodipagamento.descrizione(objPagingRS("pagamento"))%></td>
          </tr>    <%
	rs.movenext
	loop
	rs.close
     %>     
</table>