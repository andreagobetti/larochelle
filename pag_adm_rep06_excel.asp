<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
if session("idadmin") = "" then call login()

tdebug=false
mese=request("mese")
anno=request("anno")
if mese="" then
	mese="14"
end if
if anno="" then
	anno=year(date())
end if

select case mese
	case "14"
	nomemese="Tutte"
	case "13"
	nomemese="Prossime"
	case else
	nomemese=nome_mese(mese)
	
end select
if not tdebug then
	Response.ContentType = "application/vnd.ms-excel"
else
	Response.ContentType = "text/html"
end if

if not tdebug then Response.AddHeader "content-disposition", "attachment; filename=scadenze_avere_"&nomemese&"_"&anno&".xls"

sql="select utenti.*,scadenze.*, fatture.Nfat, i.azienda FROM utenti inner join utenti_intestazioni i on utenti.idintestazione = i.id INNER JOIN (scadenze INNER JOIN fatture ON scadenze.idfat = fatture.IDFat) ON utenti.iduser = fatture.iduser where importo>0 "
if mese=13 then
	sql=sql & "and scadenze.data>=date() "
elseif mese=14 then
	'sql=sql & "where  scadenze.causale=0 "
else
	sql=sql & "and month(scadenze.data)="&mese &" and year(scadenze.data)="&anno &"  "
end if

'sql=sql & "where scadenze.causale=0 and importo > 0 and month(scadenze.data)>="&damese &" and month(scadenze.data)<="&amese &" and year(scadenze.data)="&anno


sql=sql&" order by scadenze.data"

set rs = conn.execute(sql&ordine)
lngTotalRecords=clng(conn.Execute("Select Found_Rows();")(0).Value)
%>
        <table width="100%" border="1" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
		<%if mese=13 then%>
            <td colspan="8" class="ui-widget-header">Prossime scadenze: <%=lngtotalrecords%> elementi</td>
	<%
elseif mese=14 then %>
            <td colspan="8" class="ui-widget-header">Tutte le scadenze: <%=lngtotalrecords%> elementi</td>
<%else %>
            <td colspan="8" class="ui-widget-header">scadenze mese di <%=nome_mese(mese)%> <%=anno%>: <%=lngtotalrecords%> elementi</td>
<%end if %>

          </tr>
          <tr>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;</td> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Data</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Fatt.</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">ID Cli.</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Nominativo</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Tot. fattura</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Modalit&agrave; pagamento</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Pagato</td>
          </tr>
          <%
totale=0
solo_una_volta=true
Do While Not rs.EOF
	importo=cdbl(rs("importo"))
	if rs("data")<now() then
		colore="bgcolor='#FFCCCC'" 'rosino
	else
		colore=""
	end if
	%>
		  <%if rs("data")>now() and solo_una_volta then%>
		  <tr bgcolor="#FFCCCC">
            <td colspan="8" style="border-bottom:1px solid;"><b>Totale avere scaduto: <%=formatnumber(totale,2)%></b></td>
          </tr>
	  <%
	  solo_una_volta=false
	  end if%>
	
          <tr <%=colore%>>
            <td style="border-bottom:1px solid;"><input name="" type="checkbox" value=""></td> 
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=formatDateTime(rs("data"), vbShortDate)%></td>
            <td style="border-bottom:1px solid;"><%=rs("NFAT")%></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=rs("iduser")%></td>
            <td nowrap style="border-bottom:1px solid;"><%=rs("azienda")%></td>
            <td align="right" class="corpobox" style="border-bottom:1px solid;"><%if importo<>"" then%><%=formatnumber(rs("importo"),2)%><%end if%></td>
            <td style="border-bottom:1px solid;" class="corpobox"><%if rs("causale")=0 then%><%=rs("note_pagamento")%><%else%><%=rs("tipo_pagamento")%><%end if%></td>
            <td style="border-bottom:1px solid;" class="corpobox"><%=rs("tipo_pagamento")%></td>
          </tr>
		  
		  
		  
		  
          <% 	
		  if importo<>"" then totale=totale+importo
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		rs.MoveNext
	Loop
	Set rs = Nothing
	call connclose()
	%>
                  <tr>
            <td colspan="8" style="border-bottom:1px solid;"><b>Totale avere: <%=formatnumber(totale,2)%></b></td>
          </tr>
</table>