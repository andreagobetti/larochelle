<!--#include virtual="/setup.asp" -->
<%
sql="select * FROM ordini_dett_esteso_originale where idord="&request("idord")
set rs=conn.execute(sql)
%>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
      <tr>
        <td colspan="6" class="ui-widget-header">Dettaglio ordine:</td>
      </tr>
      <tr>
        <td style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Codice</b></td>
        <td style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Articolo</b></td>
        <td align="center" class=testotabella style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;"><b>Prezzo<br>
          unitario</b></td>
        <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Quantit&agrave;</b></td>
        <td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Prezzo<br>
          totale</b></td>
      </tr>
      <%totord=0
	  do while not rs.eof
			%>
      <tr valign="middle">
        <td align="center" style="border-bottom: 1px solid gray;">
              <b><a href="dettaglio.asp?idpro=<%=RS("idpro")%>" title="Clicca per la scheda prodotto" ><%=RS("codice_ordine")%></a></b>          </td>
        <td style="border-bottom: 1px solid gray;"><%=rs("articolo")%><br>
          <%
					txt=""
					if len(rs("var1"))>0 then
						txt ="&nbsp;" & rs("variante1") & ": " & rs("var1")
					end if
					if len(rs("var2"))>0 then
						txt =txt & "&nbsp;" & rs("variante2") & ": " & rs("var2")
					end if
					response.write txt
			  %></td>
        <td align="right" style="border-bottom: 1px solid gray;"><%
                promozione=false
				if RS("promozione")=True and (RS("prodata")>date() or isnull(RS("prodata"))) then promozione=true
				
				
				
				if promozione then
					prezzoeff=RS("prezzo")*(1-RS("sconto")/100)
					txtprezzo="<s>"&prezzole(RS("prezzo"))&"</s><br><b>" & prezzole(prezzoeff)&"</b>"
				else
					prezzoeff=RS("prezzo")
					txtprezzo=prezzole(prezzoeff)
				end if
            	response.write txtprezzo
				%></td>
        <td align="center" valign="middle" style="border-bottom: 1px solid gray;"><%=rs("um")%><br><%if stato>1 then %><%=rs("quantita")%><%else%>
          <input type="text" name="quantita<%=rs("iddett")%>" size="4" value="<%=rs("quantita")%>"><%end if%></td>
        <td align="right" valign="middle" nowrap style="border-bottom: 1px solid gray;"><%response.write prezzole(prezzoeff*rs("quantita"))%></td>
      </tr>
      <%
            totord=totord+(prezzoeff*rs("quantita"))

            rs.movenext
			loop%>
      <tr valign="middle" align="right" >
        <td colspan="5" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">Totale IVA esclusa: <%=prezzole(totord)%></td></tr>

         </table>