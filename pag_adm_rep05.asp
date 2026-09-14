<!--#include virtual="/setup.asp" -->
<!--#include virtual="/paginazione.asp" -->
<%
if session("idadmin") = "" then call login()

mese=request.form("mese")
if mese="" then
	mese="0"
	anno=year(date())
else
	anno=request.form("anno")
end if
%>
<!--#include virtual="/sub_head_adm.asp" -->
</head>
<body> 
<div id="wrap">
<div id="header">
<%=titolo_top%>
         <%barra=0%>
	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Scadenze</a></div>

                <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin:0px 0px 0px 0px;" >

               <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="4" class="ui-widget-header">scadenze: </td>
          </tr>
          <tr > 
            <td colspan="4" style="border-bottom:1px solid;">Mese
              <select name="MESE" onChange="this.form.submit()">
                <option value="">Selezionare</option>
                <%
for n=1 to nome_mese(0)
response.write "<option value='"&n&"'"
if n=cint(mese) then response.write " selected"
response.write  ">" & nome_mese(n)&"</option>"
next
%>
              </select>
               Anno
              <input type="text" name="anno" value="<%=anno%>">
              <input type="submit" name="Submit" value="Aggiorna" style="FONT: 12px;"></td>
          </tr></table></form>
        <%if mese<>"0" then%>
		<%

sql="select incassi.*, ordini.*, fatture.* FROM ((incassi INNER JOIN ordini ON incassi.idord = ordini.idord) LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDfat where ordini.stato=6 and ordini.eliminato=0 " 
sql=sql & " and month(incassi.data)="&mese &" and year(incassi.data)="&request.form("anno")
sql=sql&" order by incassi.data"


call paginazione_start(sql,200,"access")
%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr> 
            <td colspan="6" class="ui-widget-header">incassi mese di <%=nome_mese(mese)%> <%=anno%>: <%=lngtotalrecords%> elementi</td>
          </tr>
          <tr>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;</td> 
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;"> 
              <p>&nbsp;Data incasso</p></td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Ord./Fatt.</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Cliente</td>
            <td align="right" bgcolor="#E5E5E5" style="border-bottom:1px solid;">Importo</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Pagamento</td>
          </tr>
          <%
iRecordsShown = 0
Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
%>
          <tr>
            <td style="border-bottom:1px solid;"><input name="" type="checkbox" value=""></td> 
            <td style="border-bottom:1px solid;"><%if objPagingRS("incassi.data")<>"" then%><%=formatDateTime(objPagingRS("incassi.data"), vbShortDate)%><%end if%></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=objPagingRS("incassi.idord")%>/<%=objPagingRS("fatture.idfat")%></td>
            <td style="padding : 0 1 0 1; border-bottom:1px solid;"><%=objPagingRS("azienda")%><br><%=objPagingRS("cognome")%> <%=objPagingRS("nome")%></td>
            <td align="right" nowrap style="border-bottom:1px solid;"> <%=formatcurrency(objPagingRS("importo"),2)%></td>
            <td style="border-bottom:1px solid;" class="corpobox"><%=objPagingRS("tipo_pagamento")%></td>
          </tr>
          <% 	
		iRecordsShown = iRecordsShown + 1 ' Increment the number of records we've shown
		objPagingRS.MoveNext
	Loop
	call paginazione_end(6,true)
	%>
          
        </table><%end if%>
        <!-- FINE BLOCCO CENTRALE -->
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
