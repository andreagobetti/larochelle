<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
dim raggruppasettori
tdebug=false

if session("idadmin") = "" then call login()

settori=false
if request.querystring("settori")<>"" then
	settori=true
	
end if

if not tdebug then
	Response.ContentType = "application/vnd.ms-excel"
else
	Response.ContentType = "text/html"
end if



	if not tdebug then Response.AddHeader "content-disposition", "attachment; filename=Situazione_magazzino_"&formatdatetime(date(),2)&".xls"
	if Request.QueryString("settori")<>"" then raggruppasettori=true	
%>
<table>
	<tr>
		<td colspan="11">
			<strong>Situazione magazzino al <%=FormatDateTime(date(),2)%>
		</td>
	</tr>
</table>

<table  border="1">
	<tr bgcolor="#C0C0C0">
		<th nowrap="nowrap" >Codice</th>
		<th nowrap="nowrap">Articolo</th>
		<th nowrap="nowrap">Variante 1</th>
		<th nowrap="nowrap">Variante 2</th>
		<th nowrap="nowrap">Prezzo acquisto</th>
		<th nowrap="nowrap">Riordino</th>
		<th nowrap="nowrap">Ultimo aggiornamento </th>
		<th nowrap="nowrap">Giacenza attuale</th>
		<th nowrap="nowrap">Valore attuale</th>
		<th nowrap="nowrap">Giacenza al <%=formatdatetime(Application("data_inventario"),2)%></th>
		<th nowrap="nowrap">Valore al <%=formatdatetime(Application("data_inventario"),2)%></th>
      </tr>

<%
	sub_tot_valore1=0
	sub_tot_valore2=0
	
	if settori then	
		sql="select DISTINCT settori.Nome_Settore, settori.idsettore FROM prodotti INNER JOIN settori ON prodotti.idcat = settori.idsettore "
		set rs_settori=conn.execute(sql)
		do while not rs_settori.EOF
		%>
		<tr bgcolor="#C0C0C0">
			<th nowrap="nowrap" colspan="11" align="left">Settore: <%=rs_settori("nome_settore")%></th>
	      </tr>
		
	<%	
	sql="select prodotti.IDpro, prodotti.codice, prodotti.articolo, varianti_a.variante_a, varianti_b.variante_b, prodotti.costo, magazzino.inventario, magazzino.quantita_riordino AS riordino, magazzino.data_aggiornamento, magazzino.quantita_magazzino, magazzino.db_ven, varianti_a.prezzo_ac_va FROM ((magazzino RIGHT JOIN prodotti ON magazzino.idpro = prodotti.IDpro) LEFT JOIN varianti_a ON magazzino.idvara = varianti_a.IDvara) LEFT JOIN varianti_b ON magazzino.idvarb = varianti_b.IDvarb where db_ven=0 and idcat="&rs_settori("idsettore")
	
	call righe_magazzino(sql)
	
	
	     %>    
		<tr bgcolor="">
			<th nowrap="nowrap" colspan="8" align="left">Totale settore: <%=rs_settori("nome_settore")%></th>
			<th nowrap="nowrap" align="right"><%=formatnumber(sub_tot_valore2,2)%></th>
			<th nowrap="nowrap"></th>
			<th nowrap="nowrap" align="right"><%=formatnumber(sub_tot_valore1,2)%></th>
	      </tr>
		  <tr >
			<th nowrap="nowrap" colspan="11" ></th>
	      </tr>
	<%
		tot_valore1=tot_valore1+sub_tot_valore1
		tot_valore2=tot_valore2+sub_tot_valore2
		rs_settori.MoveNext
		Loop
	else	'No settori
		sql="select prodotti.IDpro, prodotti.codice, prodotti.articolo, varianti_a.variante_a, varianti_b.variante_b, prodotti.costo, magazzino.inventario, magazzino.quantita_riordino AS riordino, magazzino.data_aggiornamento, magazzino.quantita_magazzino, magazzino.db_ven, varianti_a.prezzo_ac_va FROM ((magazzino RIGHT JOIN prodotti ON magazzino.idpro = prodotti.IDpro) LEFT JOIN varianti_a ON magazzino.idvara = varianti_a.IDvara) LEFT JOIN varianti_b ON magazzino.idvarb = varianti_b.IDvarb where db_ven=0 " 

		call righe_magazzino(sql)
		tot_valore1=sub_tot_valore1
		tot_valore2=sub_tot_valore2

		
	end if
	%>     
	<tr bgcolor="#C0C0C0">
		<th nowrap="nowrap" colspan="8" align="left">Totale</th>
		<th nowrap="nowrap"><%=formatnumber(tot_valore2,2)%></th>
		<th nowrap="nowrap"></th>
		<th nowrap="nowrap"><%=formatnumber(tot_valore1,2)%></th>
      </tr>
	 
</table>
<%
call connclose()

sub righe_magazzino(sql)
	ordinaper=request.querystring("ordinaper")
	if ordinaper="codice" then
	
		sql=sql&" order by codice"
	elseif ordinaper="dataaggiornamento" then
		sql=sql&" order by data_aggiornamento"
	end if
	
	
	set rs=conn.execute(sql)
	%>
    <%
	totale=0
	do until rs.EOF
	%>
	<tr>
		<td nowrap="nowrap">'<%=rs("codice")%></td>
		<td nowrap="nowrap"><%=rs("articolo")%></td>
		<td nowrap="nowrap"><%=rs("variante_a")%></td>
		<td nowrap="nowrap"><%=rs("variante_b")%></td>
	 <%
		 prezzo_ac_va=converti_typevar14(rs("prezzo_ac_va"))
	 if prezzo_ac_va>0 then
		costo=converti_typevar14(rs("prezzo_ac_va"))
	else
		costo=converti_typevar14(rs("costo"))
	end if
	 %>
       <td nowrap="nowrap"><%if rs("costo")<>"" then%><%=formatnumber(costo,2)%><%end if%></td>
		<td nowrap="nowrap"><%=rs("riordino")%></td>
		<td nowrap="nowrap"><%=rs("data_aggiornamento")%></td>
		<% 
		if not isnull(rs("quantita_magazzino")) then
			quantita_magazzino=cdbl(rs("quantita_magazzino"))
			if quantita_magazzino<0 then quantita_magazzino=0

		else
			quantita_magazzino=0
		end if
		%>
        <td nowrap="nowrap"><%=quantita_magazzino%></td>
<%
	if isnumeric(costo) and isnumeric(quantita_magazzino) then
		valore2=round(costo*quantita_magazzino,2)
		sub_tot_valore2=sub_tot_valore2+valore2
		valore2=formatnumber(valore2,2)
	else
		valore2="Non calcolabile"
	end if
	%>
        <td nowrap="nowrap"><%=valore2%></td>
  		<td nowrap="nowrap"><%=rs("inventario")%></td>
	<%
	if isnumeric(costo) and isnumeric(rs("inventario")) then
		valore1=round(costo*rs("inventario"),2)
		sub_tot_valore1=sub_tot_valore1+valore1
		valore1=formatnumber(valore1,2)
	else
		valore1="Non calcolabile"
	end if
	%>
        <td nowrap="nowrap"><%=valore1%></td>
	
      </tr>
    <%
	rs.movenext
	loop
	rs.close
end sub


%>
 