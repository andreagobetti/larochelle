<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

oper=request.querystring ("oper")
tdebug=false
if not tdebug then
	Response.ContentType = "application/vnd.ms-excel"
else
	Response.ContentType = "text/html"
end if

	if not tdebug then Response.AddHeader "content-disposition", "attachment; filename=articoli_"&replace(formatdatetime(date,2),"/","_")&".xls"
	where=""
    sql= "select prodotti.*, settori.nome_settore FROM prodotti left join settori on prodotti.idcat = settori.idsettore "
    mostra_settore=true
	if request.querystring("idset")<>"" then
		if request.querystring("idset")>0 then
			mostra_settore=false
			call concatena_stringa( where," and ","idcat="&request.querystring("idset"))
			%>
			Articoli settore <b><%=conn.execute("select nome_settore from settori where idsettore="&request.querystring("idset"))(0)%>
			<%
		end if
	end if
	if where<>"" then
		sql=sql & " where " & where
	end if
		
		
    sql= sql&" order by codice;"
	set rs=conn.execute(sql)
	if not rs.eof then
	%>
	
    <table  border="1">
      <tr bgcolor="#D3D3D3">
	    <th >ID</th>
	    <th >Codice</th>
	    <th>Articolo</th>
	    <th>Prezzo vendita</th>
	    <th>Data modifica</th>
	    <th>Prezzo acquisto</th>
	    <th>Data modifica</th>
	    <%if mostra_settore then %>
	    <th>Settore principale</th>
	    <%end if %>
      </tr>
    <%
	do until rs.EOF
	%>
	<tr>
	    <td><%=rs("idpro")%></td>
	    <td><%=rs("codice")%></td>
	    <td><%=rs("articolo")%></td>
	    <td><%=formatnumber(rs("prezzo"),2,,,0)%></td>
	    <td><% if not isnull(rs("data_agg_prezzo")) then response.write formatdatetime(rs("data_agg_prezzo"),2)%></td>
	    <td><%=formatnumber(rs("costo"),2,,,0)%></td>
	    <td><% if not isnull(rs("data_agg_costo")) then response.write formatdatetime(rs("data_agg_costo"),2)%></td>
	    <%if mostra_settore then %>
	    <td><%=rs("nome_settore")%></td>
		<%end if %>
	</tr>
    <%
	rs.movenext
	loop
	rs.close
     %>     
</table>

<%end if %>
