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

select case oper
case "utenti"
	if not tdebug then Response.AddHeader "content-disposition", "attachment; filename=utenti.xls"
        sql= "select * FROM utenti;"
	set rs=conn.execute(sql)
	%>
    <table  border="1">
      <tr bgcolor="#FF66FF">
    <th >Nominativo</th>
    <th>Azienda</th>
    <th>Citta'</th>
    <th>indirizzo</th>
     <th>Provincia</th>
     <th>Telefono</th>
    <th>Numero visite</th>
    <th>Mailing list</th>
    <th>Email</th>
        <th>Sito suggerito</th>
        <th>Ultima_visita</th>
      </tr>
    <%
	do until rs.EOF
	%>
      <tr>
    <td><%=rs("cognome")%>&nbsp;<%=rs("nome")%></td>
    <td><%=rs("azienda")%></td>
    <td><%=rs("citta")%></td>
     <td><%=rs("indirizzo")%></td>
       <td><%=rs("provincia")%></td>
    <td><%=rs("telefono")%></td>
    <td><%=rs("numvisit")%></td>
    <td><%=rs("mailing")%></td>
    <td><%=rs("email")%></td>
        <td><%=rs("sito_suggerito")%></td>
        <td><%if rs("lastvisit")<>"" then response.write formatDateTime(rs("lastvisit"), vbShortDate)%></td>
      </tr>
    <%
	rs.movenext
	loop
	rs.close
case "articoli"
	if not tdebug then Response.AddHeader "content-disposition", "attachment; filename=articoli.xls"
        sql= "select * FROM prodotti order by codice;"
	set rs=conn.execute(sql)
	%>
    <table  border="1">
      <tr bgcolor="#FF66FF">
    <th >Codice</th>
    <th>Articolo</th>
    <th>Descrizione</th>
    <th>Prezzo</th>
    <th></th>
     <th></th>
     <th></th>
    <th></th>
    <th></th>
    <th></th>
        <th></th>
        <th></th>
      </tr>
    <%
	do until rs.EOF
	%>
      <tr>
    <td><%=rs("codice")%></td>
    <td><%=rs("articolo")%></td>
    <td><%=rs("descrizione")%></td>
    <td><%=rs("prezzo")%></td>
     <td></td>
       <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
        <td></td>
        <td></td>
      </tr>
    <%
	rs.movenext
	loop
	rs.close
end select
     %>     
</table>
