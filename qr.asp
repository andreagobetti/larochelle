<!--#include virtual="/setup.asp" -->
<%
idpro=request.querystring("idpro")
if idpro<>"" and isnumeric(idpro) then
	add2log "utilizzato codice QR per visualizzare "&codiceProdotto(idpro),1
	response.redirect "product.asp?idpro="&idpro
end if
idord=request.querystring("ordine")
if idord<>"" and isnumeric(idord) then
	'add2log "utilizzato codice QR per visualizzare "&codiceProdotto(idpro),1
	response.redirect "pag_adm_ordine_mobile.asp?idord="&idord
end if
idpro=request.querystring("mag")
if idpro<>"" and isnumeric(idpro) and session("idord_mobile")<>"" then
	set rs_articolo=conn.execute ("select prodotti.* from prodotti where idpro="&idpro)
	set rs_ordine=conn.execute ("select ordini.* from ordini where idord="&session("idord_mobile"))
	%>
<!DOCTYPE html>
<html>
<head>
    <title>Ordine <%=rs_ordine("nord")%></title>
    <meta name="viewport" content="width=device-width"/>
</head>
 
<body>
<p><%
	response.write "scarico articolo "&rs_articolo("articolo")
	response.write  " da ordine "&rs_ordine("nord")
	
	%></p>
</body>
</html>
	
	
	<%
	

	response.end
end if
%>