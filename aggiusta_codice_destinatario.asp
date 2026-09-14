<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->

Reimposta codice destinatario nellel fatture<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "fatture",conn,3,3
	do while not rs.EOF
		if rs("pa")=1 then
			
			codicedestinatario=get_codice_destinatario("",conn.execute("select utenti_clienti.FatturaPACodiceDestinatario from utenti_clienti where iduser="&rs("iduser"))(0))
			rs("codicedestinatario")=codicedestinatario
			response.write "Trovato "&codicedestinatario&"<br>"
			rs.update
			
			
		end if

	
	
		rs.MoveNext
	loop
		
	
	
	response.write "ESEGUITO"
end if
call connclose()

%>