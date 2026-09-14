<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
'Riordina i dettagli ordine
sql="select * from ordini order by idord"
Set rs_or = Server.CreateObject("ADODB.Recordset")

rs_or.Open sql, conn, 3, 3
do while not rs_or.eof
	set agente=conn.execute("select * from utenti where iduser="&rs_or("iduser"))
	if agente("idagente")>0 then
		response.write "idord:"&rs_or("idord")&" idagente: "&agente("idagente")&"<br>"
		rs_or("idagente")=agente("idagente")
		rs_or.update
		End If
		

	rs_or.movenext
loop
rs_or.close




%>