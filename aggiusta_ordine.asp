<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
'Riordina i dettagli ordine
sql="select * from ordini order by idord"
Set rs_or = Server.CreateObject("ADODB.Recordset")

rs_or.Open sql, conn, 3, 3
do while not rs_or.eof
	call riordina(rs_or("idord"))
	response.write (rs_or("idord")&"<br>")
	rs_or.movenext
loop
rs_or.close



	sub riordina(idord)
		sql="select * from ordini_dett where idord="&idord&" order by ordine, iddett"
		rs.Open sql, conn, 3, 3
		i=0
		do while not rs.eof
		rs("ordine")=i
		rs.update
		rs.movenext
		i=i+2
		loop
		rs.close
	end sub

%>