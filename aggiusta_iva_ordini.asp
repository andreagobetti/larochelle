<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
sql="select ordini.idord,ordini.iva_ordine,ordini.data, fatture.IDfat, fatture.aliquota_iva FROM (ordini LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDfat;"
Set rs_ordini = Server.CreateObject("ADODB.Recordset")

rs_ordini.Open sql, conn, 3, 3
do while not rs_ordini.eof
	if isnull(rs_ordini("idfat")) then
		rs_ordini("iva_ordine")=iva(rs_ordini("data"))
		else
		rs_ordini("iva_ordine")=rs_ordini("aliquota_iva")
		end if
		rs_ordini.update
	
	
	rs_ordini.update
	rs_ordini.movenext
loop
%>