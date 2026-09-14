<!--#include virtual="/setup.asp" -->
<%
da=request("da")
a=request("a")
if (a<>"" and da<>"") then
	conn.execute ("delete from traduzioni where pagina='"&a&"'")
	n=1

	rs.Open "select * from traduzioni", conn, 1, 3


	set rs_tr=conn.execute("select * from traduzioni where (pagina='"&da&"') ")
	'session("questofile")=questofile
	'Dim dizionario
	Set dizionario=Server.CreateObject("Scripting.Dictionary")
	do while not rs_tr.eof
		rs.addnew
		rs("pagina")=a
		rs("chiave")=rs_tr("chiave")
		rs("valore_en")=rs_tr("valore_en")
		rs("valore")=rs_tr("valore")

		rs.update
		rs_tr.movenext
		n=n+1
	loop
	response.write "Duplicate "&n& " chiavi"
	
else
	response.write "Manca da e a"
end if
	
	%>