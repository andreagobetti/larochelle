<!--#include virtual="/setup.asp" -->
<%   
	dim rs_tag
	Set rs_settori = Server.CreateObject("ADODB.Recordset")
	Set rs_tag = Server.CreateObject("ADODB.Recordset")
		rs_tag.Open "select * from tags", conn, 3, 3
	
call scansionasettoriAvanti(0,247,0)
		rs_tag.Close


sub scansionasettoriAvanti(layer,id,idpadre)
	'declaring
	dim sql_order,str,MessageSpacing,mesid,spaceing,tid
	'create recordset
	set rs_order=server.createObject("adodb.recordset")
	'find child with thread_parent=his parent id
	if escludi="" then escludi=0
	sql_order = "select * FROM settori WHERE idpadre = " & id
	rs_order.open sql_order,conn
	nodo=true
	
	do until rs_order.eof
		rs_tag.addNew
		rs_tag("NomeTag")=rs_order("nome_settore")
		rs_tag("id_tmp")=rs_order("idsettore")
		rs_tag("idpadre_tmp")=rs_order("idpadre")
		rs_tag("idpadre")=idpadre
		rs_tag("idnonno")=0

		rs_tag.Update 
		
		
	    call scansionasettoriAvanti(layer+1,rs_order("idsettore"),rs_tag("idtag"))

		rs_order.movenext
	loop
	'closing object
	rs_order.close
	set rs_order=nothing
End Sub	

%>
