<!--#include virtual="/setup.asp" -->
<%   
	dim rs_tag
	Set rs_tag = Server.CreateObject("ADODB.Recordset")
	set rs_settori=conn.Execute("select tags.IdTag, settori_prodotti.IDpro, tags.id_tmp FROM tags INNER JOIN settori_prodotti ON tags.id_tmp = settori_prodotti.IDsettore;")
	rs_tag.Open "select * from prodotti_tags", conn, 3, 3	
	do while not rs_settori.EOF
	rs_tag.addNew
	rs_tag("idpro")=rs_settori("idpro")
	rs_tag("idtag")=rs_settori("idtag")
	rs_tag.update
	
	
	rs_settori.MoveNext
	Loop
%>



