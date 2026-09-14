<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->

Crea i documenti per i ddt su ddt<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	set rs=conn.execute ("ddt")
	Set rs_ordini = Server.CreateObject("ADODB.Recordset")
	rs_ordini.open "ordini",conn,3,3
	do while not rs.EOF
	
		sub_idord=rs("idord").value
		set rs2=conn.execute("select ordini.* from ordini where idord="&sub_idord)
		idintestazione=rs2("idintestazione")
		idconsegna=rs2("idconsegna")
		rs_ordini.addnew
		rs_ordini("tipo_documento")="ddt"
		rs_ordini("data")=rs("data")
		rs_ordini("anno")=year(rs("data"))
		rs_ordini("nord")=rs("nddt")
		rs_ordini("iduser")=rs("iduser")
		rs_ordini("idintestazione")=idintestazione
		rs_ordini("idconsegna")=idconsegna
		
		rs_ordini.update
		idord=Get_last_id("ordini")
		conn.execute("update ddt set idord="&idord &", sub_idord="&sub_idord&" where idddt="&rs("idddt"))
	
	
		rs.MoveNext
	loop
		
	
	
	response.write "ESEGUITO"
end if
call connclose()

%>