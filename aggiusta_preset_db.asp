<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Imposta db_ved e db_acq nella tabella magazzino<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then

	call esegui()


	response.write "ESEGUITO"
end if


sub esegui()
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "select * from magazzino",conn,3,3
	
	do while not rs.EOF
	
		idpro=rs("idpro")
		idvara=rs("idvara")
		idvarb=rs("idvarb")
		db_acq=clng(conn.execute ("select count(*) from distinta_base where idpro_acq="&idpro&" and idvara_acq="&idvara&" and idvarb_acq="&idvarb)(0))
		db_ven=clng(conn.execute ("select count(*) from distinta_base where idpro_ven="&idpro&" and idvara_ven="&idvara&" and idvarb_ven="&idvarb)(0))
		if db_acq>0 then db_acq=1
		if db_ven>0 then db_ven=1
			
			
		if db_acq<> rs("db_acq")	then
			response.write "idmag:"&rs("idmag")&" idpro:"&rs("idpro")&" db_acq non corrisponde db_acq:"&rs("db_acq")&" ricalcolo:"&db_acq&"<br>" 	
		end if
		if db_ven<> rs("db_ven")	then
			response.write "idmag:"&rs("idmag")&" idpro:"&rs("idpro")&" db_ven non corrisponde db_acq:"&rs("db_ven")&" ricalcolo:"&db_ven&"<br>"	
		end if

	
	
		rs.MoveNext
	loop
	
	
	
	call connclose()

end sub
%>