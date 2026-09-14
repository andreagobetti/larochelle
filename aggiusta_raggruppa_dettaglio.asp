<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Raggruppa gli articoli in ordini_dett<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	idord=14
	sql="select *,min(iddett) from ordini_dett  where idord="&idord&" group by idpro,idvara,idvarb "
	
	
	set rs= conn.execute (sql)
	do while not rs.eof
	
		response.write "iddett:"&rs("iddett")&" idpro:"&rs("idpro")&" idvara:"&rs("idvara")&" idvarb:"&rs("idvarb")&"<br>"
		set rs2=conn.execute("select * from ordini_dett where idord="&idord&" and idpro="&rs("idpro")&" and idvara="&rs("idvara")&" and idvarb="&rs("idvarb")&" and iddett<>"&rs("iddett"))
		do while not rs2.EOF
		
			response.write rs2("iddett")&": quantita "&rs2("quantita")&"<br>"
			'Incrementa quantita su iddett esistente
			sql="update ordini_dett set quantita=quantita+"&rs2("quantita")&" where iddett="&rs("iddett")
			conn.execute (sql)
			
			'Elimina iddett duplicato
			sql="delete from ordini_dett where iddett="&rs2("iddett")
			conn.execute (sql)
			
			'Modifica spettanza
			sql="update ordini_dett_spettanze set iddett="&rs("iddett")&" where iddett="&rs2("iddett")
			conn.execute (sql)
			
		
			rs2.MoveNext
		loop



		rs.movenext
	loop
	rs.close
	set rs = nothing
	response.write "ESEGUITO"
end if
call connclose()

%>