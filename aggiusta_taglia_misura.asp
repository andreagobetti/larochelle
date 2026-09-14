<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Rivalorizza il campo taglia misura in ordini_dett_spettanze<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	Server.ScriptTimeout=180
	sql="select * from dipendenti"
	
	
	set rs= conn.execute (sql)
	do while not rs.eof

		x=aggiorna_misure_in_ordini(rs("iddip"))

		rs.movenext
	loop
	rs.close
	set rs = nothing
	response.write "ESEGUITO"
end if
call connclose()

%>