<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Imposta campo primo in ordini<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	Server.ScriptTimeout=180
	sql="select idord, stato, tipo_documento, count(idord) as conteggio from ordini group by iduser having conteggio=1"
	
	
	set rs= conn.execute (sql)
	do while not rs.eof
		if rs("stato")<6 and rs("tipo_documento")="ordine" then
			response.write rs("idord")&" stato:"&rs("stato")&" tipo documento:"&rs("tipo_documento")&"<br>"
			conn.execute("update ordini set primo=1 where idord="&rs("idord"))
		end if
		rs.movenext
	loop
	rs.close
	set rs = nothing
	response.write "ESEGUITO"
end if
call connclose()

%>