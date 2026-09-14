<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
crea foreink keys<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	
	sql="select * from utenti_foglio_spettanze "
	set rs= conn.execute (sql)
	do while not rs.eof
		conn.execute("update utenti_spettanze set idfoglio="&rs("id")&" where iduser="&rs("iduser"))
		rs.movenext
	loop
	rs.close
	set rs = nothing
	conn.execute("update utenti_foglio_spettanze set creato=curdate()")
	conn.execute("update utenti_foglio_spettanze set creato=completo where creato>completo")
	response.write "ESEGUITO"
end if
call connclose()

%>