<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Imposta il campo verde della tabella ordini<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	sql="select * from ddt"
	
	Set ccausale_ddt = new cl_causale_ddt
	
	set rs= conn.execute (sql)
	do while not rs.eof
		verde=ccausale_ddt.verde(rs("causale"))
		reso_tutto=rs("reso_tutto")
		if verde=false then
			verde=reso_tutto
		end if
		call imposta_verde(rs("idord"),verde)
		rs.movenext
	loop
	rs.close
	set rs = nothing
	response.write "ESEGUITO"
end if
call connclose()

%>