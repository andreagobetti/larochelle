<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Rivalorizza il campo cinturone nelle misure dipendente con vita giacca * 2 + 20<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	Server.ScriptTimeout=180
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from dipendenti_misure"
	
	rs.Open sql,conn,3,3
	do while not rs.eof

		if rs("giacca_vita")<>"" then
			giacca_vita=rs("giacca_vita")
			if isnumeric(giacca_vita) then
				cinturone=cstr((cint(giacca_vita)*2)+20)
				'response.write "giacca_vita="&giacca_vita&" misura cinturone="&cinturone&"<br>"
				rs("cinturone")=cinturone
			else
				response.write "Saltato dipendente iddip:"&rs("iddip")&"<br>"
				
			end if
			rs.update			
		end if

		rs.movenext
	loop
	rs.close
	set rs = nothing
	response.write "ESEGUITO"
end if
call connclose()

%>