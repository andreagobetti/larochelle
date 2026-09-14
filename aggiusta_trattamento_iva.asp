<!--#include virtual="/setup.asp" -->
<%
	if request.form("esegui")<>"" then
		call aggiusta_trattamento_iva()
	end if
	
	
	
	%>


<form action="<%=questofile%>" method="post">

Imposta trattamento iva in utenti con 'comune' in azienda<br>
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 

<%
	
function aggiusta_trattamento_iva()

	sql="update  utenti set trattamento_iva=10  where azienda like '%comune%' "
	conn.execute sql,num
	response.write num

end function
	
	%>