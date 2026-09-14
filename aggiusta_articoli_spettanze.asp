<!--#include virtual="/setup.asp" -->
<%
	

		

%>


<form name="form1" method="post" action="<%=questofile%>">
	<input type="submit" name="b" value="Esegui"> Elimina i articoli spettanze duplicati<br>
</form>
      
      <%
	      
	      
	      
	      
	      
	
if request.form("b")<>"" then
	response.write "Inizio<br>"
	sql_base="select count(id), id  from prodotti_spettanze    group by idpro, iduser  having count(id)>1 "
	set rs_settori=conn.execute (sql_base)
	do while not rs_settori.eof
		response.write "Elimino id:"&rs_settori("id")&"<br>"
		conn.execute ("delete from prodotti_spettanze where id="&rs_settori("id"))
		rs_settori.movenext
	loop
end if  	
	      
	      
%>