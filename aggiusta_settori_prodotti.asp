<!--#include virtual="/setup.asp" -->
<%
	

		

%>


<form name="form1" method="post" action="<%=questofile%>">
	<input type="submit" name="a" value="Esegui"> Crea i settori per i livelli superiori<br>
	<input type="submit" name="b" value="Esegui"> Elimina i settori duplicati<br>
</form>
      
      <%
	      
	      
	      
	      
	      
if request.form("a")<>"" then
	response.write "Inizio<br>"
	sql_base="select settori_prodotti.* from settori_prodotti"
	set rs_settori=conn.execute (sql_base)
	do while not rs_settori.eof
		idpro=rs_settori("idpro")
		call scorri_settori_bkw(rs_settori("idsettore"))
		rs_settori.movenext
	loop
end if      
	
if request.form("b")<>"" then
	response.write "Inizio<br>"
	sql_base="select count(id), id  from settori_prodotti    group by idpro, idsettore  having count(id)>1 "
	set rs_settori=conn.execute (sql_base)
	do while not rs_settori.eof
		response.write "Elimino id:"&rs_settori("id")&"<br>"
		conn.execute ("delete from settori_prodotti where id="&rs_settori("id"))
		rs_settori.movenext
	loop
end if  	
	      
	      
sub scorri_settori_bkw(id)
	on error resume next

	set rs_tmp=conn.execute ("select * FROM settori WHERE idsettore = " & id&";")
	if not rs_tmp.eof then
			conn.execute ("insert into settori_prodotti (idpro,idsettore,ordine) values ("&idpro&","&rs_tmp("idsettore")&",1)")
			if err.number<>0 then
				response.write "Esiste gia<br>"
			else
				response.write "Aggiunto<br>"
			end if
			call scorri_settori_bkw(rs_tmp("idpadre"))
	end if
	set rs_tmp=nothing
	on error goto 0
end sub
%>