<!--#include virtual="/setup.asp" -->
<%
	
if request.form("ok")<>"" then
	call esegui()
	
		response.write("Eseguito")

end if
		

%>
Imposta settore predefinito sugli articoli prendendo il primo settore<br>

<form name="form1" method="post" action="<%=questofile%>">
	Suffisso tabella <input type="text" name="suffisso" value="wl3d">
	<input type="submit" name="ok" value="Esegui">
</form>
      
      <%
sub esegui()
	
	
	
	
	
	suffisso=request.form("suffisso")
	tabelle="settori|prodotti|settori_settori|settori_prodotti|prodotti_dettagli|files"
	array_tabelle=split(tabelle,"|")
	for n=0 to ubound(array_tabelle)
	
	
		if not esiste_tabella(array_tabelle(n)&suffisso) then
			response.write "Manca la tabella:"&array_tabelle(n)
			response.end
		end if
	next

	
	dim nome_campo()

	if true then
		'settori
		tabella="settori"
		
		'Imposto idsettore_tmp
		'if not esiste_campo(tabella&suffisso,"idsettore_tmp") then
			response.write "creo idsettore_tmp"
			conn.execute "ALTER TABLE  `"&tabella&suffisso&"` ADD  `idsettore_tmp` INT NOT NULL"	
		'end if
		
		conn.execute ("update "&tabella&suffisso&" set idsettore_tmp=idsettore")
		
		if not esiste_campo(tabella,"idsettore_tmp") then
			conn.execute "ALTER TABLE  '"&tabella&"' ADD  `idsettore_tmp` INT NOT NULL"	
		end if
	
		
		
		
		
		set rs_from=conn.execute ("select * from settori"&suffisso)
		n=rs_from.fields.count-1
		sql="select * from settori"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		
		'creo array con nomi campo
		redim nome_campo(n+1)
		'memorizzo i nomi dei campi e i valori attuali
		response.write "elenco campi destinazione:"
		for i=0 to n
			response.write rs.Fields(i).Name&", "
			nome_campo(i)=rs.Fields(i).Name
		next
		response.write "<br>"

		do while not rs_from.EOF
			rs.addnew
			for i=1 to n
				
					'response.write nome_campo(i)&"<br>"
					rs(nome_campo(i))=rs_from(nome_campo(i))
				next			
			rs.update
		
		rs_from.MoveNext
		loop
		response.write "<br>Settori importati<br>"
	end if
	'PRODOTTI
	if true then
		tabella="prodotti"
		'Imposto idpro_tmp
		
		if not esiste_campo(tabella&suffisso,"idpro_tmp") then
			conn.execute "ALTER TABLE  `"&tabella&suffisso&"` ADD  `idpro_tmp` INT NOT NULL"	
		end if
		
		conn.execute ("update "&tabella&suffisso&" set idpro_tmp=idpro")
		
		if not esiste_campo(tabella,"idpro_tmp") then
			conn.execute "ALTER TABLE  "&tabella&" ADD  idpro_tmp INT NOT NULL"	
		end if
		
		
		set rs_from=conn.execute ("select * from "&tabella&suffisso)
		n=rs_from.fields.count-1
		sql="select * from "&tabella
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		
		'creo array con nomi campo
		redim nome_campo(n+1)
		'memorizzo i nomi dei campi e i valori attuali
		response.write "elenco campi "&tabella&":"
		for i=0 to n
			response.write rs.Fields(i).Name&", "
			nome_campo(i)=rs.Fields(i).Name
		next
		response.write "<br>"
		on error resume next
		do while not rs_from.EOF
			rs.addnew
			for i=1 to n
				
					'response.write nome_campo(i)&"<br>"
					rs(nome_campo(i))=rs_from(nome_campo(i))
				next			
			rs.update
		
		rs_from.MoveNext
		loop
		on error goto 0
		response.write "<br>"&tabella&" importati<br>"
	end if
	'settori_settori
	if true then
		tabella="settori_settori"
		if true then
			sql="select settori_settori"&suffisso&".*, settori.idsettore as nuovo_idpadre from "&tabella&suffisso&" inner join settori on settori_settori"&suffisso&".idpadre = settori.idsettore_tmp where settori_settori"&suffisso&".idpadre>0 "
			
			set rs= conn.execute(sql)
			do while not rs.EOF
				conn.execute ("update "&tabella&suffisso&" set idpadre="&rs("nuovo_idpadre")&" where id="&rs("id"))
				rs.movenext
			loop
			response.write "<br>idpadrea aggiornato<br>"
		end if
		if true then
			sql="select settori_settori"&suffisso&".*, settori.idsettore as nuovo_idfiglio from "&tabella&suffisso&" inner join settori on settori_settori"&suffisso&".idfiglio = settori.idsettore_tmp where settori_settori"&suffisso&".idfiglio>0 "
			
			response.write sql 
			set rs= conn.execute(sql)
			do while not rs.EOF
				conn.execute ("update "&tabella&suffisso&" set idfiglio="&rs("nuovo_idfiglio")&" where id="&rs("id"))
				rs.movenext
			loop
			response.write "<br>idfiglio aggiornato<br>"
		end if
		
		
		
		
		
		set rs_from=conn.execute ("select * from "&tabella&suffisso)
		n=rs_from.fields.count-1
		sql="select * from "&tabella
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		
		'creo array con nomi campo
		redim nome_campo(n+1)
		'memorizzo i nomi dei campi e i valori attuali
		response.write "elenco campi "&tabella&":"
		for i=0 to n
			response.write rs.Fields(i).Name&", "
			nome_campo(i)=rs.Fields(i).Name
		next
		response.write "<br>"

		do while not rs_from.EOF
			rs.addnew
			for i=1 to n
				
					'response.write nome_campo(i)&"<br>"
					rs(nome_campo(i))=rs_from(nome_campo(i))
				next			
			rs.update
		
		rs_from.MoveNext
		loop
		response.write "<br>"&tabella&" importati<br>"
	end if
	'settori_prodotti
	if true then
		tabella="settori_prodotti"
		if true then
			sql="select "&tabella&suffisso&".*, settori.idsettore as nuovo_idpadre from "&tabella&suffisso&" inner join settori on "&tabella&suffisso&".idsettore = settori.idsettore_tmp  "
			
			set rs= conn.execute(sql)
			do while not rs.EOF
				conn.execute ("update "&tabella&suffisso&" set idsettore="&rs("nuovo_idpadre")&" where id="&rs("id"))
				rs.movenext
			loop
			response.write "<br>idsettore aggiornato<br>"
		end if
		if true then
			sql="select "&tabella&suffisso&".*, prodotti.idpro as nuovo_idpro from "&tabella&suffisso&" inner join prodotti on "&tabella&suffisso&".idpro = prodotti.idpro_tmp  "
			
			response.write sql 
			set rs= conn.execute(sql)
			do while not rs.EOF
				conn.execute ("update "&tabella&suffisso&" set idpro="&rs("nuovo_idpro")&" where id="&rs("id"))
				rs.movenext
			loop
			response.write "<br>idpro aggiornato<br>"
		end if
		
		
		
		
		set rs_from=conn.execute ("select * from "&tabella&suffisso)
		n=rs_from.fields.count-1
		sql="select * from "&tabella
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		
		'creo array con nomi campo
		redim nome_campo(n+1)
		'memorizzo i nomi dei campi e i valori attuali
		response.write "elenco campi "&tabella&":"
		for i=0 to n
			response.write rs.Fields(i).Name&", "
			nome_campo(i)=rs.Fields(i).Name
		next
		response.write "<br>"

		do while not rs_from.EOF
			rs.addnew
			for i=1 to n
				
					'response.write nome_campo(i)&"<br>"
					rs(nome_campo(i))=rs_from(nome_campo(i))
				next			
			rs.update
		
		rs_from.MoveNext
		loop
		response.write "<br>"&tabella&" importati<br>"
	end if
	'prodotti_dettagli
	if true then
		tabella="prodotti_dettagli"
		if true then
			sql="select "&tabella&suffisso&".*, prodotti.idpro as nuovo_idpro from "&tabella&suffisso&" inner join prodotti on "&tabella&suffisso&".idpro = prodotti.idpro_tmp  "
			
			response.write sql 
			set rs= conn.execute(sql)
			do while not rs.EOF
				conn.execute ("update "&tabella&suffisso&" set idpro="&rs("nuovo_idpro")&" where id="&rs("id"))
				rs.movenext
			loop
			response.write "<br>idpro aggiornato<br>"
		end if
		
		
		
		set rs_from=conn.execute ("select * from "&tabella&suffisso)
		n=rs_from.fields.count-1
		sql="select * from "&tabella
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		
		'creo array con nomi campo
		redim nome_campo(n+1)
		'memorizzo i nomi dei campi e i valori attuali
		response.write "elenco campi "&tabella&":"
		for i=0 to n
			response.write rs.Fields(i).Name&", "
			nome_campo(i)=rs.Fields(i).Name
		next
		response.write "<br>"

		do while not rs_from.EOF
			rs.addnew
			for i=1 to n
				
					'response.write nome_campo(i)&"<br>"
					rs(nome_campo(i))=rs_from(nome_campo(i))
				next			
			rs.update
		
		rs_from.MoveNext
		loop
		response.write "<br>"&tabella&" importati<br>"
	end if
	'files
	if true then
		tabella="files"
		conn.Execute("delete from files"&suffisso&" where cosa<>4")
		
		if true then
			sql="select "&tabella&suffisso&".*, prodotti.idpro as nuovo_idpro from "&tabella&suffisso&" inner join prodotti on "&tabella&suffisso&".idcosa = prodotti.idpro_tmp  "
			
			response.write sql 
			set rs= conn.execute(sql)
			do while not rs.EOF
				conn.execute ("update "&tabella&suffisso&" set idcosa="&rs("nuovo_idpro")&" where idfiles="&rs("idfiles"))
				rs.movenext
			loop
			response.write "<br>idpro aggiornato<br>"
		end if
		
		
		
		set rs_from=conn.execute ("select * from "&tabella&suffisso)
		n=rs_from.fields.count-1
		sql="select * from "&tabella
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql, conn, 3, 3
		
		'creo array con nomi campo
		redim nome_campo(n+1)
		'memorizzo i nomi dei campi e i valori attuali
		response.write "elenco campi "&tabella&":"
		for i=0 to n
			response.write rs.Fields(i).Name&", "
			nome_campo(i)=rs.Fields(i).Name
		next
		response.write "<br>"

		do while not rs_from.EOF
			rs.addnew
			for i=1 to n
				
					'response.write nome_campo(i)&"<br>"
					rs(nome_campo(i))=rs_from(nome_campo(i))
				next			
			rs.update
		
		rs_from.MoveNext
		loop
		response.write "<br>"&tabella&" importati<br>"
	end if
end sub

Function esiste_campo(tabella,campo)
	sql = "select * from "&tabella&";"
	set rs = conn.execute(sql)
	For Each field in rs.Fields
	   if lcase(field.Name) = lcase(campo) then 
		  esiste_campo = true 
	  exit for
	   else 
		  esiste_campo = true
	   end if
	next
	rs.close
	set rs=nothing
End Function

Function esiste_tabella(tabella)
	dim rs
	tabella=lcase(tabella)
	esiste_tabella=false
	set rs = conn.OpenSchema(20)
      do while not rs.eof
          if rs("table_type") = "TABLE" then
              'response.write "Tabella " & rs("table_name")
              if tabella=lcase(rs("table_name")) then
	              esiste_tabella=true
	              exit do
              end if
          end if
          rs.movenext
      loop
	rs.close
	set rs=nothing
End Function


%>