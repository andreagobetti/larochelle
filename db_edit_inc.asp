<%

function tb_edit(parametri)
	'parametri="sTableName: Record_locking, crearetabella:si, sNewName: campo, sName:, lType:tipocampo, lSize:0, sAutoIncrement:si, chkPrimaryKey: si, chkIndexed:si, chkIndexUnique:si"
	array_parametri=split(parametri,",")
	bNewField=true
	bEliminareCampo=false
	valoreDefault=""
	For i = 0 To UBound(array_parametri)
		parametro=split(array_parametri(i),":")
		parametro(0)=trim(parametro(0))
		parametro(1)=trim(parametro(1))
		'response.write "Parametro="&parametro(0)&"  valore="&trim(parametro(1))&"<br>"
			if parametro(0)="NomeTabella" then
				sTableName=parametro(1)
			end if
			if parametro(0)="CreareTabella" and parametro(1)<>"" then
				crearetabella=true
			end if

			if parametro(0)="NomeNuovoCampo" then
				sNewName=parametro(1)
			end if
			
			if parametro(0)="NomeCampo" and parametro(1)<>"" then
				sName=parametro(1)
				bNewField=false
			end if
			if parametro(0)="Tipo" then
				lType=lcase(parametro(1))
				select case lType
					case "integer"			: sFieldType = "SMALLINT"
					case "long"			: sFieldType = "INTEGER"
					case "boolean"			: sFieldType = "BIT"
					case "date"				: sFieldType = "DATE"
					case "currency"			: sFieldType = "CURRENCY"
					case "text"			: sFieldType = "VARCHAR"
					case "memo"		: sFieldType = "TEXT"
					case "ole"	: sFieldType = "LONGVARBINARY"
					case "guid"				: sFieldType = "GUID"
					case "byte"	: sFieldType = "TINYINT"
					case "single"	: sFieldType = "SINGLE"
				end select

			end if

			if parametro(0)="NCaratteri" and parametro(1)<>"" then
				lSize=CLng(parametro(1))
			end if
			if parametro(0)="Contatore" and parametro(1)<>"" then
				sAutoIncrement=true
			end if
			if parametro(0)="AccettaNull" and parametro(1)<>"" then
				sAcceptNulls=true
			end if
			if parametro(0)="ChiavePrimaria" and parametro(1)<>"" then
				sPrimaryKey=true
			end if
			if parametro(0)="Indicizzato" and parametro(1)<>"" then
				sIndexed=true
			end if
			if parametro(0)="ValoriUnici" and parametro(1)<>"" then
				sIndexUnique=true
			end if
			if parametro(0)="EliminareCampo" and parametro(1)<>"" then
				bEliminareCampo=true
				bNewField=false
				sFieldName=sName
			end if			
			if parametro(0)="Default" and parametro(1)<>"" then
				valoreDefault=parametro(1)
			end if			

	Next
	'response.end
	
	'sTableName : Nome nuova tabella
	'sName: Nome tabella esistente
	
		if sTableName = "" then
			sErr = "Error : <br>Please specify a table name."
			response.write sErr
			response.end
		end if
		if crearetabella then
			sSQL = "CREATE TABLE [" & sTableName & "]"
			on error resume next	

			conn.execute sSQL
			if err <> 0 then
				sErr = "Errore : <br>" & err.description
			end if
			On Error Goto 0
		end if
	
	
	
	if sName = "" then bFldAdd = True else bFldEdit=true
	'on error resume next	

		
		
		sSQL = "ALTER TABLE [" & sTableName & "] "
		
		if bNewField then
			'We are adding a new field
			sSQL = sSQL & " ADD COLUMN [" & sNewName & "] "
		else
			'We are editing a field
			sSQL = sSQL & " ALTER COLUMN [" & sName & "] "
		end if
		
		
		sSQL = sSQL & sFieldType
		
		if sFieldType = "VARCHAR" then
			'Set field size
			if lSize = "" then lSize = 50
			sSQL = sSQL & "(" & lSize & ") "
		end if
		
		if sAcceptNulls then
			sSQL = sSQL & " NULL "
		else
			sSQL = sSQL & " NOT NULL "
		end if
		
		if sAutoIncrement  then
			sSQL = sSQL & " IDENTITY "
		end if

		if valoreDefault<>"" then
			sSQL = sSQL & " DEFAULT "
			if sFieldType = "VARCHAR" then
				sSQL = sSQL & "'"&valoreDefault&"' "
			else
				sSQL = sSQL &valoreDefault&" "
			end if
		end if
		
		if sPrimaryKey then
			'If the field has a primary key index
			sSQL = sSQL & " CONSTRAINT [pk" & sNewName & "] PRIMARY KEY "
		else
			'If the field is indexed (not PK)
			if sIndexed  then
				if sIndexUnique <> "" then
					sSQL = sSQL & " CONSTRAINT [idx" & sNewName & "] UNIQUE"
				else
					'Index is no constraint, so another sql stat required
					sSQL2 = "CREATE INDEX [idx" & sNewName & "] ON [" & sTableName & "] ([" & sNewName & "]) "
				end if
			end if
		end if
		if bEliminareCampo then
			'DELETE field
			'Delete the indexes if any
			'response.write "adSchemaIndexes:"&adSchemaIndexes&"<br>"
			set rs = conn.openSchema(adSchemaIndexes)
			do while not rs.eof
				'response.write "table_name:"&rs("table_name")&"-"&sTableName&"<br>"
				if lcase(rs("table_name")) = lcase(sTableName) then
					if (rs("column_name") = sFieldName) then
						sSQL = "ALTER TABLE [" & sTableName & "] DROP CONSTRAINT [" & rs("index_name") & "]"
						response.write "rimuovo indice "&rs("index_name")&"<br>"
						response.write sSQL&"<br>"
						conn.execute sSQL
					end if
				end if
				rs.movenext
			loop

			sSQL = "ALTER TABLE [" & sTableName & "] DROP COLUMN [" & sFieldName & "]"
			response.write sSQL&"<br>"
			conn.execute sSQL
		
			'sErr = "Cannot rename field."
			'bErr = True
		end if
		if not bEliminareCampo then
			if (sName <> "") and (sNewName <> sName)  then
				'Rename field
				sSQL = "ALTER TABLE [" & sTableName & "] ALTER COLUMN RENAME [" & sName & "] TO [" & sNewName & "]"
				'sErr = "Cannot rename field."
				'bErr = True
			end if
			
			if bErr = True then
				'conn.close
				set rs = nothing
				set conn = nothing
				response.write "Errore:"&sErr
			else
				response.write sSQL&"<br>"
				
				conn.execute sSQL 
				if sSQL2 <> "" then
					response.write sSQL2&"<br>"
					conn.execute sSQL2
				end if
		
				if err <> 0 then
					%><%
					response.write err.description & "<br>" & sSQL
				else
					'response.redirect "te_tableedit.asp?cid=" & lConnID & "&tablename=" & server.urlencode(request("tablename"))
				end if
			end if
		end if

	tb_edit = true
end function
function query_edit(sQueryName,sSQL)
	dim rsquery
	dim rsproc
	dim cmd
	dim cat
	'Cerco nelle Query
	response.write "Cerco nelle query: "&sQueryName&", adSchemaViews:"&adSchemaViews&"<br>"
	
	set rsquery = conn.OpenSchema(adSchemaViews)
	do while not rsquery.eof
		response.write rsquery("table_name")&"<br>"
		if lcase(rsquery("table_name"))=lcase(sQueryName) then
			response.write "inizio Elimino query "&sQueryName&"<br>"
			dropSQL = "DROP TABLE ["&sQueryName&"]"
			conn.execute dropSQL
			response.write "Elimino query "&sQueryName&"<br>"
		end if
		rsquery.movenext
	loop
	rsquery.close
	
	'Cerco nelle Procedures
	response.write "Cerco nelle procedure: "&sQueryName&", adSchemaViews:"&adSchemaProcedures&"<br>"
	set rsproc = conn.OpenSchema(adSchemaProcedures)

	do while not rsproc.eof
		if lcase( rsproc("procedure_name"))=lcase(sQueryName) then
			response.write "inizio Elimino procedura "&sQueryName&"<br>"
			dropSQL = "DROP TABLE ["&sQueryName&"]"
			conn.execute dropSQL
			response.write "Elimino procedura "&sQueryName&"<br>"
		end if
		rsproc.movenext
	loop
	rsproc.close

	response.write "Creo query :"&sQueryName&"<br>"
	
	set cmd = server.createobject("adodb.command")
	set cat = server.createobject("adox.catalog")
	set cat.ActiveConnection = conn
	set cmd.ActiveConnection = conn
	
	cmd.CommandType = adCmdText
	cmd.CommandText = sSQL

	set views = cat.views
	
	views.append sQueryName, cmd
	set cmd=nothing
	set cat=nothing
		
	query_edit=true
	set rs=nothing
end function
	
Function esiste_campo(tabella,campo)
	sql = "select * from "&tabella&";"
	set rs = conn.execute(sql)
	For Each field in rs.Fields
	   if lcase(field.Name) = lcase(campo) then 
		  esiste_campo = true 
	  exit for
	   else 
		  esiste_campo = false
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
function elimina_tabella(tabella)
	if esiste_tabella(tabella) then
		conn.execute ("drop table "&tabella)
		response.write tabella&" eliminata<br>"
	else
		response.write tabella&" non esiste<br>"
	end if
end function


sub imposta_versione(versione)
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "impostazioni", conn, 3, 3
	rs.movefirst
	rs("versione")=versione
	rs.update
	rs.close
	set rs=nothing
	conn.execute("insert into log (data,evento,causale) values( now(),'Aggiornato database a versione "&versione&"',3) ")
	
	
end sub
%>