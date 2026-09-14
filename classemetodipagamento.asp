<%
'-------------------------Funzioni per incassi INIZIO

class ClasseMetodipagamento
	'Proprietà
	'Public n_Elementi
	dim Dict,rs
	
	'Costruttore
	Private Sub Class_Initialize()
		Set Dict = New MultiDimensionalDictionary 'Create an Instance of the Class
		set rs=conn.execute("select * from metodi_pagamento order by descrizione")
		
		do while not rs.eof
			Dict.SetKey = cint(rs("id").value) 'Key record
			Dict.SetField "descrizione", rs("descrizione").value
'			Dict.SetField "genera_scadenze", rs("genera_scadenze").value
			Dict.SetField "numero_scadenze", rs("numero_scadenze").value
			Dict.SetField "inizio_scadenze", rs("inizio_scadenze").value
			Dict.SetField "riba", rs("riba").value
			Dict.SetField "pagamentopa", rs("pagamentopa").value
			
			Dict.Update 'Bind the record and preapre for the next record
			rs.MoveNext
		Loop
		set rs = nothing
'		Response.Write "Dict.Count = " &  Dict.Count &  "<br/>"
'		'######################################################################
'		'ENUMERATE RECORDS
'		'This method shows you how to enumerate the record/field collection.
'		'######################################################################
'		Response.Write("-- ENUMERATE RECORDS --<br>")
'		For Each Record in Dict.Records
'		Response.Write("<strong>Record: " &  Record &  "</strong><br>")
'		For Each Field in Dict.Fields(Record)
'		Response.Write("&nbsp; &nbsp; &nbsp;<strong>Field Name:</strong> " & Field &  " : <strong> Value:</strong> " &  Dict.Item(Record, Field) &  "<br>")
'		Next
'		Response.Write("<hr>")
'		Next


	end sub
	
	'Distruttore
	Private Sub Class_Terminate()
		set Dict = nothing
	End Sub
	
	Public function descrizione(metodo_pagamento)
	if isnull(metodo_pagamento) then
		descrizione="Non specificato"
	else
		descrizione=Dict.Item(cint(metodo_pagamento), "descrizione") 
	end if
	end function
	
	public Property Get count ()
	   count=Dict.count
	End Property
	
	public sub stampa_option (valore)
		if isnull(valore) then valore=0
		For Each Record in Dict.Records
			Response.Write("<option value="""&Record&"""")
			if cint(valore)=cint(Record) then response.write " selected"
			Response.Write ">" &  Dict.Item(Record, "descrizione") &  "</option>"&vbcrlf
		Next

	End sub
	
	
	Public function campo(IDpagamento,nome_campo)
		if isnull(IDpagamento) then exit function
		IDpagamento=cint(IDpagamento)
		campo=Dict.Item(IDpagamento, nome_campo)
	end function
	
	Public function numeroScadenze(metodo_pagamento)
	if isnull(metodo_pagamento) then
		scadenze=0
	else
		scadenze=Dict.Item(cint(metodo_pagamento), "numero_scadenze") 
	end if
	end function


	Public function inizio(metodo_pagamento)
	if isnull(metodo_pagamento) then
		inizio=0
	else
		inizio=Dict.Item(cint(metodo_pagamento), "inizio_scadenze") 
	end if
	end function


	public function genera_scadenze(idfat,nfat,idord,data_fattura,importo,metodo_pagamento)
		call add2log("importo:"&importo,0)
		genera_scadenze=0
		metodo_pagamento=cint(metodo_pagamento)
		dim riba
		sql=""
		if idfat>0 then
			sql = "delete FROM scadenze WHERE idfat="&idfat
		elseif idord>0 then
			sql = "delete FROM scadenze WHERE idord="&idord
		end if
		if sql<>"" then conn.execute (sql)
		
		numero_scadenze=cint(Dict.Item(metodo_pagamento, "numero_scadenze"))
		
		if  numero_scadenze>0 then
		
			sql=""
			importo_scadenza=roundup(importo/numero_scadenze,2)
			call add2log("importo_scadenza:"&importo_scadenza,0)
			importo_cumulato=0
			sql = "select * FROM scadenze"
			Set rs_incassi = Server.CreateObject("ADODB.Recordset")
			rs_incassi.Open sql, conn, 3, 3
			for n=1 to numero_scadenze
				rs_incassi.addnew
				data=finemese(data_fattura,n+Dict.Item(metodo_pagamento, "inizio_scadenze"))
				rs_incassi("data")=data
				if n=numero_scadenze then importo_scadenza=roundup(importo-importo_cumulato,2)
			call add2log("importo_scadenza 2:"&importo_scadenza,0)
				rs_incassi("importo")=importo_scadenza
			call add2log("importo:"&rs_incassi("importo"),0)
				rs_incassi("idord")=idord
				rs_incassi("idfat")=IDfat
				rs_incassi("causale")=0
				rs_incassi("note_pagamento")=Dict.Item(metodo_pagamento, "descrizione")&" FATTURA "&pre_fattura(data_fattura)&nfat&" "&n&"&#47;"&numero_scadenze
				riba=converti_bool(Dict.Item(metodo_pagamento, "riba"))
				'response.write "riba="&riba
				rs_incassi("riba")=riba
				rs_incassi.update
				idscadenza=get_last_id("scadenze")
				testo=testo&"Creata [scadenza="&idscadenza&"] riba="&riba&" " &importo_scadenza&"<br>"
				importo_cumulato=importo_cumulato+importo_scadenza
				'response.write "Genero scadenza data "&data&" importo "&importo_scadenza&"<br>"
				
			next
			rs_incassi.close
			set rs_incassi=nothing
			testo=testo&"per [fattura="&idfat&"]"&nfat&"[/fattura]"
			add2log testo,2
			genera_scadenze=scadenze
		end if

		
	end function
	
		
	Public function spese_bancarie(IDpagamento)
		IDpagamento=cint(IDpagamento)
		if Dict.Item(IDpagamento, "riba")=1 then
			spese_bancarie=3.7*Dict.Item(IDpagamento, "numero_scadenze")
		else
			spese_bancarie=0
		end if
	end function

	

	
	
	
end class


%>
