<!--#include virtual="/setup.asp" -->
<%

file_yatego

sub file_yatego()
	idpro=5206

	'response.ContentType="application/vnd.ms-excel"
	Dim objFSO, objTextFile 
	'Creazione dell'istanza filesystem 
	Set objFSO = CreateObject("Scripting.filesystemObject") 
	'Apertura del file di testo 
	Set objTextFile = objFSO.CreateTextFile(Server.MapPath("\public\yatego.csv")) 
	add2log "Iniziata esportazione <b>Yatego</b>",1
	sql="select prodotti.* FROM (prodotti) "
	sql=sql&" where yatego=true "
	Set rs_var_a = Server.CreateObject("ADODB.Recordset")
	Set rs_var_b = Server.CreateObject("ADODB.Recordset")
	set rs=conn.execute(sql)
	n_articoli=0
	n_righe=0
	do while not rs.eof
		rs_var_a.Open "select * from varianti_a where idpro=" & rs("idpro"), conn, 3, 3
		rs_var_b.Open "select * from varianti_b where idpro=" & rs("idpro"), conn, 3, 3
		variante_a=false
		variante_b=false
		if not rs_var_a.eof then variante_a=true
		if not rs_var_b.eof then variante_b=true

		'Scrivo la prima riga di codice
		riga=""
		'add2log "variante_a"&variante_a&"variante_b"&variante_b,1
		if not variante_a then
			varianti=""
							objTextFile.Write record_yatego(rs("codice")&varianti,rs("idpro"),rs("articolo")&varianti,rs("prezzo"),rs("prezzo"),rs("Descrizione_listino"),rs("idpro"),rs("idcat"))
							n_righe=n_righe+1

		else
				rs_var_a.movefirst
				do until rs_var_a.eof
					varianti="-"&rs_var_a("variante_a")
					if not variante_b then
							varianti="-"&rs_var_a("variante_a")
							objTextFile.Write record_yatego(rs("codice")&varianti,rs("idpro"),rs("articolo")&varianti,rs("prezzo"),rs("prezzo"),rs("Descrizione_listino"),rs("idpro"),rs("idcat"))
							n_righe=n_righe+1
					else
						rs_var_b.movefirst
						do until rs_var_b.eof
							varianti="-"&rs_var_a("variante_a")&"-"&rs_var_b("variante_b")
							objTextFile.Write record_yatego(rs("codice")&varianti,rs("idpro"),rs("articolo")&varianti,rs("prezzo"),rs("prezzo"),rs("Descrizione_listino"),rs("idpro"),rs("idcat"))
							rs_var_b.movenext
							n_righe=n_righe+1
						loop
					end if
					rs_var_a.movenext
				loop
			end if	
		rs_var_a.close
		rs_var_b.close
		
		
		rs.movenext
		n_articoli=n_articoli+1

	loop
	
		
		'Scrivo la riga
		objTextFile.Write riga 
		
		'Chiudo il file e i vari oggetti/istanze 
		objTextFile.Close 
		Set objTextFile = Nothing 
		Set objFSO = Nothing
		
		if true then
		Response.Buffer = True
		Response.Clear
		
		Dim strFilePath, strFileName
		
		Const adTypeBinary = 1
		
		Set objStream = Server.CreateObject("ADODB.Stream")
		
		strFilePath = Server.MapPath("\public\yatego.csv")
		
		objStream.Open
		objStream.Type = adTypeBinary
		objStream.LoadFromFile strFilePath 
		
		Response.AddHeader "Content-Disposition", "attachment; filename=" & "yatego.csv"
		Response.Charset = "UTF-8"
		Response.ContentType = "text/plain"
		'Google "Mime Types" for addtional Content Type definitions
		
		Response.BinaryWrite objStream.Read
		Response.flush
		objStream.Close
		Set objStream = Nothing
		else
		response.write riga
		end if
		if session("idadmin")="" then
			generato_da=", richiesta da <b>Yatego</b>"
		else
			generato_da=", richiesta da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"
		end if
	add2log "Terminata esportazione Yatego, articol:"&n_articoli&", righe totali:"&n_righe&generato_da,2
end sub

function record_yatego(foreign_id,article_nr,title,price,price_uvp,long_desc,picture,categories)
	separatore="|"

	'foreign_id
	lunghezza_campo=30
	riga_tmp=riga_tmp&left(foreign_id,lunghezza_campo)&varianti&separatore
	'article_nr
	lunghezza_campo=30
	riga_tmp=riga_tmp&left(article_nr,lunghezza_campo)&separatore
	'title
	lunghezza_campo=255
	riga_tmp=riga_tmp&left(title,lunghezza_campo)&separatore
	'tax
	lunghezza_campo=255
	riga_tmp=riga_tmp&left(iva(date()),lunghezza_campo)&separatore
	'price
	riga_tmp=riga_tmp&formatnumber(price*(1+(iva(date())/100)),2)&separatore
	'price_uvp
	riga_tmp=riga_tmp&formatnumber(price_uvp*(1+(iva(date())/100)),2)&separatore
	'units
	riga_tmp=riga_tmp&separatore
	'delivery_surcharge
	riga_tmp=riga_tmp&"7"&separatore
	'delivery_calc_once
	riga_tmp=riga_tmp&"0"&separatore
	'short_desc	
	riga_tmp=riga_tmp&separatore
	'long_desc	
	lunghezza_campo=255
	long_desc=replace(long_desc,vbcrlf," ")
	riga_tmp=riga_tmp&left(long_desc,lunghezza_campo)&separatore
	'url
	riga_tmp=riga_tmp&separatore
	'picture
	riga_tmp=riga_tmp&"http://"&nomesito&upl_img_cat &picture&"_b.jpg"&separatore
	'picture2
	riga_tmp=riga_tmp&separatore
	'picture3
	riga_tmp=riga_tmp&separatore
	'picture4
	riga_tmp=riga_tmp&separatore
	'picture5
	riga_tmp=riga_tmp&separatore
	'categories
	riga_tmp=riga_tmp&categoria_1(categories,"")&separatore
	'stock
	riga_tmp=riga_tmp&"10"&separatore
	record_yatego=riga_tmp&vbcrlf
end function

%>