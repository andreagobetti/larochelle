<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<%
idord=request("idord")

n=pdf_nota_credito(idord,"pdf")
	
function pdf_nota_credito(idord,out)
	dim rs
	
	sql="select ordini.*,  utenti.cellulare,utenti.idbanca, utenti.fax, ordini.trasporto, ordini.mepa_testo, ordini.mepa_tipo, ordini.cig,  ordini.tipo_documento, utenti_clienti.FatturaPACodiceDestinatario, utenti_clienti.stato_estero, ordini.idord, ordini.nord, utenti.email, utenti.telefono,  ordini.note, ordini.tipo_trasporto, ordini.noteacq FROM (ordini INNER JOIN utenti ON ordini.iduser = utenti.iduser) LEFT JOIN utenti_clienti ON utenti.iduser = utenti_clienti.iduser"
	sql=sql&" WHERE ordini.idord= " &idord  &";"
	
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	
	
	if not rs.eof then
		if session("idadmin")="" and request("id2")<>chkDataUser(rs("data")) then
			add2log "Tentativo di visualizare PDF fattura, id2 incongruente:"&request("id2")&vbcrlf&allCookies(),3
			call esci(8,idfat)
		end if
		if session("idadmin")="" then
			conn.execute ("update fatture set download=now() where idfat="&idfat)
			add2log "[utente="&rs("iduser")&"]"&nominativo(rs("nome"),rs("cognome"),rs("azienda"))&"[/utente] ha scaricato la [notacredito="&rs("idord")&"]"&rs("nord")&"[/notacredito]",2
		end if
	
		dim vecchio_conteggio
		if DateDiff("d","01/01/2015",rs("data"))<0 then 'prima del 01/01/2015
			vecchio_conteggio=true
		else
			vecchio_conteggio=false
		end if
		Set Pdf = Server.CreateObject("Persits.PDFManager")
		Set Doc = Pdf.CreateDocument
		Set TextParam = PDF.CreateParam
		Set ImageParam = PDF.CreateParam
		'Conversione decimal in dbl
		totale_fattura=cdbl(rs("totale"))
		trasporto=cdbl(rs("trasporto"))
		spese_bancariev=cdbl(rs("spese_bancarie_ordine"))
		sconto_ordine=rs("sconto_ordine")
		if not isnull(sconto_ordine) then
			sconto_ordine=cdbl(sconto_ordine)
		else
			sconto_ordine=0
		end if
		imposta=cdbl(rs("imposta_ordine"))
		totale_merce=cdbl(rs("totale_merce_ordine"))
		
		Doc.Title = "Nota Credito "&nord
		titolo_doc= "Nota credito "
			
		Doc.Creator = nomesito
		nord=rs("nord")
		
		codicedestinatario=get_codice_destinatario("",rs("FatturaPACodiceDestinatario"))
		if codicedestinatario<>"" then
			codiceufficio="<br>Codice univoco ufficio: "&codicedestinatario
		else
			codiceufficio=""
		end if
		
		
		Set Page = Doc.Pages.Add
		'logo
		Set fs = Server.CreateObject("Scripting.filesystemObject")
		if fs.FileExists(Server.MapPath( "/public/files/logo_documenti.jpg")) then
			esiste_logo=true
			
			'Parte nuova
			margine=5
			Set logo = Doc.OpenImage( Server.MapPath( "/public/files/logo_documenti.jpg") )
			'calcolo misure dell'immagine considerando la risoluzione DPI
			larghezza=logo.width*72/logo.resolutionx
			altezza=logo.height*72/logo.resolutiony
			
			
			scala=(dimensione_logo_fattura)/larghezza
			
			'calcolo nuova altezza considerando la scala per centrare poi l'immagine
			nuova_altezza=altezza*scala
			nuova_larghezza=larghezza*scala
			
			Set Param_logo = PDF.CreateParam
			
			'Trovo la x per centrare l'immagine
			'Param_logo("x") =  (row.cells(1).width - nuova_larghezza ) / 2 'Centra l'immagine
			'Param_logo("x") =  (row.cells(1).width - nuova_larghezza ) + margine'Allinea a SX
			Param_logo("x") =(Page.Width/ 2) - margine - dimensione_logo_fattura
			
			'Sotto laschio metà del margine
			ImageParam("y") = margine/2
			Param_logo("y") = Page.Height - 110
			
			Param_logo("ScaleX") = scala
			Param_logo("ScaleY") = scala
		    'session("pdf_fatturaa")= "larghezza immagine: "&larghezza&" larghezza colonna: "&dimensione_logo_fattura&" scala:"& scala & "larghezza immagine: "&(image.width*scalex)&" scala_txt:"& scalex_txt&"ImageParam.scalex"&Param_logo("scalex"), "alignment=0; size=8; html=true; expand=true;"
		end if
		
		
			Bordo_sx=25
			page.height=825
			Set Param = Pdf.CreateParam
			Set Param2 = Pdf.CreateParam
			'TABELLA 1
			Set Table1 = Doc.CreateTable("width=300; height=600; Rows=1; Cols=3; Border=1; CellSpacing=-1; cellpadding=2 ")
			Table1.Font = Doc.Fonts("Arial")
			'table1.rows.add(15)
			With Table1.Rows(1)
				.BGColor = &HCCCCCC
				.Cells(1).Width = page.width/6+12
				.Cells(1).AddText titolo_doc&pre_fattura(rs("data"))&aggiungi_zeri(rs("nord"))&post_fattura&" / "&year(rs("data")), Param
				.Cells(2).Width = page.width/6-18
				.Cells(2).AddText "del "& formatDateTime(rs("data"), vbShortDate) , Param
				.Cells(3).Width = page.width/6-14
				.Cells(3).AddText "Pagina 1 di 1", Param
				.height=20
			End With
			'Page.Canvas.DrawTable Table1, "x="& page.width/2 &", y="&page.height-20
			'TABELLA 2
			Set Table2 = Doc.CreateTable("width=300; height=600; Rows=2; Cols=4; Border=1; CellSpacing=-1; cellpadding=2 ")
			Table2.Font = Doc.Fonts("Arial")
			'table1.rows.add(15)
			param.add "html=true"
			With Table2.Rows(1)
				.cells(1).colspan=3
				.cells(1).border=0
				.Cells(1).Width = (page.width/2-34)/4
				.Cells(4).Width = (page.width/2-34)/4
				.Cells(1).AddText "SPETT.LE", "size=9; html=true;"
				
				.Cells(4).AddText "COD.CLIENTE<br><b>"& rs("iduser")&"</b>", "size=8; html=true; Expand=true;"
				.height=28
			End With
			With Table2.Rows(2)
				.cells(1).colspan=4
				.cells(1).border=0
				.Cells(1).AddText dati_fatturazione_pdf(rs("idintestazione"))&codiceufficio, "size=9; html=true; expand=true;"
				
				'testo="P.Iva "&rs("piva")&" "&"Cod. fisc. "&rs("cf")
				'.Cells(1).AddText "<b>"&ucase(rs("azienda")&"<P></P>"&rs("indirizzo")&"<br>"&rs("cap")&" "&rs("citta")&" "&rs("provincia")&"</b>")&"<br>"&testo, "size=9; html=true;"
				.height=70
			End With
			'TABELLA 21
			banca_appoggio_arr=banca_appoggio(rs("idbanca"))
			
			Set Table21 = Doc.CreateTable("width="&page.width-40&"; height=100; Rows=2; Cols=2; Border=1; CellSpacing=-1; cellpadding=2 ")
			Table21.Font = Doc.Fonts("Arial")
			With Table21.Rows(1)
				if metodo_pagamento(rs("tipopagamento"),"riba")=1 then
					testo=rs("banca_appoggio")
				else
					testo=banca_appoggio_arr(0)
				end if
				.Cells(1).AddText "BANCA D'APPOGGIO<br><b>"&ucase(testo)&"</b>", "size=8; html=true; Expand=true;"
				.height=28
				.Cells(1).Width = (page.width-40)/2
				.Cells(2).Width = (page.width-40)/2
				'.Cells(2).AddText "ANNOTAZIONI<br><b>"& rs("annotazioni")&"</b>", "size=8; html=true; Expand=true;"
				.Cells(2).AddText "ANNOTAZIONI<br>", "size=8; html=true; Expand=true;"
			End With
			With Table21.Rows(2)
				'.BGColor = &HCCCCCC
				if rs("tipopagamento") >= 10 and rs("tipopagamento")<=14 then
					testo=rs("iban")
				else
					testo=banca_appoggio_arr(1)
				end if
				.Cells(1).AddText "IBAN<br><b>"&testo&"</b>", "size=8; html=true; Expand=true;"
				.Cells(2).AddText "MODALITA' DI PAGAMENTO<br><b>"& ucase( metodo_pagamento(rs("tipopagamento"),"descrizione"))&"</b>", "size=8; html=true; Expand=true;"
				.height=28
			End With
			'TABELLA5
			Set Table5 = Doc.CreateTable("width="&page.width-40&"; height=600; Rows=1; Cols=8; Border=1; CellSpacing=-1; cellpadding=2 ")
			Table5.Font = Doc.Fonts("Arial")
			With Table5.Rows(1)
				.BGColor = &HCCCCCC
				.Cells(1).Width = 60
				.Cells(1).AddText "COD. ARTICOLO", "alignment=2; size=8;"
				.Cells(2).Width = page.width-60-30-40-50-50-55-30-34
				.Cells(2).AddText "DESCRIZIONE", "alignment=2; size=8;"
				.Cells(3).Width = 30
				.Cells(3).AddText "U.M.", "alignment=2; size=8;"
				.Cells(4).Width = 40
				.Cells(4).AddText "Q.TA'", "alignment=2; size=8;"
				.Cells(5).Width = 50
				.Cells(5).AddText "PR.UNIT.", "alignment=2; size=8;"
				.Cells(6).Width = 50
				.Cells(6).AddText "SCONTO", "alignment=2; size=8;"
				.Cells(7).Width = 55
				.Cells(7).AddText "TOTALE", "alignment=2; size=8;"
				.Cells(8).Width = 30
				.Cells(8).AddText "IVA", "alignment=2; size=8;"
				.height=20
			End With
			Table5.Rows.Add(450)
			Set Table4 = Doc.CreateTable("width="&page.width-40&"; height=600; Rows=1; Cols=8; Border=1; CellSpacing=-1; cellpadding=2 ")
			Table4.Font = Doc.Fonts("Arial")
			With Table4.Rows(1)
				.Cells(1).Width = table5.rows(1).cells(1).width
				.Cells(1).border =0
				.Cells(2).Width = table5.rows(1).cells(2).width
				.Cells(2).border =0
				.Cells(3).Width = table5.rows(1).cells(3).width
				.Cells(3).border =0
				.Cells(4).Width = table5.rows(1).cells(4).width
				.Cells(4).border =0
				.Cells(5).Width = table5.rows(1).cells(5).width
				.Cells(5).border =0
				.Cells(6).Width = table5.rows(1).cells(6).width
				.Cells(6).border =0
				.Cells(7).Width = table5.rows(1).cells(7).width
				.Cells(7).border =0
				.Cells(8).Width = table5.rows(1).cells(8).width
				.Cells(8).border =0
				.height=20
			End With
			if rs("cig")<>"" or rs("mepa_testo")<>"" then		
				Set Row = Table4.Rows.Add(10) ' row height
				Row.Cells(1).border =0
				Row.Cells(2).border =0
				Row.Cells(3).border =0
				Row.Cells(4).border =0
				Row.Cells(5).border =0
				Row.Cells(6).border =0
				Row.Cells(7).border =0
				Row.Cells(8).border =0
				if rs("cig")<>"" then
					testo_riga=" CIG: "&rs("cig")
				end if
				if rs("mepa_testo")<>"" then
					testo_riga=testo_riga&" "&mepa_tipo(rs("mepa_tipo"))&":"&rs("mepa_testo")
				end if
				Row.Cells(2).AddText testo_riga, "alignment=0; size="&size_chr_pdf&"; Expand=true;"
			end if
			
			
			
			table4.border=0
			
			sconto_totale=0
			totord=0
			n=2
			
			

'			sql_ordine=get_sql_ordine_fattura(rs("tipo_fattura"),idfat)
'			
'			
'			sconto_totale=0
'			set rs_ordini=conn.execute(sql_ordine)
'			do while not rs_ordini.eof
'					'Set Row = Table4.Rows.Add(10) ' row height
'					'Row.Cells(2).AddText "TOTORD"&totord, "alignment=0; size=7; Expand=true;"
'					
'		        if rs_ordini("tipo_documento")="ddt" then
'			        tipo_ddt=rs_ordini("tipo")
'					Set Row = Table4.Rows.Add(10) ' row height
'					Row.Cells(1).border =0
'					Row.Cells(2).AddText "<b>Rif. d.d.t. "&rs_ordini("nddt")&" del "&formatDateTime(rs_ordini("dataddt"), vbShortDate)&"</b>" , "alignment=0; size=7; Expand=true; html=true;"
'					Row.Cells(2).border =0
'					Row.Cells(3).border =0
'					Row.Cells(4).border =0
'					Row.Cells(5).border =0
'					Row.Cells(6).border =0
'					Row.Cells(7).border =0
'					Row.Cells(8).border =0
'				end if
'				
'				
'				
'				if rs_ordini("tipo")<>"2" then
'					Set Row = Table4.Rows.Add(10) ' row height
'					Row.Cells(2).border =0
'					Row.Cells(3).border =0
'					Row.Cells(4).border =0
'					Row.Cells(5).border =0
'					Row.Cells(6).border =0
'					Row.Cells(7).border =0
'					Row.Cells(8).border =0
'					Row.Cells(1).border =0
'					Row.Cells(2).AddText "<b>Rif. ordine "&rs_ordini("nord")&" del "&formatDateTime(rs_ordini("dataordine"), vbShortDate) &"</b>", "alignment=0; size=7; Expand=true; html=true;"
'				end if
'				if rs_ordini("annotazioni")<>"" then
'					Set Row = Table4.Rows.Add(10) ' row height
'					Row.Cells(1).border =0
'					Row.Cells(2).AddText "Annotazioni "&rs_ordini("annotazioni") , "alignment=0; size=7; Expand=true;"
'					Row.Cells(2).border =0
'					Row.Cells(3).border =0
'					Row.Cells(4).border =0
'					Row.Cells(5).border =0
'					Row.Cells(6).border =0
'					Row.Cells(7).border =0
'					Row.Cells(8).border =0
'				end if
'				sconto_ordine=cdbl(rs_ordini("sconto_ordine"))
'				sconto_totale=sconto_totale+sconto_ordine
				
				
				
			'if rs_ordini("tipo_documento")="ddt" then
			'	if tipo_ddt="1" then
			'		sql_dettaglio="select ordini_dett.*,ddt_dett_ordini.quantita, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara inner join ddt_dett_ordini on ordini_dett.iddett = ddt_dett_ordini.iddett  where ddt_dett_ordini.idddt="&rs_ordini("idddt")&" order by ordine,iddett;"
			'	else
			'		sql_dettaglio="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara  where idord="&rs_ordini("sub_idord")&" order by ordine,iddett;"
			'	end if
			'else
					sql_dettaglio="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara  where idord="&idord&" order by ordine,iddett;"
				
			'end if
			set rs1=conn.execute(sql_dettaglio)
			Do While Not rs1.EOF
				testo=""
				Set Row = Table4.Rows.Add(10) ' row height
				Row.Cells(1).border =0
				if rs1("codice_ordine")<>"" then
					Row.Cells(1).AddText codice_articolo_e_variante( rs1("codice_ordine"),rs1("codicevara")), "alignment=1; size="&size_chr_pdf&"; Expand=true; html=true;"
				end if
		    	if rs1("varianti_ordine")<>"" then testo=testo&vbcrlf&rs1("varianti_ordine")
				if cint(rs1("numeri_di_serie"))>0 then
					testo=testo&vbcrlf&elenca_seriali(rs1("iddett"),0)
				end if
				Row.Cells(2).AddText rs1("articolo_ordine")&testo, "alignment=0; size="&size_chr_pdf&"; Expand=true;"
				Row.Cells(2).border =0
				if rs1("um")<>"" then
					Row.Cells(3).AddText rs1("um"), "alignment=2; size="&size_chr_pdf&"; Expand=true;"
				end if
				Row.Cells(3).border =0
				if cdbl(rs1("quantita"))>0 then Row.Cells(4).AddText rs1("quantita"), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
				Row.Cells(4).border =0
				nascondi=true
				if cdbl(rs1("prezzo"))<>0 then
					Row.Cells(5).AddText formatcurrency(rs1("prezzo"),2), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
					nascondi=false
				end if
				Row.Cells(5).border =0
				testo=""
				if cdbl(rs1("sconto_prodotto"))<>0 then
					Row.Cells(6).AddText formatnumber(rs1("sconto_prodotto"),2)&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"
				end if
				Row.Cells(6).border =0
				Row.Cells(7).border =0
				
	            if vecchio_conteggio then
		            
		            prezzo=cdbl(rs1("prezzo"))
					totord=totord+round((prezzo-(cdbl(rs1("sconto_prodotto"))/100)*prezzo)*rs1("quantita"),2)
		        else
					totale_riga_tmp=totale_riga(rs1("prezzo"),rs1("sconto_prodotto"),rs1("quantita"))
				end if
				
				
				if nascondi=false then
					Row.Cells(7).AddText formatcurrency(totale_riga_tmp,2), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
					Row.Cells(8).AddText rs("iva_ordine")&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"
				end if
				Row.Cells(8).border =0
				totord=totord+totale_riga_tmp
					'Set Row = Table4.Rows.Add(10) ' row height
					'Row.Cells(2).AddText "TOTORDB"&totord, "alignment=0; size=7; Expand=true;"
				rs1.movenext
				n=n+1
			Loop
			set rs1= nothing
			trasporto_ordine=cdbl(rs("trasporto"))
			if trasporto_ordine>0 then
				Set Row = Table4.Rows.Add(10) ' row height
				Row.Cells(1).border =0
				Row.Cells(2).AddText "Trasporto", "alignment=0; size=7; Expand=true;"
				Row.Cells(2).border =0
				Row.Cells(3).border =0
				Row.Cells(4).border =0
				Row.Cells(5).AddText formatcurrency(rs("trasporto"),2), "alignment=1; size=7; Expand=true;"
				Row.Cells(5).border =0
				Row.Cells(6).border =0
				Row.Cells(7).border =0
				Row.Cells(7).AddText formatcurrency(rs("trasporto"),2), "alignment=1; size=7; Expand=true;"
				Row.Cells(8).border =0
				Row.Cells(8).AddText rs("iva_ordine")&"%", "alignment=1; size=7; Expand=true;"
				totord=totord+trasporto_ordine
			end if
			if sconto_ordine>0 then
				Set Row = Table4.Rows.Add(10) ' row height
				Row.Cells(1).border =0
				Row.Cells(2).AddText "Sconto ordine", "alignment=0; size=7; Expand=true;"
				Row.Cells(2).border =0
				Row.Cells(3).border =0
				Row.Cells(4).border =0
				Row.Cells(5).AddText formatcurrency(rs("sconto_ordine"),2), "alignment=1; size=7; Expand=true;"
				Row.Cells(5).border =0
				Row.Cells(6).border =0
				Row.Cells(7).border =0
				Row.Cells(7).AddText formatcurrency(rs("sconto_ordine"),2), "alignment=1; size=7; Expand=true;"
				Row.Cells(8).border =0
				Row.Cells(8).AddText rs("aliquota_iva")&"%", "alignment=1; size=7; Expand=true;"
				sconto_totale=sconto_totale+sconto_ordine
			end if
			'rs_ordini.movenext
			'loop
			'set rs_ordini = Nothing
			
			if mod_larochelle then 
				Set Row = Table4.Rows.Add(10) ' row height
				Row.Cells(1).border =0
				Row.Cells(2).AddText txt_restituzione, "alignment=0; html=true; size=8; Expand=true;"
				Row.Cells(2).border =0
				Row.Cells(3).border =0
				Row.Cells(4).border =0
				Row.Cells(5).border =0
				Row.Cells(6).border =0
				Row.Cells(7).border =0
				Row.Cells(8).border =0
			end if 
			
			
			
			
			Set T_totali = Doc.CreateTable("width="&page.width-40&"; height=100; Rows=4; Cols=5; Border=1; CellSpacing=-1; cellpadding=1 ")
			T_totali.Font = Doc.Fonts("Arial")
			With T_totali.Rows(1)
				.Cells(1).Width = (page.width-40)/5+0.75
				.Cells(1).AddText "SPESE BANCARIE", "alignment=2; size=7; Expand=true;"
				.cells(1).border=0
				.Cells(2).AddText "IMPONIBILE", "alignment=2; size=7; Expand=true;"
				.cells(2).SetBorderParams "bottom=False" 
				.Cells(2).Width = .Cells(1).Width
				.Cells(3).AddText "IVA %", "alignment=2; size=7; Expand=true;"
				.Cells(3).Width = .Cells(1).Width
				.cells(3).SetBorderParams "bottom=False" 
				.Cells(4).AddText "IMPOSTA", "alignment=2; size=7; Expand=true;"
				.Cells(4).Width = .Cells(1).Width
				.cells(4).SetBorderParams "bottom=False" 
				.Cells(5).Width = .Cells(1).Width
				'.Cells(5).AddText "TOTALE MERCE", "alignment=2; size=7; Expand=true;"
				.cells(5).SetBorderParams "bottom=False" 
				.height=14
			End With
			With T_totali.Rows(2)
				'.Cells(1).AddText "", "alignment=2; size=7; Expand=true;"
				.cells(1).border=0
				.cells(2).SetBorderParams "Top=False" 
				.cells(2).rowspan=3
				.cells(3).rowspan=3
				.cells(3).SetBorderParams "Top=False" 
				.cells(4).rowspan=3
				.cells(4).SetBorderParams "Top=False" 
				.cells(5).SetBorderParams "Top=False" 
				.height=19
			End With
			With T_totali.Rows(3)
				if sconto_ordine>0 then
					.Cells(1).AddText "SCONTO SU ORDINE", "alignment=2; size=7; Expand=true;"
				end if
				.cells(1).SetBorderParams "bottom=False" 
				.cells(5).SetBorderParams "bottom=False" 
				.Cells(5).AddText "TOTALE FATTURA", "alignment=2; size=7; Expand=true;"
				.height=14
			End With
			With T_totali.Rows(4)
				.Cells(1).AddText formatcurrency(sconto_totale,2), "alignment=1; size=7; Expand=true;"
				.cells(1).SetBorderParams "Top=False" 
				.cells(5).SetBorderParams "Top=False" 
				.height=19
			End With
			'TABELLA 8
			Set Table8 = Doc.CreateTable("width="&page.width-40&"; height=30; Rows=1; Cols=1; Border=1; CellSpacing=-1; cellpadding=2 ")
			Table8.Font = Doc.Fonts("Arial")
			With Table8.Rows(1)
				.Cells(1).Width = (page.width-40)
			End With
			param.clear
			Param("x") = 20
			Param("y") = page.height-204
			Param("MaxHeight") = 455
			FirstRow = 2
			Do While True
				LastRow = Page.Canvas.DrawTable (Table4, param)
			   if LastRow >= Table4.Rows.Count Then Exit Do ' entire table displayed
			   ' Display remaining part of table on the next page
			   Set Page = Page.NextPage
			   Param.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
			   Param("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
			   FirstRow = LastRow + 1
			Loop
			pagine=doc.pages.count
			for pagina=1 to pagine
				Set Page = Doc.Pages.item(pagina)
				table1.rows(1).Cells(3).cleartext
				table1.rows(1).Cells(3).AddText "Pagina "&pagina&" di "&pagine, ""
				Page.Canvas.DrawTable Table1, "x="& page.width/2 &", y="&page.height-20
				if pagina=pagine then
					totord=round(totord,2)
					call T_totali_compila(t_totali, totord, sconto_ordine, rs("trattamento_iva_ordine"), rs("iva_ordine"), spese_bancariev,0)
			
					
					'T_totali.Rows(2).Cells(5).AddText formatcurrency(totord,2), "alignment=1; size=10; Expand=true;" 'TOTALE MERCE
					'if spese_bancariev>0 then
					'	T_totali.Rows(2).Cells(1).AddText formatcurrency(spese_bancariev,2), "alignment=1; size=8; Expand=true;" 'SPESE BANCARIE
					'	totord=totord+spese_bancariev
					'end if
					'T_totali.Rows(2).Cells(2).AddText formatcurrency(totord,2), "alignment=1; size=8; Expand=true;" 'IMPONIBILE
					'if esenzione_iva(rs("trattamento_iva")) then
					'	valore_iva=0
					'	if rs("trattamento_iva")=10 then
					'		valore_iva=formatcurrency(totord*rs("aliquota_iva")/100,2)&vbcrlf&vbcrlf&"-"&formatcurrency(totord*rs("aliquota_iva")/100,2)
					'	end if
					'else
					'	valore_iva=roundup(totord*rs("aliquota_iva")/100,2)
					'	totord=totord+valore_iva
					'	valore_iva=FormatCurrency(valore_iva,2)
					'end if
					'testo_esenzione=replace(trattamento_iva(rs("trattamento_iva")),"Normale","")
					
					'T_totali.Rows(2).Cells(3).AddText rs("aliquota_iva")&"%"&vbcrlf&testo_esenzione, "alignment=2; size=8; Expand=true;"
					'T_totali.Rows(2).Cells(4).AddText valore_iva, "alignment=1; size=8; Expand=true;" 'IVA
					'totord=totord-sconto_totale
					'T_totali.Rows(4).Cells(5).AddText formatcurrency(totord,2), "alignment=1; size=10; Expand=true;" 'TOTALE FATTURA
					
					'Set ccausale_ddt = new cl_causale_ddt
					'CAUSALE=ccausale_ddt.elemento(rs("causale"))
					'set ccausale_ddt = Nothing
					'porto=porto_ddt(rs("porto"))
					'ASPETTO=imballo_ddt(rs("imballo"))
					'peso=rs("peso")
					'dimensione=rs("dimensione")
					'colli=rs("colli")
					'annotazioni=ucase(rs("annotazioni"))
					'if rs("regione")=21 then
					'	annotazioni="<b>Remark:   The arrived amount should be "&simbolo_valuta&FormatNumber(totale_fattura,2)&" exactly.</b>"
					'end if
					'if rs("incaricato")=3 then
					'	INCARICATO=ucase(rs("vettore"))
					'else
				'		INCARICATO=ucase(incaricato_ddt(rs("incaricato")))
					'end if
					'if isdate(rs("data_inizio")) then
					'	inizio= formatDateTime(rs("data_inizio"), vbShortDate)
					'else
					'	inizio=""
					'end if
				end if
				
				
				'Cerco le scadenze
				set rs_scadenze=conn.execute("select * from scadenze where idord="&idord&" order by data")
				txt_scadenze=""
				do while not rs_scadenze.EOF
					call concatena_stringa(txt_scadenze," - ",rs_scadenze("data")&" importo "&simbolo_valuta&FormatNumber(rs_scadenze("importo"),2))
					rs_scadenze.movenext
				loop
				if txt_scadenze<>"" then txt_scadenze="<br>"&txt_scadenze
				Table8.Rows(1).Cells(1).AddText "Scadenze "&txt_scadenze, "alignment=2; size=8; Expand=true; html=true;"
				Page.Canvas.DrawTable Table1, "x="& page.width/2 &", y="&page.height-20
				Page.Canvas.DrawTable Table2, "x="& page.width/2 &", y="&page.height-43
				Page.Canvas.DrawTable Table21, "x=20, y="&page.height-144
				Page.Canvas.DrawTable Table5, "x=20, y="&page.height-203
				Page.Canvas.DrawText "CONTRIBUTO CONAI ASSOLTO, OVE DOVUTO", "X=85; Y="&Page.Height-660&"; size=7;" , doc.fonts("Arial")
				if rs("trattamento_iva_ordine")=2 then
					Page.Canvas.DrawText trattamento_iva(2), "X=85; Y="&Page.Height-550&"; size=7;" , doc.fonts("Arial")
				end if
				Page.Canvas.DrawTable T_totali, "x=20, y="&page.height-675
				Page.Canvas.DrawTable Table8, "x=20, y="&page.height-740
				call intestazione_pdf_ddt_fat(doc, page,rs("data"))
			next
		'Parte conclusiva comune
		
		nome_file= "Nota_di_credito_"&pre_fattura(rs("data"))&nord&".pdf" 
		
		
		if out="pdf" then
			Doc.SaveHttp "attachment;filename="&nome_file
			pdf_fattura=""
		else
			cartella_fatture="\public\fatturepdf"
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			If not objFSO.FolderExists(Server.MapPath(cartella_fatture)) Then
				set f=objFSO.CreateFolder(Server.MapPath(cartella_fatture))
				set f=nothing
			End if 
			set objFSO=nothing
			pdf_fattura = Doc.Save( Server.MapPath(cartella_fatture&"\"&nome_file), True )
		end if
		
		'call CheckConnChiusa()
	else	'not rs.eof
		call esci(6,idfat)
	end if 	'not rs.eof
	Set rs = nothing
end function
%>
