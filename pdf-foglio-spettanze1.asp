<%
'Verifica chiusure 30_11_2015
'Migliorabile con getrows su rs
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="jsonObject.class.asp"-->

<%
	
	
iduser=session("iduser")
if utente_admin and request("iduser")<>"" then
	iduser=request("iduser")
end if
'if not isnumeric(iduser) or iduser="" then
'	call esci(0,0)
'end if

set rs=conn.execute("select * FROM impostazioni")
rs.movefirst
fondo_listino=rs("fondo_listino")
rs.close



	const margine_x=15
	Set Pdf = Server.CreateObject("Persits.Pdf")
	Set Doc = Pdf.CreateDocument
	Set TextParam = PDF.CreateParam
	Set ImageParam = PDF.CreateParam
	Doc.Title = "Richiesta fornitura "&nord
	Doc.Creator = nomesito
	Bordo_sx=25
	Set Page = Doc.Pages.Add (792,612)
	pos_Y=0
	Set fs = Server.CreateObject("Scripting.filesystemObject")
	const file_logo="/public/files/logo_documenti.jpg"
	Set fs = Server.CreateObject("Scripting.filesystemObject")
	if fs.FileExists(Server.MapPath( file_logo)) then
		esiste_logo=true
		
		'Parte nuova
		dimensione_logo=100
		margine=5
		Set logo = Doc.OpenImage( Server.MapPath( file_logo) )
		'calcolo misure dell'immagine considerando la risoluzione DPI
		larghezza=logo.width*72/logo.resolutionx
		altezza=logo.height*72/logo.resolutiony
		
		
		scala=(dimensione_logo)/larghezza
		
		'calcolo nuova altezza considerando la scala per centrare poi l'immagine
		nuova_altezza=altezza*scala
		nuova_larghezza=larghezza*scala
		
		
		param_x =((Page.Width/ 2)-nuova_larghezza)/2
		param_x=margine_x
		'Sotto laschio metà del margine
		param_y = Page.Height - nuova_altezza - 10
		
		
		param_ScaleX = scala
		param_ScaleY = scala
		parametri_logo="x="&param_x&"; y="&param_y&"; ScaleX="&param_ScaleX&"; ScaleY="&param_ScaleY&";"
		parametri_logo=replace(parametri_logo,",",".")
	else
		param_Y=Page.Height-20
	end if
	Set fs = Nothing


	denominazioneV=get_denominazione(iduser)
	
	
	tabella_titolo_Y=page.height-15
	tabella_titolo_X=nuova_larghezza+margine_x
	Set tabella_titolo = Doc.CreateTable("width="&int(page.width-(margine_x*2)-nuova_larghezza)&"; height=600; Rows=2; Cols=1; Border=1; CellSpacing=-1; cellpadding=0 ")
	tabella_titolo.Font = Doc.Fonts("Arial")
	With tabella_titolo.Rows(1)
		'.BGColor = &H000000
		.Cells(1).AddText "RICHIESTA DI FORNITURA", "size=10; alignment=2;"
		.Cells(1).SetBorderParams "bottom=False"
		.height=25
	End With
	With tabella_titolo.Rows(2)
		'.BGColor = &H000000
		.Cells(1).AddText ucase(denominazioneV), "size=12; alignment=2;"
		.Cells(1).SetBorderParams "top=False"
		.height=25
	End With
	
	
	set rs=conn.execute ("select utenti_foglio_spettanze.* from utenti_foglio_spettanze where iduser="&iduser)

	tabella_accessori_Y=Page.Height-20-tabella_titolo.height
	width_tabelle=int(page.width-(margine_x*2))
	Set tabella__accessori = Doc.CreateTable("width="&width_tabelle&"; height=600; Rows=1; Cols=4; Border=0; CellSpacing=-1; cellpadding=0 ")
	tabella__accessori.Font = Doc.Fonts("Arial")
	parametri="size=8;"
	With tabella__accessori.Rows(1)
		'.BGColor = &H000000
		
		if rs("tipo_ordine")=0 then
			.Cells(1).AddText "TRATTATIVA DIRETTA MePA"&vbcrlf&"con PREZZO A CORPO",  parametri&" alignment=2; VAlignment=2;"
		else
			.Cells(1).AddText "ORDINE DIRETTO"&vbcrlf&"FUORI MePA",  parametri&" alignment=1; VAlignment=2;"
		end if
		'.Cells(1).SetBorderParams "bottom=False"

		.Cells(2).AddText "CIG:"&vbcrlf&" "&rs("cig"),  parametri&" alignment=2; VAlignment=2;"

		.Cells(3).AddText "Determinazione:"&vbcrlf&" "&rs("determinazione"),  parametri&" alignment=2; VAlignment=2;"

		.Cells(4).AddText "Completato:"&vbcrlf&rs("completo"),  parametri&" alignment=2; VAlignment=2;"
		

		.height=30
		note=rs("note")
		note=""
		if note<>"" then
			Set Row = tabella__accessori.Rows.Add(10) ' row height
			with row
				.Cells(1).AddText note,  parametri
				.cells(1).colspan=4
				.height=25

			End With
		end if
	End With




	'Carico elenco dipendenti
    set rs_dipendenti=conn.execute ("(select dipendenti.iddip, cognome, nome,sesso from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.dimesso=0 and utenti_dipendenti.iduser="&iduser&" order by sesso desc, cognome,nome)  order by sesso desc, cognome, nome")
    arrDipendenti=rs_dipendenti.getrows()
    ubound_arrDipendenti=ubound(arrDipendenti,2)

	dim array_totali
	redim array_totali(ubound_arrDipendenti)
	for n = 0 to ubound_arrDipendenti
		array_totali(n)=0
	Next	
	
	
	
	
	tabellaArticoli_Y=param_y-10
	Set tabellaArticoli = Doc.CreateTable("width="&width_tabelle&"; height=600; Rows=1; Cols="&ubound_arrDipendenti+4&"; Border=1; CellSpacing=-1; cellpadding=0 ")
	tabellaArticoli.Font = Doc.Fonts("Arial")
	parametri_titolo="size=6; alignment=2;"
	With tabellaArticoli.Rows(1)
		.Cells(1).AddText "Articoli", "size=8; alignment=2; VAlignment=2;"

		'.Cells(1).AddText L&"/"&ubound_arrDipendenti+3, "size=8; alignment=2; VAlignment=2;"

		.Cells(2).AddText "Prezzo", "size=8; alignment=2; VAlignment=2;"
		.Cells(2).width=45




		for n = 0 to ubound_arrDipendenti
		'.BGColor = &H000000
		'.Cells(n).Canvas.SetCTM 0, 1, -1, 0, 15, 5
		.Cells(n+3).width=25
		.Cells(n+3).canvas.drawtext nominativo_breve(arrDipendenti(1,n),arrDipendenti(2,n)), "X=4; Y=3; size=7; Angle=90; " , doc.fonts("Arial")
		next
		n=n+3
		.Cells(n).width=25
		.Cells(n).canvas.drawtext "TOTALE", "X=7; Y=3; size=7; Angle=90;" , doc.fonts("Arial")

		.height=100

	'calcolo larghezza celle
		L=0
		for n = 2 to tabellaArticoli.colnum
			L=L+.Cells(n).width
		next
		.Cells(1).width=width_tabelle-L
		
		
	End With
	
	'call add2log("L:"&l&" width_tabelle:"&width_tabelle&" tabellaArticoli.width:"&tabellaArticoli.width&" L_articoli:"&L_articoli&" tabellaArticoli.colnum:"&tabellaArticoli.colnum,0)


	
	'larghezza_tabella=tabellaArticoli.width+(margine_x*2)
	'if page.width<larghezza_tabella then
	'	page.width=larghezza_tabella
	'end if
	
	parametri_titolo="size=5; alignment=2;"

	sql="select preferiti.idpre, prodotti.idpro, prodotti.codice, prodotti.articolo, prodotti.prezzo, nome_settore from preferiti INNER JOIN prodotti on preferiti.idpro=prodotti.idpro left join settori on prodotti.idcat= settori.idsettore INNER JOIN (select distinct idpro from  utenti_spettanze where iduser="&iduser&") as t on preferiti.idpro=t.idpro  where preferiti.iduser="&iduser&" order by codice"
	set rs_prodotti=conn.execute(sql)
	importoTotale=0
	do while not rs_prodotti.eof 
	
		Set Row = tabellaArticoli.Rows.Add(30) ' row height
		with row

	
		articolo=rs_prodotti("codice")&" "&rs_prodotti("articolo")
		prezzo=cdbl(rs_prodotti("prezzo"))

		quantita_tot=0
		'Loop su dipendenti
		quantita_riga=0

	    for n = 0 to ubound_arrDipendenti
			quantita=0

	
	        'if isnull(rs_prodotti("idpro")    ) then
			'	set rs_spettanze=conn.execute("select ordini_dett_spettanze.*,ordini_dett.iddett, ordini_dett.varianti_ordine from ordini_dett_spettanze inner join ordini_dett on ordini_dett_spettanze.iddett = ordini_dett.iddett where ordini_dett.iddett= "&rs_prodotti("iddett")&" and iddip="&arrDipendenti(0,n)&" and idord="&idord)
			'else
			'	set rs_spettanze=conn.execute("select ordini_dett_spettanze.*,ordini_dett.iddett, ordini_dett.varianti_ordine from ordini_dett_spettanze inner join ordini_dett on ordini_dett_spettanze.iddett = ordini_dett.iddett where ordini_dett.idpro= "&rs_prodotti("idpro")&" and iddip="&arrDipendenti(0,n)&" and idord="&idord)
			'end if
			sql="select utenti_spettanze.* from utenti_spettanze where iduser="&iduser&" and idpro= "&rs_prodotti("idpro")&" and iddip="&arrDipendenti(0,n)
			set rs_spettanze=conn.execute(sql)
			if not rs_spettanze.eof then
				
				
				
				
				val=0
				classe=""
				quantita=cint(rs_spettanze("quantita"))
				.Cells(n+3).AddText quantita, "size=8; alignment=2; VAlignment=2;"
				array_totali(n)=array_totali(n)+quantita*prezzo
				quantita_riga=quantita_riga+quantita
			end if
        Next	
			.Cells(1).AddText articolo, "size=7; html=true; VAlignment=2;"
			.cells(1).CellPadding=2
			.Cells(2).AddText FormatNumber(prezzo,2), "size=8; alignment=1; VAlignment=2;"
			.cells(2).CellPadding=2
	
			if quantita_riga>0 then
				.Cells(n+3).AddText quantita_riga,"size=8; alignment=2; VAlignment=2;"
				importoTotale=importoTotale+(quantita_riga*prezzo)
	
			end if
		
		
	end with
	rs_prodotti.MoveNext
	Loop
	
	set rs_prodotti = Nothing
	'TRASPORTO
	Set Row = tabellaArticoli.Rows.Add(30) ' row height
	with row
		.cells(1).CellPadding=2
		.Cells(1).AddText "Trasporto", "size=7; html=true; VAlignment=2;"
		.cells(2).CellPadding=2
		.Cells(2).AddText FormatNumber(12,2), "size=8; alignment=1; VAlignment=2;"
			importoTotale=importoTotale+12
	end with
	'TOTALI PER DIPENDENTE
	Set Row = tabellaArticoli.Rows.Add(30) ' row height
	with row
		.cells(1).colspan=2
    		.cells(1).CellPadding=2
		
		.Cells(1).AddText "Totale senza iva per dipendente", "size=8; alignment=1; VAlignment=2;"

	    for n = 0 to ubound_arrDipendenti
    		.cells(n+3).CellPadding=0

			.Cells(n+3).AddText formatnumber(array_totali(n),2), "size=6; alignment=1; VAlignment=2; bold=true;"
        Next	
	end with
	'Totale iva esclusa:
	Set Row = tabellaArticoli.Rows.Add(30) ' row height
	with row
		.cells(1).CellPadding=2
		.Cells(1).AddText "<b>Totale iva esclusa</b>", "size=8; html=true; VAlignment=2; html=true;"
		.Cells(2).AddText FormatNumber(importoTotale,2), "size=8; alignment=1; VAlignment=2;"
		.cells(2).CellPadding=2

		.cells(3).colspan=ubound_arrDipendenti+4

			importoTotale=importoTotale
	end with
	
	'IVA
	valore_iva=RoundUp(importoTotale*iva(date())/100, 2)
	Set Row = tabellaArticoli.Rows.Add(30) ' row height
	with row
		.cells(1).CellPadding=2
		.Cells(1).AddText "<b>Iva "&iva(date())&"%</b>", "size=8; html=true; VAlignment=2; html=true;"
		.Cells(2).AddText FormatNumber(valore_iva,2), "size=8; alignment=1; VAlignment=2;"
		.cells(2).CellPadding=2
		.cells(3).colspan=ubound_arrDipendenti+4
	end with
	
	'TOTALE CON IVA
	totale=importoTotale+valore_iva
	Set Row = tabellaArticoli.Rows.Add(30) ' row height
	with row
		.cells(1).CellPadding=2
		.Cells(1).AddText "<b>Totale con iva:</b>", "size=8; html=true; VAlignment=2; html=true;"
		.Cells(2).AddText FormatNumber(totale,2), "size=8; alignment=1; VAlignment=2;"
		.cells(2).CellPadding=2
		.cells(3).colspan=ubound_arrDipendenti+4
	end with


'Immagine timbro
	Set fs = Server.CreateObject("Scripting.filesystemObject")
	file_timbro=Server.MapPath("/images/Timbro_e_firma.png")
	if fs.FileExists(file_timbro ) then
		'Parte nuova
		dimensione_logo=100
		margine=5
		Set timbro = Doc.OpenImage( file_timbro )
		'calcolo misure dell'immagine considerando la risoluzione DPI
		larghezza=timbro.width*72/timbro.resolutionx
		altezza=timbro.height*72/timbro.resolutiony
		
		
		scala=(dimensione_logo)/larghezza
		
		'calcolo nuova altezza considerando la scala per centrare poi l'immagine
		nuova_altezza=altezza*scala
		nuova_larghezza=larghezza*scala
		
		'Centrato
		'param_x =((Page.Width/ 2)-nuova_larghezza)/2
		
		'A sinistra
		'param_x=margine_x
		
		'A destra
		param_x = Page.Width - margine_x - nuova_larghezza
		
		'Sotto laschio metà del margine
		param_y = Page.Height - nuova_altezza - 5
		
		
		param_ScaleX = scala
		param_ScaleY = scala

	end if
	set fs = Nothing

'Tabella footer
	Set tabellaFooter = Doc.CreateTable("width="&width_tabelle&"; height=25; Rows=1; Cols=2; Border=1; CellSpacing=-1; cellpadding=1 ")
	tabellaFooter.Font = Doc.Fonts("Arial")
	Set Row = tabellaFooter.Rows(1) ' row height
	with row
		.Cells(1).AddText "VISTO SI CONFERMA LA PRESENTE FORNITURA", "size=9; alignment=2;"
		.height=25
		.Cells(1).SetBorderParams "Bottom=false;"
	end with
	Set Row = tabellaFooter.Rows.Add(30) ' row height
	with row
		'.cells(1).CellPadding=2
		.Cells(1).AddText "Timbro"&vbcrlf&"Ente", "size=7; alignment=2;  VAlignment=2; "
		.height=30
		.Cells(1).SetBorderParams "top=false; bottom=false;"

	end with
	Set Row = tabellaFooter.Rows.Add(30) ' row height
	with row
		'.cells(1).CellPadding=2
		.Cells(1).AddText "    DATA:", "size=9; alignment=0;  VAlignment=2; "
		.height=30
		.Cells(1).SetBorderParams "top=false;"

	end with
	'tabellaFooter.rows(1).Cells(2).rowspan=4
	tabellaFooter.rows(1).Cells(2).AddText "Fatti salvi casi particolari preventivamente autorizzati; la merce viene consegnata franco LA ROCHELLE SNC che anticiper&agrave; le spese di spedizione e fatturer&agrave; al cliente con la merce. LA ROCHELLE SNC non sar&agrave; tenuta alla sostituzione della merce in caso di evidente errore nell&rsquo;ordine da parte del cliente. Eventuali contestazioni per vizi o difformit&agrave; delle merci dovranno pervenire alla LA ROCHELLE SNC entro 30 giorni dalla consegna oltre i quali LA ROCHELLE SNC non &egrave; pi&ugrave; tenuta a sanare i difetti riscontrati nella fornitura. Il pagamento delle merci dovr&agrave; avvenire entro 30 gg. da fine mese dalla data della fattura.", "size=8; alignment=0; expand=true; html=true;"
	tabellaFooter.rows(1).Cells(2).SetBorderParams "bottom=false;"

	tabellaFooter.rows(2).Cells(2).AddText "<b>Per rispondere al meglio alle vostre necessit&agrave; ed alle vostre esigenze, la rilevazione delle taglie potr&agrave; essere eseguita presso i vostri uffici da nostro personale<br>Le richieste di riparazioni o sostituzioni del materiale per cambio taglie, si accettano fino ad un mese dalla consegna<br>Per ragioni fiscali, non si garantiscono sostituzioni per errate consegne dovute ad errori di codifica degli articoli da parte del cliente</b>", "size=8; alignment=0; expand=true; html=true;"
	tabellaFooter.rows(2).Cells(2).SetBorderParams "top=false;"
	tabellaFooter.rows(3).Cells(2).AddText "IL COMMITTENTE:", "size=9; alignment=0;  VAlignment=2; "


	
	Set Param = Pdf.CreateParam
	
	param.clear
	Param("x") = margine_x
	y_tabella_articoli=page.height-75-tabella__accessori.height
	Param("y") = y_tabella_articoli
	
	Param("MaxHeight") =page.height-75-40
	FirstRow = 1
	pagina=1
	Do While True
		if pagina=1 then
			start_Y = page.height-75-tabella__accessori.height
		else
			start_Y = page.height-75
		end if
		Param("y") = start_Y
		
		LastRow = Page.Canvas.DrawTable (tabellaArticoli, param)
		
		
	
		
		
		
		
	   if LastRow >= tabellaArticoli.Rows.Count Then
	   	call add2log("start_Y:"&start_Y,0)
	   		'Calcolo pos Y della tabella nell'ultima pagina visualizzata
	   		if pagina>1 then
		   		
		   		start_Y=start_Y-tabellaArticoli.rows(1).height
		   	end if
			for n=FirstRow to LastRow
				start_Y=start_Y-tabellaArticoli.rows(n).height+1
				call add2log("altezza riga "&n&":"&tabellaArticoli.rows(n).height,0)
			next

	   	call add2log("start_Y:"&start_Y,0)
	   
	    Exit Do ' entire table displayed
	    
	   end if
	   ' Display remaining part of table on the next page
	   Set Page = Page.NextPage
	   Param.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
	   Param("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
	   FirstRow = LastRow + 1
	   pagina=pagina+1
	Loop
	Y_timbro=start_Y
	call add2log("Y_timbro:"&Y_timbro,0)
	
	Y_timbro=Y_timbro-nuova_altezza-5 '5 Margine
	if Y_timbro<0 then
		Set Page = Doc.Pages.Add (792,612)
	   pagina=pagina+1
	   Y_timbro=page.height-75-nuova_altezza-5 '5 Margine
	end if
	
	parametri_timbro="x="&param_x&"; y="&Y_timbro&"; ScaleX="&param_ScaleX&"; ScaleY="&param_ScaleY&";"
	parametri_timbro=replace(parametri_timbro,",",".")

	Page.Canvas.DrawImage timbro, parametri_timbro
	
	
	
	Y_tabella_footer=Y_timbro

	
	Y_min_tabella_footer=Y_tabella_footer-tabellaFooter.height
	
	if Y_min_tabella_footer<0 then
		Set Page = Doc.Pages.Add (792,612)
	   pagina=pagina+1
	   Y_tabella_footer=page.height-75
	end if
	
	
	Y_tabella_footer=replace(Y_tabella_footer,",",".")
	Page.Canvas.DrawTable tabellaFooter, "x=20, y="&Y_tabella_footer











	pagine=doc.pages.count
	'call add2log("pagine:"&pagine,0)

	for pagina=1 to pagine
		Set Page = Doc.Pages.item(pagina)
		Page.Canvas.DrawTable tabella_titolo, "x="&tabella_titolo_X&", y="&replace(tabella_titolo_Y,",",".")
		'Page.Canvas.DrawTable tabellaArticoli, "x="&margine_x&", y="&replace(tabellaArticoli_Y,",",".")
		if esiste_logo then
			Page.Canvas.DrawImage logo, parametri_logo
		
		end if
		if pagina=1 then
			Page.Canvas.DrawTable tabella__accessori, "x="&margine_x&", y="&replace(tabella_accessori_Y,",",".")
		end if

		
		
		
		Page.Canvas.DrawText fondo_listino,  "x="&(page.width-350)/2&"; y=25; width=350; alignment=center; size=8;html=true", doc.fonts("Arial")
		
		Page.Canvas.DrawText "Pagina "&pagina&"/"&pagine, "X="&page.width-100&"; Y=20;" , doc.fonts("Arial")

		
	next
	set rs_rilievi = Nothing
	set rs_dipendenti = Nothing
	call connclose()



	Doc.SaveHttp "attachment;filename="&replace(Doc.Title," ","_")&".pdf" 

	Set Doc = Nothing
	Set TextParam = Nothing
	Set ImageParam = Nothing
	Set Param_logo = Nothing
	Set Pdf = Nothing
	
	
	


%>


