<%
'Verifica chiusure 30_11_2015
'Migliorabile con getrows su rs
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="jsonObject.class.asp"-->

<%
	
'---------------------------------------------	
'V3.00 genera più pagine
'---------------------------------------------


idord=request("idord")
if request.querystring("prezzi")<>"" then
	prezzi=true
else
	prezzi=false
end if


set rs_ordine=conn.execute ("select ordini.nord, ordini.iduser from ordini where idord="&idord)	 
iduser=rs_ordine("iduser")
nord=rs_ordine("nord")
dim lettera_pagina()
set rs=conn.execute("select * FROM impostazioni")
rs.movefirst
fondo_listino=rs("fondo_listino")
rs.close






'Preset variabili
start_Pagina=1
importoTotale=0
txtLog=""

const margine_x=15
Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument
Set TextParam = PDF.CreateParam
Set ImageParam = PDF.CreateParam
Doc.Title = "Spettanze ordine "&nord
Doc.Creator = nomesito
Set Page = Doc.Pages.Add (792,612)
pos_Y=0
Set fs = Server.CreateObject("Scripting.filesystemObject")
const file_logo="/public/files/logo_documenti.jpg"
Set fs = Server.CreateObject("Scripting.filesystemObject")
if fs.FileExists(Server.MapPath( file_logo)) then
	
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
	.Cells(1).AddText "ELENCO SPETTANZE ORDINE "&nord, "size=10; alignment=2;"
	.Cells(1).SetBorderParams "bottom=False"
	.height=25
End With
With tabella_titolo.Rows(2)
	'.BGColor = &H000000
	.Cells(1).AddText ucase(denominazioneV), "size=12; alignment=2;"
	.Cells(1).SetBorderParams "top=False"
	.height=25
End With


'Carico elenco dipendenti
    sql="(select dipendenti.iddip, cognome, nome,sesso from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip inner join (select distinct iddip from ordini_dett_spettanze inner join ordini_dett on ordini_dett_spettanze.iddett = ordini_dett.iddett where idord="&idord&") t on dipendenti.iddip = t.iddip  order by sesso desc, cognome,nome) union ( select 0 as iddip , 'Scorta tecnica' as cognome,'' as nome, 'A' as sesso ) order by sesso desc, cognome, nome"


set rs_dipendenti=conn.execute (sql)
arrDipendenti=rs_dipendenti.getrows()
ubound_arrDipendenti=ubound(arrDipendenti,2)

dim array_totali
redim array_totali(ubound_arrDipendenti)
for n = 0 to ubound_arrDipendenti
	array_totali(n)=0
Next	
	
const dipperpagina=16

'Determino cicli

resto= ubound_arrDipendenti Mod dipperpagina


cicli=int(ubound_arrDipendenti/(dipperpagina))

cicliPrima=cicli

if resto>0 then
	cicli=cicli+1
	
end if	

add2log "Calcolo cicli: dipendenti:" &ubound_arrDipendenti&" per pagina:"&dipperpagina& " cicliPrima:"&cicliPrima&" resto:"&resto&" cili dopo: "&cicli,1

'if ubound_arrDipendenti mod dipperpagina>0 then clicli = cicli +1


txtLog=txtLog&" ubound_arrDipendenti:"&ubound_arrDipendenti&" cicli:"&cicli

width_tabelle=int(page.width-(margine_x*2))


	set rs= conn.execute ("select mod_dip from utenti_clienti where iduser="&iduser)
	if not rs.eof then
		jsonString=rs(0)
		if jsonString="" then
		jsonString="{}"
		end if
	else
		jsonString="{}"
	end if
	set rs = nothing
	'jsonString="{""velcrocamicia"":1}"
	'Response.write(jsonString)
	set JSON = New JSONobject
	JSON.Parse(jsonString)
	tabella_accessori_Y=Page.Height-20-tabella_titolo.height

	Set tabella__accessori = Doc.CreateTable("width="&width_tabelle&"; height=600; Rows=1; Cols=19; Border=0; CellSpacing=-1; cellpadding=0 ")
	tabella__accessori.Font = Doc.Fonts("Arial")
	parametri="size=8;"
	With tabella__accessori.Rows(1)
		'.BGColor = &H000000
		.Cells(1).AddText "<b>Velcro DX </b>",  parametri&" alignment=1; html=true;"
		.Cells(1).Width = 48
		'.Cells(1).SetBorderParams "bottom=False"
		.Cells(2).AddText "Camicia",  parametri&" alignment=2;"
		.Cells(2).Width = 38

		.Cells(2).BGColor = verde(json("velcrocamicia"))
		.Cells(3).AddText "Giacca",  parametri&" alignment=2;"
		.Cells(3).Width = 38
		.Cells(3).BGColor = verde(json("velcrocamicia"))

		.Cells(4).AddText "Giubbino",  parametri&" alignment=2;"
		.Cells(4).Width = 38
		.Cells(4).BGColor = verde(json("velcrocamicia"))
		
		.Cells(5).Width = 13
		.Cells(5).SetBorderParams "bottom=False; top=False;"


		.Cells(6).AddText "<b>Cintura e passanti </b>",  parametri&" alignment=1; html=true;"
		.Cells(6).Width = 78

		.Cells(7).AddText "SI",  parametri&" alignment=2;"
		.Cells(7).Width = 22
		.Cells(7).BGColor = verde(json("cintura"))

		.Cells(8).AddText "NO",  parametri&" alignment=2;"
		.Cells(8).Width = 22
		if json("cintura")=0 then
			.Cells(8).BGColor = &H99FF99
		else
			.Cells(8).BGColor = &Hffffff
		end if
		
		.Cells(9).Width = 13
		.Cells(9).SetBorderParams "bottom=False; top=False;"
		
		.Cells(10).AddText "<b>Camicia </b>",  parametri&" alignment=1; html=true;"
		.Cells(11).Width = 42
		
		.Cells(11).AddText "Militare",  parametri&" alignment=2;"
		.Cells(11).BGColor = verde(json("camiciam"))
		.Cells(11).Width = 40


		.Cells(12).AddText "Civile",  parametri&" alignment=2;"
		.Cells(12).BGColor = verde(json("camiciac"))
		.Cells(12).Width = 40
		
		
		.Cells(13).Width = 12
		.Cells(13).SetBorderParams "bottom=False; top=False;"
	
		
		
		.Cells(14).AddText "<b>Camiciotto </b>",  parametri&" alignment=1; html=true;"
		.Cells(14).Width = 60
		
		.Cells(15).AddText "Tutto bottoni",  parametri&" alignment=2;"
		.Cells(15).BGColor = verde(json("camiciottob"))
		.Cells(15).Width = 40

		.Cells(16).AddText "Polo",  parametri&" alignment=2;"
		.Cells(16).BGColor = verde(json("camiciottop"))
		.Cells(16).Width = 40
		
		.Cells(17).Width = 12
		.Cells(17).SetBorderParams "bottom=False; top=False;"
		
		.Cells(18).AddText "<b>Interasse fori berretto</b> ",  parametri&" alignment=1; html=true;"
		.Cells(18).Width = 100
		.Cells(19).AddText json("interasse")&" cm",  parametri
		.Cells(19).Width = 85
		.height=20

		if json("note")<>"" then
			Set Row = tabella__accessori.Rows.Add(10) ' row height
			with row
				.Cells(1).Width = 90
				.Cells(1).AddText json("note"),  parametri
				.cells(1).colspan=19
				.height=25

			End With
		end if
	End With




for ciclo = 1 to cicli
		
	if ciclo>1 then
		Set Page = Doc.Pages.Add (792,612)
	end if
	
	
	
	'Determino coldip: quante colonne dipendenti visualizzare sulla pagina
	if cicli=1 then
		'Solo una pagina
		colDip=ubound_arrDipendenti
	else
		if ciclo<cicli then
			'Pagine intermedie
			colDip=dipperpagina
		else
			'Ultima pagina
			colDip= ubound_arrDipendenti -( (dipperpagina)*(cicli-1))+1	'-1 per scorta tecnica'
			call add2log("coldip ultima pagina:"&coldip&" cili:"&cicli&" ubound_arrDipendenti:"&ubound_arrDipendenti,0)
		end if
		
		enddipp=offsDip(ciclo,colDip)
		if enddipp>= ubound_arrDipendenti then
			enddipp=ubound_arrDipendenti
		end if
		
	
		'nominativiPagina=nominativo_breve(arrDipendenti(1,offsDip(ciclo,0)),arrDipendenti(2,offsDip(ciclo,0)))&" a "&nominativo_breve(arrDipendenti(1,enddipp),arrDipendenti(2,enddipp))

	end if
			call add2log("ernder pagina:"&ciclo &" coldip:"&coldip,0)


	
	L=0
	tabellaArticoli_Y=param_y-10
	Set tabellaArticoli = Doc.CreateTable("width="&width_tabelle&"; height=600; Rows=1; Cols="&colDip+3&"; Border=1; CellSpacing=-1; cellpadding=0 ")
	tabellaArticoli.Font = Doc.Fonts("Arial")
	parametri_titolo="size=6; alignment=2;"
	With tabellaArticoli.Rows(1)
		.Cells(1).AddText "Articoli", "size=8; alignment=2; VAlignment=2;"
		if prezzi then
			.Cells(2).AddText "Prezzo", "size=8; alignment=2; VAlignment=2;"
		else			
			.Cells(1).colspan=2
		end if
		.Cells(2).width=45
		L=L+.Cells(2).width-1

		'Intestazione
		for n = 0 to colDip-1
			'.BGColor = &H000000
			'.Cells(n).Canvas.SetCTM 0, 1, -1, 0, 15, 5
			.Cells(n+3).width=22.5
			L=L+.Cells(n+3).width-1
			colore=dipendente_colorato( arrDipendenti(3,n) )

			
			
			.Cells(n+3).canvas.drawtext nominativo_breve(arrDipendenti(1,offsDip(ciclo,n)),arrDipendenti(2,offsDip(ciclo,n))), "X=4; Y=3; size=7; Angle=90; Color="&colore&"; " , doc.fonts("Arial")
			
	
			
			if offsDip(ciclo,n)>=ubound_arrDipendenti then
				n=n+1
				exit for
				
			end if
			
			
			
		next
		n=n+3
		.Cells(n).width=25
		L=L+.Cells(n).width
		.Cells(n).canvas.drawtext "TOTALE", "X=7; Y=3; size=7; Angle=90;" , doc.fonts("Arial")
		.height=100


'response.write width_tabelle&"-"&L&"<br>"

		.Cells(1).width=width_tabelle-L
	End With
	
	parametri_titolo="size=5; alignment=2;"

	
	sql="select idpro, codice_ordine as codice, articolo_ordine as articolo, prezzo, iddett, quantita from ordini_dett where idord="&idord&" and pronto>0  order by ordine"

	
	
	set rs_prodotti=conn.execute(sql)
	do while not rs_prodotti.eof 
	
		Set Row = tabellaArticoli.Rows.Add(30) ' row height
		with row
			articolo=rs_prodotti("codice")&" "&rs_prodotti("articolo")
			quantita_articolo=rs_prodotti("quantita")

			quantita_tot=0
			quantita_riga=0
			tot_in_ordine=0
			tot_pronto=0
			tot_consegnato=0
			classe_articolo_v=&HFFFFFF
				
				
			'Loop quantita su dipendenti
			for n = 0 to colDip-1
				classe=&HFFFFFF
				quantita=0
				
				
		       ' if isnull(rs_prodotti("idpro")    ) then
					sql = "select ordini_dett_spettanze.*,ordini_dett.iddett,ordini_dett.prezzo, ordini_dett.varianti_ordine from ordini_dett_spettanze inner join ordini_dett on ordini_dett_spettanze.iddett = ordini_dett.iddett where ordini_dett.iddett= "&rs_prodotti("iddett")&" and iddip="&arrDipendenti(0,offsDip(ciclo,n))&" and idord="&idord&" and ordini_dett_spettanze.pronto>0"
				'else
				'	sql = "select ordini_dett_spettanze.*,ordini_dett.iddett,ordini_dett.prezzo, ordini_dett.varianti_ordine from ordini_dett_spettanze inner join ordini_dett on ordini_dett_spettanze.iddett = ordini_dett.iddett where ordini_dett.idpro= "&rs_prodotti("idpro")&" and iddip="&arrDipendenti(0,n)&" and idord="&idord
				'end if
				set rs_spettanze=conn.execute(sql)
				if not rs_spettanze.eof then
					prezzo=cdbl(rs_spettanze("prezzo"))

					val=0
					classe=""
					quantita=cint(rs_spettanze("pronto"))
					tot_in_ordine=tot_in_ordine+rs_spettanze("in_ordine")
					tot_pronto=tot_pronto+rs_spettanze("pronto")
					tot_consegnato=tot_consegnato+rs_spettanze("consegnato")
					txt=avanzamento(quantita,rs_spettanze("in_ordine"),rs_spettanze("pronto"),rs_spettanze("consegnato"),classe)
					
					
					.cells(n+3).BGColor=classe
					.Cells(n+3).AddText quantita, "size=8; alignment=2; VAlignment=2;"
					array_totali(n)=array_totali(n)+quantita*prezzo
					quantita_riga=quantita_riga+quantita
				end if
				
				if offsDip(ciclo,n)>=ubound_arrDipendenti then
									n=n+1

					exit for
					
				end if
	        Next	
			classe_articolo_v=classe_articolo(quantita_articolo,tot_in_ordine,tot_pronto,tot_consegnato)

	        'call add2log("classe_articolo_v:"&classe_articolo_v,0)
			'.cells(1).bgcolor=classe_articolo_v
			.Cells(1).AddText articolo, "size=7; html=true; VAlignment=2;"
			.cells(1).CellPadding=2
			if prezzi then
				.Cells(2).AddText FormatNumber(prezzo,2), "size=8; alignment=1; VAlignment=2;"
				.cells(2).CellPadding=2
			else
				.cells(1).colspan=2
			end if
			if quantita_articolo>0 then
				.Cells(n+3).AddText quantita_articolo,"size=8; alignment=2; VAlignment=2;"
				importoTotale=importoTotale+(quantita_riga*prezzo)
	
			end if
		end with
		rs_prodotti.MoveNext
	Loop
	
	set rs_prodotti = Nothing
	if prezzi then
	'TOTALI PER DIPENDENTE
		Set Row = tabellaArticoli.Rows.Add(30) ' row height
		with row
			.cells(1).colspan=2
			.cells(1).CellPadding=2
			.Cells(1).AddText "Totale senza iva per dipendente", "size=8; alignment=1; VAlignment=2;"

			'totali
			for n = 0 to colDip-1
				.cells(n+3).CellPadding=0
				.Cells(n+3).AddText formatnumber(array_totali(offsDip(ciclo,n)),2), "size=6; alignment=1; VAlignment=2; bold=true;"
				if offsDip(ciclo,n)>=ubound_arrDipendenti then
					n=n+1

				exit for
				
			end if

			Next	
		end with
	end if

	
	
	
	
	
	if ciclo=cicli and prezzi then
		call add2log("importoTotale:"&importoTotale,0)
		set rs_ordine=conn.execute ("select * from ordini where idord="&idord)	 

		'Ultimo ciclo visualizzo totali
		'TRASPORTO
		Set Row = tabellaArticoli.Rows.Add(30) ' row height
		with row
			.cells(1).CellPadding=2
			.Cells(1).AddText "Trasporto", "size=7; html=true; VAlignment=2;"
			.cells(2).CellPadding=2
			.Cells(2).AddText FormatNumber(rs_ordine("trasporto"),2), "size=8; alignment=1; VAlignment=2;"
		end with
		'Totale iva esclusa:
		Set Row = tabellaArticoli.Rows.Add(30) ' row height
		with row
			.cells(1).CellPadding=2
			.Cells(1).AddText "<b>Totale iva esclusa</b>", "size=8; html=true; VAlignment=2; html=true;"
			.Cells(2).AddText FormatNumber(cdbl(rs_ordine("totale_merce_ordine"))+cdbl(rs_ordine("trasporto")),2), "size=8; alignment=1; VAlignment=2;"
			.cells(2).CellPadding=2
	
			.cells(3).colspan=ubound_arrDipendenti+4
	
		end with
		if false then
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
	end if

		'Immagine timbro
		Set fs = Server.CreateObject("Scripting.filesystemObject")
		file_timbro=Server.MapPath("/images/Timbro_e_firma.png")
		if fs.FileExists(file_timbro ) then
			'Parte nuova
			dimensione_logo=260
			margine=5
			Set timbro = Doc.OpenImage( file_timbro )
			'calcolo misure dell'immagine considerando la risoluzione DPI
			larghezza=timbro.width*72/timbro.resolutionx
			altezza=timbro.height*72/timbro.resolutiony
			txtLog=txtLog&"larghezza:"&larghezza&" altezza:"&altezza
			
			scala=(dimensione_logo)/larghezza
			
			'calcolo nuova altezza considerando la scala per centrare poi l'immagine
			nuova_altezza=altezza*scala
			nuova_larghezza=larghezza*scala
			txtLog=txtLog&" nuova_larghezza:"&nuova_larghezza&" nuova_altezza:"&nuova_altezza
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
			'.Cells(1).AddText "VISTO SI CONFERMA LA PRESENTE FORNITURA", "size=9; alignment=2;"
			.height=25
			.Cells(1).SetBorderParams "Bottom=false;"
		end with
		'Set Row = tabellaFooter.Rows.Add(30) ' row height
		'with row
		'	'.cells(1).CellPadding=2
		'	.Cells(1).AddText "Timbro"&vbcrlf&"Ente", "size=7; alignment=2;  VAlignment=2; "
		'	.height=30
		'	.Cells(1).SetBorderParams "top=false; bottom=false;"
		'end with
		'Set Row = tabellaFooter.Rows.Add(30) ' row height
		'with row
		'	'.cells(1).CellPadding=2
		'	.Cells(1).AddText "    DATA:", "size=9; alignment=0;  VAlignment=2; "
		'	.height=30
		'	.Cells(1).SetBorderParams "top=false;"
		'end with
		'tabellaFooter.rows(1).Cells(2).AddText "Fatti salvi casi particolari preventivamente autorizzati; la merce viene consegnata franco LA ROCHELLE SNC che anticiper&agrave; le spese di spedizione e fatturer&agrave; al cliente con la merce. LA ROCHELLE SNC non sar&agrave; tenuta alla sostituzione della merce in caso di evidente errore nell&rsquo;ordine da parte del cliente. Eventuali contestazioni per vizi o difformit&agrave; delle merci dovranno pervenire alla LA ROCHELLE SNC entro 30 giorni dalla consegna oltre i quali LA ROCHELLE SNC non &egrave; pi&ugrave; tenuta a sanare i difetti riscontrati nella fornitura. Il pagamento delle merci dovr&agrave; avvenire entro 30 gg. da fine mese dalla data della fattura.", "size=8; alignment=0; expand=true; html=true;"
		'tabellaFooter.rows(1).Cells(2).SetBorderParams "bottom=false;"
	
		'tabellaFooter.rows(2).Cells(2).AddText "<b>Per rispondere al meglio alle vostre necessit&agrave; ed alle vostre esigenze, la rilevazione delle taglie potr&agrave; essere eseguita presso i vostri uffici da nostro personale<br>Le richieste di riparazioni o sostituzioni del materiale per cambio taglie, si accettano fino ad un mese dalla consegna<br>Per ragioni fiscali, non si garantiscono sostituzioni per errate consegne dovute ad errori di codifica degli articoli da parte del cliente</b>", "size=8; alignment=0; expand=true; html=true;"
		'tabellaFooter.rows(2).Cells(2).SetBorderParams "top=false;"
		'tabellaFooter.rows(3).Cells(2).AddText "IL COMMITTENTE:", "size=9; alignment=0;  VAlignment=2; "
		if false then	'if txtLog<>"" and utente_andrea then
			Set Row = tabellaFooter.Rows.Add(30)
			row.Cells(1).AddText txtLog, "size=8; alignment=0; expand=true; html=true;"
		end if

	end if

	
	Set Param = Pdf.CreateParam
	param.clear
	Param("x") = margine_x
	y_tabella_articoli=page.height-75-tabella__accessori.height
	Param("y") = y_tabella_articoli
	Param("MaxHeight") =page.height-75-100
	FirstRow = 1
	pagina=1
	Do While True
		if pagina=1 and ciclo=1 then
			start_Y = page.height-75-tabella__accessori.height
		else
			start_Y = page.height-75
		end if
		Param("y") = start_Y
		Param("MaxHeight") =start_Y-30	'30 altezza testo dati azienda
		
		LastRow = Page.Canvas.DrawTable (tabellaArticoli, param)
		
		if LastRow >= tabellaArticoli.Rows.Count Then
			'Calcolo pos Y della tabella nell'ultima pagina visualizzata
			if pagina>1 then
				'Considero la prima riga che ha altezza diversa
				start_Y=start_Y-tabellaArticoli.rows(1).height
			end if
			for n=FirstRow to LastRow
				'altezza delle altre righe
				start_Y=start_Y-tabellaArticoli.rows(n).height+1
			next
			redim Preserve lettera_pagina(doc.pages.count)
			'lettera_pagina(doc.pages.count)=chr(ciclo+64)
			lettera_pagina(doc.pages.count)=nominativiPagina
			
			Exit Do ' entire table displayed
		end if
		redim Preserve lettera_pagina(doc.pages.count)
		'lettera_pagina(doc.pages.count)=chr(ciclo+64)
		lettera_pagina(doc.pages.count)=nominativiPagina
		' Display remaining part of table on the next page
		Set Page = Page.NextPage
  
	   Param.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
	   Param("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
	   FirstRow = LastRow + 1
	   
	   
	   pagina=pagina+1
	Loop
	
	if false then 	'Timbto
	
		Y_timbro=start_Y
		Y_timbro=Y_timbro-nuova_altezza-5 '5 Margine
		if Y_timbro<0 then
			Set Page = Doc.Pages.Add (792,612)
			redim Preserve lettera_pagina(doc.pages.count)
			lettera_pagina(doc.pages.count)="Totali"
		   pagina=pagina+1
		   Y_timbro=page.height-75-nuova_altezza-5 '5 Margine
		end if
		
		parametri_timbro="x="&param_x&"; y="&Y_timbro&"; ScaleX="&param_ScaleX&"; ScaleY="&param_ScaleY&";"
		parametri_timbro=replace(parametri_timbro,",",".")
		if IsObject(timbro) then
			Page.Canvas.DrawImage timbro, parametri_timbro
		end if
	
		Page.Canvas.DrawText "<b>Nell'attesa di un vostro cenno di riscontro porgiamo cordiali saluti.</b>",  "x="&margine_x&"; y="&start_Y-10&"; width=350; alignment=center; size=10; html=true", doc.fonts("Arial")
	end if
	
	if ciclo=cicli and false then 'Tabella footer su ultima pagina
		Y_tabella_footer=Y_timbro
	
		Y_min_tabella_footer=Y_tabella_footer-tabellaFooter.height-30	'30 altezza testo dati azienda
	
		if Y_min_tabella_footer<0 then
			Set Page = Doc.Pages.Add (792,612)
			redim Preserve lettera_pagina(doc.pages.count)
			lettera_pagina(doc.pages.count)="Totali"
		   pagina=pagina+1
		   Y_tabella_footer=page.height-75
		end if
	
		Y_tabella_footer=replace(Y_tabella_footer,",",".")
		Page.Canvas.DrawTable tabellaFooter, "x=20, y="&Y_tabella_footer
	end if
	'Cicli successivi imposto inizio pagina a ultima pagina ciclo precedente +1
	if start_Pagina>1 then
		start_Pagina=pagine+1
	end if

	
	pagine=doc.pages.count

	for pagina=start_Pagina to pagine
		Set Page = Doc.Pages.item(pagina)
		Page.Canvas.DrawTable tabella_titolo, "x="&tabella_titolo_X&", y="&replace(tabella_titolo_Y,",",".")
		if isobject(logo) then
			Page.Canvas.DrawImage logo, parametri_logo
		end if
		if pagina=1 then
			Page.Canvas.DrawTable tabella__accessori, "x="&margine_x&", y="&replace(tabella_accessori_Y,",",".")
		end if
	next
	
next
'Numero tutte le pagine
for pagina=1 to pagine
	Set Page = Doc.Pages.item(pagina)
	Page.Canvas.DrawText fondo_listino,  "x="&(page.width-350)/2&"; y=25; width=350; alignment=center; size=8;html=true", doc.fonts("Arial")
	
	if cicli>1 then
	
		testoPagina="Pagina "&pagina&" da "&lettera_pagina(pagina)&" di "&pagine
	else
		testoPagina="Pagina "&pagina&" di "&pagine
	end if
	width=page.width/2-margine_x
	'Fondo pagina
	'Page.Canvas.DrawText testoPagina, "X="&replace(width+margine_x,",",".")&"; Y=20; Alignment=1; width="&replace(width,",",".")&"; size=8;" , doc.fonts("Arial")
	Page.Canvas.DrawText testoPagina, "X="&replace(width+margine_x-10,",",".")&"; Y=560; Alignment=1; width="&replace(width,",",".")&"; size=8;" , doc.fonts("Arial")
	
	
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
	
	
function startPagina(ciclo)
	if ciclo=1 then
		startDip=0
	else
		startDip=dipperpagina*(ciclo-1)
	end if
end function
function endDip(ciclo)
	endDip=dipperpagina*ciclo-1
end function
function offsDip(ciclo,n)
	
	offsDip=(ciclo-1)*(dipperpagina)+n
end function







function grado_colorato(grado,colore)
	if colore<>"" then
		grado_colorato="<font color=""#"&colore&""">"&grado&"</font>"
		
	else
		grado_colorato="<font color=""black"">"&grado&"</font>"
		
	end if
end function
function verde(valore)
	if valore=1 then 
		verde=&H99FF99
	else
		verde=&Hffffff
	end if
	
end function
	
function avanzamento(quantita, in_ordine,pronto,consegnato,byref classe)
	classe=&HFFFFFF
	txt2=""
		if in_ordine>0 then
			if in_ordine>=quantita then
				classe=&Hf8c200
			else
				call concatena_stringa(txt2,",","<span class=""in_ordines"">"&in_ordine&"</span>")
			end if
		end if
		if pronto>0 then
			if pronto>=quantita then
				classe=&H0eff00
			else
				call concatena_stringa(txt2,",","<span class=""prontos"">"&pronto&"</span>")
			end if
		end if
		if consegnato>0 then
			if consegnato>=quantita then
				classe=&H297bff
			else
				call concatena_stringa(txt2,",","<span class=""consegnatos"">"&consegnato&"</span>")
			end if
		end if
		avanzamento=txt2
end function
function classe_articolo(quantita,tot_in_ordine,tot_pronto,tot_consegnato)
	quantita=cdbl(quantita)
	classe_articolo=&HFFFFFF
	if tot_consegnato=quantita then
		classe_articolo=&H297bff
	elseif tot_pronto=quantita then
		classe_articolo=&H0eff00
	elseif tot_in_ordine=quantita then
		classe_articolo=&Hf8c200
	end if
end function

function dipendente_colorato(sesso)
	if sesso="F" then 
		dipendente_colorato=&Hf74ca5
	else
		dipendente_colorato=&H000000
	end if
end function




%>


