<%
'Verifica chiusure 30_11_2015
'Migliorabile con getrows su rs
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/ClasseOrdine.asp" -->
<%
Set Pdf = Server.CreateObject("Persits.PDFManager")
Set Doc = Pdf.CreateDocument
Set TextParam = PDF.CreateParam
Set ImageParam = PDF.CreateParam

idord=request("idord")
m_tabella=request("tabella")
per_cliente=request("per_cliente")
if per_cliente="" then
	per_cliente=false
else
	per_cliente=true
end if
	
set ordine= (new ClasseOrdine)(array("apri",m_tabella,idord))

nord=ordine.campo("nord")
if m_tabella="ordini" then
	Doc.Title=firstup(ordine.campo("tipo_documento"))&" "&nord
else
Doc.Title = firstup(tabella_singolare(m_tabella)&" "&nord)
end if
Doc.Creator = nomesito
Bordo_sx=25
Set Page = Doc.Pages.Add
'page.height=825
Set Param = Pdf.CreateParam
Set Param2 = Pdf.CreateParam
Set fs = Server.CreateObject("Scripting.filesystemObject")
if fs.FileExists(Server.MapPath( "/public/files/logo_documenti.jpg")) then
	esiste_logo=true
	
	'Parte nuova
	dimensione_logo=170
	margine=5
	Set logo = Doc.OpenImage( Server.MapPath( "/public/files/logo_documenti.jpg") )
	'calcolo misure dell'immagine considerando la risoluzione DPI
	larghezza=logo.width*72/logo.resolutionx
	altezza=logo.height*72/logo.resolutiony
	
	
	scala=(dimensione_logo)/larghezza
	
	'calcolo nuova altezza considerando la scala per centrare poi l'immagine
	nuova_altezza=altezza*scala
	nuova_larghezza=larghezza*scala
	
	Set Param_logo = PDF.CreateParam
	
	'Trovo la x per centrare l'immagine
	'Param_logo("x") =  (row.cells(1).width - nuova_larghezza ) / 2 'Centra l'immagine
	'Param_logo("x") =  (row.cells(1).width - nuova_larghezza ) + margine'Allinea a SX
	Param_logo("x") =(Page.Width/ 2) - margine - dimensione_logo
	
	'Sotto laschio metà del margine
	ImageParam("y") = margine/2
	Param_logo("y") = Page.Height - 110
	
	Param_logo("ScaleX") = scala
	Param_logo("ScaleY") = scala
end if
Set fs = Nothing
if m_tabella="ordini" then
	spettanze=ordine.campo("spettanze")
	if per_cliente then spettanze=0
end if

Bordo_sx=25
page.height=825
Set Param = Pdf.CreateParam
Set Param2 = Pdf.CreateParam

'TABELLA 1
Set T_testata = Doc.CreateTable("width="&int((page.width/2)-3)&"; height=600; Rows=1; Cols=3; Border=1; CellSpacing=-1; cellpadding=2 ")
T_testata.Font = Doc.Fonts("Arial")
With T_testata.Rows(1)
	'.BGColor = &H000000
	.Cells(1).Width = (T_testata.width/3-4)+15
	.Cells(1).AddText Doc.Title, ""
	.Cells(2).Width = (T_testata.width/3-4)-13
	.Cells(2).AddText "del "& formatDateTime(ordine.campo("data"), vbShortDate) , ""
	.Cells(3).Width = (T_testata.width/3-7)-3.5
	.Cells(3).AddText "Pagina 1 di 1", ""
	.Cells(3).border=1
	.height=20
End With

'TABELLA 2
Set Table2 = Doc.CreateTable("width="&(page.width/2)-2&"; height=600; Rows=4; Cols=4; Border=1; CellSpacing=-1; cellpadding=2 ")
Table2.Font = Doc.Fonts("Arial")
'T_testata.rows.add(15)
param.add "html=true"
With Table2.Rows(1)
	'.BGColor = &HCCCCCC
	.cells(1).colspan=3
	.Cells(1).Width = (page.width/2-34)/4
	.Cells(1).AddText "PAGAMENTO<br><b>"&ucase( metodo_pagamento(ordine.campo("tipopagamento"),"descrizione"))&"</b>", "size=8; html=true; Expand=true;"
	.Cells(4).Width = (page.width/2-34)/4
	.Cells(4).AddText "COD.CLIENTE<br><b>"& ordine.campo("iduser")&"</b>", "size=8; html=true; Expand=true;"
	.height=28
End With
With Table2.Rows(2)
	.cells(1).colspan=4
	.Cells(1).AddText "RIFERIMENTO D.D.T.", "size=8; html=true; Expand=true;"
	'.Cells(2).AddText "NUMERO COLLI<br><b>"& ordine.campo("colli")&"</b>", "size=9; html=true; Expand=true;"
	.height=28
End With

banca_appoggio_arr=banca_appoggio(ordine.campo("idbanca"))

With Table2.Rows(3)
	.cells(1).colspan=4
	if metodo_pagamento(ordine.campo("tipopagamento"),"riba")=1 then
		testo=ordine.campo("banca_appoggio")
	else
		testo=banca_appoggio_arr(0)
	end if
	.Cells(1).AddText "BANCA D'APPOGGIO<br><b>"&ucase(testo)&"</b>", "size=8; html=true; Expand=true;"
	'.Cells(2).AddText "NUMERO COLLI<br><b>"& ordine.campo("colli")&"</b>", "size=9; html=true; Expand=true;"
	.height=28
End With
With Table2.Rows(4)
	.cells(1).colspan=4
	testo=""
	if isnumeric(ordine.campo("tipopagamento")) then
		if ordine.campo("tipopagamento") >= 10 and ordine.campo("tipopagamento")<=14 then
			testo=ordine.campo("iban")
		else
			testo=banca_appoggio_arr(1)
		end if
	end if

	.Cells(1).AddText "IBAN<br><b>"&testo&"</b>", "size=8; html=true; Expand=true;"
	.height=28
End With


'TABELLA 3
Set T_INTESTAZIONE = Doc.CreateTable("width="&page.width-40&"; height=600; Rows=1; Cols=2; Border=1; CellSpacing=-1; cellpadding=2 ")
T_INTESTAZIONE.Font = Doc.Fonts("Arial")
'T_testata.rows.add(15)
With T_INTESTAZIONE.Rows(1)
	'.BGColor = &HCCCCCC
	.Cells(1).Width = (page.width-40)/2
	testo=""
	if ordine.campo("d_indirizzo")<>"" then
		testo=ucase(ordine.campo("d_azienda")&"<br>"&ordine.campo("d_indirizzo")&"<br>"&ordine.campo("d_cap")&" "&ordine.campo("d_citta")&" "&ordine.campo("d_provincia"))
	else
		testo="IDEM"
		'testo=ucase(ordine.campo("azienda")&"<br>"&ordine.campo("indirizzo")&"<br>"&ordine.campo("cap")&" "&ordine.campo("citta")&" "&ordine.campo("provincia"))
	end if
	.Cells(1).AddText "LUOGO DI DESTINAZIONE<br><b>"&testo&"</b>", "size=9; html=true; Expand=true;"
	.Cells(2).Width = (page.width-40)/2
	testo="P.Iva "&ordine.campo("piva")&" "&"Cod. fisc. "&ordine.campo("cf")
	.Cells(2).AddText "SPETT.LE "&dati_fatturazione_pdf(ordine.campo("idintestazione")), "size=9; html=true; Expand=true;"

	.height=62 '60
End With

If m_tabella="ordini" then
Set Row = T_INTESTAZIONE.Rows.Add(30) ' row height
with T_INTESTAZIONE.rows(2)
	.Cells(1).AddText "NOTE SPEDIZIONE<br>"&ordine.campo("note_spedizione"), "size=7; html=true;"
	.Cells(2).AddText "NOTE SU ORDINE<br>"&ordine.campo("note"), "size=7; html=true;"
end with
Else
Set Row = T_INTESTAZIONE.Rows.Add(30) ' row height
with T_INTESTAZIONE.rows(2)
	.Cells(1).AddText "VALIDITA' OFFERTA 30 GG.<br>"&ordine.campo("note_spedizione"), "size=7; html=true;"
	.Cells(2).AddText "NOTE SU PREVENTIVO<br>"&ordine.campo("note"), "size=7; html=true;"
end with

End if
if m_tabella="ordini" then
	colonne=-1
	colonne_txt=array("")
	if ordine.campo("cig")<>"" then
		colonne=colonne+1
		Redim PRESERVE colonne_txt(colonne)
		colonne_txt(colonne)="Codice CIG: "&ordine.campo("cig")
	end if
	
	if ordine.campo("impegno_spesa")<>"" then
		colonne=colonne+1
		Redim PRESERVE colonne_txt(colonne)
		colonne_txt(colonne)="Impegno di spesa: "&ordine.campo("impegno_spesa")
	end if
	if ordine.campo("mepa_tipo")>0 then
		colonne=colonne+1
		Redim PRESERVE colonne_txt(colonne)
		colonne_txt(colonne)=mepa_tipo(ordine.campo("mepa_tipo"))&":"&ordine.campo("mepa_testo")
	end if
	if ordine.campo("mepa_data")<>"" then
		colonne=colonne+1
		Redim PRESERVE colonne_txt(colonne)
		colonne_txt(colonne)="Data ordine: "&ordine.campo("mepa_data")
	end if
	if colonne>=0 then
		Set T_INTESTAZIONEbis = Doc.CreateTable("width="&page.width-40&"; height=20; Rows=1; Cols="&colonne+1&"; Border=1; CellSpacing=-1; cellpadding=2 ")
		T_INTESTAZIONEbis.Font = Doc.Fonts("Arial")
		col_width=(page.width-40)/(colonne+1)
		for n = 1 to colonne+1
			if n=colonne+1 then
				'diff=page.width-40-(col_width*n)
				'col_width=col_width+diff
				'diff="page.width-40:"&page.width-40&" col_width*n:"&col_width*n
				col_width=col_width+1*(n-1)-1
			end if
			 
				'on error resume next
				T_INTESTAZIONEbis.Rows(1).Cells(n).Width = col_width
				
				if err.number<>0 then
					response.write "colonne:"&n&" ubound:"&ubound(colonne_txt)
					response.end
				end if
				T_INTESTAZIONEbis.Rows(1).Cells(n).AddText colonne_txt(n-1)&" "&diff, "size=8;  alignment=0; Expand=true;"
		next
	end if
	
end if
if m_tabella="ordini" and mod_larochelle and not per_cliente then

		Set T_RESI = Doc.CreateTable("width="&page.width-41&"; height=20; Rows=1; Cols=1; Border=1; CellSpacing=-1; cellpadding=2 ")
		T_RESI.Font = Doc.Fonts("Arial")
		riga=1
		if ordine.campo("note_gestore")<>"" then
			T_RESI.Rows(riga).Cells(1).BGColor = &Hffff3a
			T_RESI.Rows(riga).Cells(1).AddText "<b>Note interne:</b><br>"&ordine.campo("note_gestore"), "alignment=0; size=8; Expand=true; html=true; color=red;"
			riga=riga+1
		end if
		if riga>1 then
			Set Row = T_RESI.Rows.Add(10) ' row height
		end if


		T_RESI.Rows(riga).Cells(1).AddText txt_restituzione, "size=8;  html=true; alignment=0; Expand=true;"
		
	
end if



''TABELLA5
'Set T_intestazione_articoli = Doc.CreateTable("width="&page.width-40&"; height=20; Rows=1; Cols=8; Border=1; CellSpacing=-1; cellpadding=2 ")
'T_intestazione_articoli.Font = Doc.Fonts("Arial")

Set T_elenco_articoli = Doc.CreateTable("width="&page.width-40&"; height=10; Rows=1; Cols=8; Border=1; CellSpacing=-1; cellpadding=2 ")
T_elenco_articoli.Font = Doc.Fonts("Arial")
With T_elenco_articoli.Rows(1)
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


T_elenco_articoli.border=0
sconto_totale=0
totord=0
n=2


'ordini_dett_esteso="select ordini_dett.*, magazzino.quantita_magazzino, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, varianti_a.variante_a, varianti_b.variante_b FRom ((magazzino RIGHT JOIN ordini_dett oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JOIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JOIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara "

ordini_dett_esteso="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b, ordini_dett_note.nota FRom (((ordini_dett left join ordini_dett_note on ordini_dett.iddett = ordini_dett_note.iddett) ) LEFT JoIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara "


'ordini_dett_esteso="select ordini_dett.*, magazzino.quantita_magazzino,magazzino.idmag, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, magazzino.db_ven, varianti_a.variante_a, varianti_b.variante_b, ordini_dett_note.nota FRom (((magazzino RIGHT JoIN (ordini_dett left join ordini_dett_note on ordini_dett.iddett = ordini_dett_note.iddett) oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JoIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara ) left join prodotti on  magazzino.idpro = prodotti.idpro where idord="&m_idord &" order by ordine,iddett"


sql_dettaglio=replace(ordini_dett_esteso,"ordini",m_tabella)&"  where idord="&idord&" order by ordine,iddett;"
set rs1=conn.execute(sql_dettaglio)
Do While Not rs1.EOF



	testo=""
	prezzo=cdbl(rs1("prezzo"))
	sconto_prodotto=cdbl(rs1("sconto_prodotto"))
	quantita=cdbl(rs1("quantita"))
	'if n>2 then
		Set Row = T_elenco_articoli.Rows.Add(10) ' row height
	'end if
	
	
	
	if mod_larochelle then
		call imposta_bordi( row, "bottom=False")
	else
		call imposta_bordi( row, "bottom=False, Top=False")
		
		
	end if
	
	
	
	
	'Row.Cells(1).border =0
	if rs1("codice_ordine")<>"" then
		Row.Cells(1).AddText codice_articolo_e_variante( rs1("codice_ordine"),rs1("codicevara")), "alignment=1; size="&size_chr_pdf&"; Expand=true; Color=black; html=true;"
	end if
	if rs1("varianti_ordine")<>"" then testo=testo&vbcrlf&rs1("varianti_ordine")
	if m_tabella="ordini" then
		if cint(rs1("numeri_di_serie"))>0 then
			
			testo=testo&vbcrlf&elenca_seriali(rs1("iddett"),quantita)
		end if
	end if
	
	
	
	
	Row.Cells(2).AddText rs1("articolo_ordine")&testo, "alignment=0; size="&size_chr_pdf&"; Expand=true; Color=black;"
	'Row.Cells(2).border =0
	if rs1("um")<>"" then
		Row.Cells(3).AddText rs1("um"), "alignment=2; size="&size_chr_pdf&"; Expand=true;"
	end if
	'Row.Cells(3).border =0
	if quantita>0 then Row.Cells(4).AddText quantita, "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	'Row.Cells(4).border =0
	nascondi=true
	if prezzo<>0 then
		Row.Cells(5).AddText formatcurrency(rs1("prezzo"),2), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
		nascondi=false
	end if

	'Row.Cells(5).border =0
	testo=""
	if sconto_prodotto<>0 then
		Row.Cells(6).AddText formatnumber(rs1("sconto_prodotto"),2)&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	end if
	'Row.Cells(6).border =0
	'Row.Cells(7).border =0
	if isnull(prezzo) then prezzo=0
	subtot=totale_riga(prezzo,sconto_prodotto,quantita)
	'subtot=(prezzo-(sconto_prodotto/100)*prezzo)*rs1("quantita")
	'subtot=round(subtot,2)
	if nascondi=false then
		Row.Cells(7).AddText formatcurrency(subtot,2), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
		Row.Cells(8).AddText iva(ordine.campo("data"))&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"

	end if
	'Row.Cells(8).border =0
	totord=totord+subtot
		'Set Row = T_elenco_articoli.Rows.Add(10) ' row height
		'Row.Cells(2).AddText "TOTORDB"&totord, "alignment=0; size=7; Expand=true;"
		
	if spettanze=1 then
		'Cerco nelle spettanze
		txt=""
		conteggio_seriali=0
		sql="select "&m_tabella&"_dett_spettanze.quantita, "&m_tabella&"_dett_spettanze.taglia_misura, dipendenti.matricola, dipendenti.cognome, dipendenti.nome, dipendenti.sesso from ("&m_tabella&"_dett_spettanze inner join dipendenti on "&m_tabella&"_dett_spettanze.iddip=dipendenti.iddip) where iddett="&RS1("iddett")&" order by sesso desc,cognome,nome"
        sql=replace(sql,"ordini",m_tabella)

		set rs_spettanze=conn.execute (sql)
		quantita_tot=quantita
		do while not rs_spettanze.eof 
			Set Row = T_elenco_articoli.Rows.Add(10) ' row height
			call imposta_bordi( row, "bottom=False, Top=False")
			if not per_cliente then
				taglia_misura=rs_spettanze("taglia_misura")
				if taglia_misura<>"" then taglia_misura="("&taglia_misura&")"
			end if
			Row.Cells(2).AddText rs_spettanze("cognome")&" "&rs_spettanze("nome")&" "&taglia_misura , "alignment=0; size=7; html=true; Expand=true;"
			Row.Cells(4).AddText rs_spettanze("quantita"), "alignment=1; size=7; html=true; Expand=true;"
			quantita_tot=quantita_tot-rs_spettanze("quantita")
			rs_spettanze.movenext
		loop
		if quantita_tot>0 then
			Set Row = T_elenco_articoli.Rows.Add(10) ' row height
			call imposta_bordi( row, "bottom=False, Top=False")
			Row.Cells(2).AddText "Scorta tecnica" , "alignment=0; size=7; html=true; Expand=true;"
			Row.Cells(4).AddText quantita_tot, "alignment=1; size=7; html=true; Expand=true;"
		end if
		'Row.Cells(1).SetBorderParams "Left=False, Right=False, Top=False"
		'Row.Cells(2).SetBorderParams "Left=False, Right=False, Top=False"
		'Row.Cells(3).SetBorderParams "Left=False, Right=False, Top=False"
		'Row.Cells(4).SetBorderParams "Left=False, Right=False, Top=False"
		'Row.Cells(5).SetBorderParams "Left=False, Right=False, Top=False"
		'Row.Cells(6).SetBorderParams "Left=False, Right=False, Top=False"
		'Row.Cells(7).SetBorderParams "Left=False, Right=False, Top=False"
		'Row.Cells(8).SetBorderParams "Left=False, Right=False, Top=False"
	end if
	
	if rs1("nota")<>"" and not per_cliente then
		Set Row = T_elenco_articoli.Rows.Add(10) ' row height
		if mod_larochelle then
			call imposta_bordi( row, "bottom=true, Top=False")
		end if
		Row.Cells(2).BGColor = &Hffff3a
		Row.Cells(2).AddText "Nota: "&rs1("nota") , "alignment=0; size=7; html=true; Expand=true; Color=red;"
	end if

	rs1.movenext
	n=n+1
Loop
Set rs1 = Nothing





'Conversione decimal in dbl
totale_ordine=cdbl(ordine.campo("totale"))
trasporto=cdbl(ordine.campo("trasporto"))
sconto_ordine=cdbl(ordine.campo("sconto_ordine"))
imposta_ordine=cdbl(ordine.campo("imposta_ordine"))
if tabella="ordini" then
	spese_bancarie_ordine=cdbl(ordine.campo("spese_bancarie_ordine"))
	arrotondamento=cdbl(ordine.campo("arrotondamento"))
	costo_totordine=cdbl(ordine.campo("costo_totordine"))
end if

if trasporto>0 then
	Set Row = T_elenco_articoli.Rows.Add(10) ' row height
	Row.Cells(2).AddText "Trasporto", "alignment=0; size=7; Expand=true; Color=black;"
	Row.Cells(5).AddText formatcurrency(ordine.campo("trasporto"),2), "alignment=1; size=7; Expand=true;"
	Row.Cells(7).AddText formatcurrency(ordine.campo("trasporto"),2), "alignment=1; size=7; Expand=true;"
	Row.Cells(8).AddText iva(ordine.campo("data"))&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	
	
	call imposta_bordi( row, "bottom=False")
	totord=totord+trasporto
end if
iva_ordine= iva(ordine.campo("data"))





if sconto_ordine>0 then
	Set Row = T_elenco_articoli.Rows.Add(10) ' row height
	call imposta_bordi( row, "bottom=False")
	Row.Cells(2).AddText "Sconto ordine", "alignment=0; size=7; Expand=true; Color=black;"
	Row.Cells(5).AddText formatcurrency(ordine.campo("sconto_ordine"),2), "alignment=1; size=7; Expand=true;"
	Row.Cells(7).AddText formatcurrency(ordine.campo("sconto_ordine"),2), "alignment=1; size=7; Expand=true;"
	
	Row.Cells(8).AddText iva_ordine&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	sconto_totale=sconto_totale+sconto_ordine
end if
'Perchè ho messo questa riga?
if mod_larochelle then
	Set Row = T_elenco_articoli.Rows.Add(1) ' row height
	call imposta_bordi( row, "bottom=False, Top=False")
end if





if ordine.campo("nascondi_totali")=1 then
	
	Set T_totali = Doc.CreateTable("width="&page.width-41&"; height=100; Rows=1; Cols=1; Border=1; CellSpacing=-1; cellpadding=1 ")
	T_totali.Font = Doc.Fonts("Arial")
	With T_totali.Rows(1)
		'.Cells(1).Width = (page.width-40)/5+0.6
		.Cells(1).AddText "I prezzi sono iva esclusa", "alignment=2; size=9; Expand=true;"
		.height=14
	End With

else
	


	Set T_totali = Doc.CreateTable("width="&page.width-40&"; height=100; Rows=4; Cols=5; Border=1; CellSpacing=-1; cellpadding=1 ")
	T_totali.Font = Doc.Fonts("Arial")
	With T_totali.Rows(1)
		.Cells(1).Width = (page.width-40)/5+0.6
		'.Cells(1).AddText "SPESE BANCARIE", "alignment=2; size=7; Expand=true;"
		.cells(1).border=0
		.Cells(2).AddText "IMPONIBILE", "alignment=2; size=7; Expand=true; Color=black;"
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
			.Cells(1).AddText "SCONTO SU "&ucase(tabella_singolare(m_tabella)), "alignment=2; size=7; Expand=true;"
		end if
		.cells(1).SetBorderParams "bottom=False"
		.cells(5).SetBorderParams "bottom=False"
		.Cells(5).AddText "TOTALE "&ucase(tabella_singolare(m_tabella)), "alignment=2; size=7; Expand=true;"
		.height=14
	End With
	With T_totali.Rows(4)
		'.Cells(1).AddText "", "alignment=2; size=7; Expand=true;"
		.cells(1).SetBorderParams "Top=False" 
		.cells(5).SetBorderParams "Top=False" 
		.height=19
	End With
	
	call T_totali_compila( T_totali, totord, sconto_ordine, ordine.campo("trattamento_iva_ordine"), iva_ordine, 0 ,0)
	
	



end if


param.clear
Param("x") = 20
page_height=page.height
pagina=1
FirstRow = 1
session("T_elenco_articoli")=""
Do While True

	'Esegue un loop ad ogni pagina
	
	Y=page.height
	
	'Intestazione per ogni pagina
	'if esiste_logo then Page.Canvas.DrawImage logo, Param_logo
	'Page.Canvas.DrawText intestazione_ddt_fat(ordine.campo("data")), "X=20; Y="&Y-20&"; html=true" , doc.fonts("Arial")
	call intestazione_pdf_ddt_fat(doc, page,ordine.campo("data"))
	'T_testata renderizzata dopo
	Y=Y-147
	Page.Canvas.DrawTable T_INTESTAZIONE, "x=20, y="&Y
	session("Y_T_INTESTAZIONE")=Y
	'Page.Canvas.DrawTable T_intestazione_articoli, "x=20, y="&page.height-233-3
	'Page.Canvas.DrawTable T_totali, "x=20, y="&page.height-715-3
	Y=Y-int(T_INTESTAZIONE.height)+1
	
	'if not ( rs_ordine is Nothing)
	if isobject(T_INTESTAZIONEbis) and pagina=1 then
		Page.Canvas.DrawTable T_INTESTAZIONEbis, "x=20, y="&Y
		session("Y_T_INTESTAZIONEbis")=Y
		Y=Y-roundup(T_INTESTAZIONEbis.height,0)+1
	end if
	
	
	
	'Page.Canvas.DrawTable Table8, "x=20, y="&page.height-217
	session("Y_T_intestazione_articoli")=Y
	if isobject(T_RESI) and pagina=1 then
		Page.Canvas.DrawTable T_RESI, "x=20, y="&Y
		session("Y_T_INTESTAZIONEbis")=Y
		Y=Y-roundup(T_RESI.height,0)+1
	end if
	'Page.Canvas.DrawTable Table8, "x=20, y="&page.height-217
	session("T_RESI")=Y
	
	
	
	
	
	
	'Page.Canvas.DrawTable T_intestazione_articoli, "x=20, y="&Y
	'Y=Y-int(T_intestazione_articoli.height)+1
	Param("y") = Y
	session("Y_T_elenco_articoli")=Y
	Y_T_elenco_articoli=Y
	'50 =Table8 Y
	Param("MaxHeight") = Y-50
	
	
	LastRow = Page.Canvas.DrawTable (T_elenco_articoli, param)
	
	T_elenco_articoli.Rows(LastRow-1).Cells(1).SetBorderParams "bottom=True, "

	
	h=0
	'Calcolo altezza tabella parziale
	for n=FirstRow to LastRow
		h=h+T_elenco_articoli.rows(n).height-1
	next
	h=int(h)
	
	
	session("T_elenco_articoli")=session("T_elenco_articoli")&","&h
	'T_intestazione_articoli.rows(2).height=h
	

	if LastRow >= T_elenco_articoli.Rows.Count Then
		'call add2log("LastRow:"&LastRow&" T_elenco_articoli.Rows.Count:"&T_elenco_articoli.Rows.Count&"doc.pages.count:"&doc.pages.count,0)
		Exit Do ' entire table displayed
	else
		'Ci sono più pagine
		
		'TABELLA 8
		Set Table8 = Doc.CreateTable("width="&page.width-40&"; height=30; Rows=1; Cols=1; Border=1; CellSpacing=-1; cellpadding=2 ")
		Table8.Font = Doc.Fonts("Arial")
		With Table8.Rows(1)
			.Cells(1).Width = (page.width-40)
		End With

		
		
		if pagina>1 then
			h=h+T_elenco_articoli.Rows(1).height
		end if
		y=Y-h-1

		Page.Canvas.DrawLine 21, y, page.width-22, y
	end if

   ' Display remaining part of table on the next page
   Set Page = Page.NextPage
   Param.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
   Param("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
   FirstRow = LastRow + 1
   pagina=pagina+1
Loop
y_T_totali=page_height-715-3
y_fine_T_elenco_articoli=Y_T_elenco_articoli-h

'call add2log("y_T_totali:"&y_T_totali&" y_fine_T_elenco_articoli:"&y_fine_T_elenco_articoli&" T_elenco_articoli.height:"&T_elenco_articoli.height&" h:"&h,0)
if y_fine_T_elenco_articoli<y_T_totali then
   Set Page = Page.NextPage
	'Intestazione per ogni pagina
	'if esiste_logo then Page.Canvas.DrawImage logo, Param_logo
	'Page.Canvas.DrawText intestazione_ddt_fat(ordine.campo("data")), "X=20; Y="&Page.Height-20&"; html=true" , doc.fonts("Arial")
	call intestazione_pdf_ddt_fat(doc, page,ordine.campo("data"))
	
	
	'T_testata renderizzata dopo
	Y=147
	Page.Canvas.DrawTable T_INTESTAZIONE, "x=20, y="&page.height-Y
	'Page.Canvas.DrawTable T_intestazione_articoli, "x=20, y="&page.height-233-3
	'Page.Canvas.DrawTable T_totali, "x=20, y="&page.height-715-3


	
end if




pagine=doc.pages.count
'call add2log("pagine:"&pagine,0)
for pagina=1 to pagine
	Set Page = Doc.Pages.item(pagina)
	
	
	T_testata.rows(1).Cells(3).cleartext
	T_testata.rows(1).Cells(3).AddText "Pagina "&pagina&" di "&pagine, ""
	if pagina=pagine then
		
		
		if tabella="ordini" then
			iva_ordine=ordine.campo("iva_ordine")
		else
			iva_ordine=iva(date())
		end if

		'Completamento griglia
		Set T_completamento_griglia= Doc.CreateTable("width="&page.width-40&"; height=20; Rows=1; Cols=8; Border=1; CellSpacing=-1; cellpadding=2 ")

		Y_T_totali=page.height-715-3
		With T_completamento_griglia.Rows(1)
			.Cells(1).Width = 60
			.Cells(2).Width = page.width-60-30-40-50-50-55-30-34
			.Cells(3).Width = 30
			.Cells(4).Width = 40
			.Cells(5).Width = 50
			.Cells(6).Width = 50
			.Cells(7).Width = 55
			.Cells(8).Width = 30
			.height=Y-Y_T_totali
		End With
		Page.Canvas.DrawTable T_completamento_griglia, "x=20, y="&Y
		
		
		




		Page.Canvas.DrawTable T_totali, "x=20, y="&Y_T_totali


	end if
	if pagine>1 and pagina<pagine and isobject(Table8)  then
		Table8.Rows(1).Cells(1).cleartext
		Table8.Rows(1).Cells(1).AddText "Segue a pagina "&pagina+1, "alignment=2; size=8; Expand=true; html=true;"
		Page.Canvas.DrawTable Table8, "x=20, y=50"
	end if
	
	Page.Canvas.DrawTable T_testata, "x="& (page.width/2)-1 &", y="&page.height-20
	Page.Canvas.DrawTable Table2, "x="& (page.width/2)-1 &", y="&page.height-39




next
Set rs = Nothing
call connclose()
Doc.Title=Doc.Title&" "&year(ordine.campo("data"))
Doc.SaveHttp "attachment;filename="&replace(Doc.Title," ","_")&".pdf" 

Set Doc = Nothing
Set TextParam = Nothing
Set ImageParam = Nothing
Set Param_logo = Nothing
Set Pdf = Nothing

call CheckConnChiusa()

function aggiungi_zeri(n)
'aggiungi_zeri=string(5-len(n),"0")&n
aggiungi_zeri=n
end function

sub imposta_bordi(ByRef row, bordo)
			Row.Cells(1).SetBorderParams bordo
			Row.Cells(2).SetBorderParams bordo
			Row.Cells(3).SetBorderParams bordo
			Row.Cells(4).SetBorderParams bordo
			Row.Cells(5).SetBorderParams bordo
			Row.Cells(6).SetBorderParams bordo
			Row.Cells(7).SetBorderParams bordo
			Row.Cells(8).SetBorderParams bordo
end sub

%>


