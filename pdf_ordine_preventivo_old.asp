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
set ordine= (new ClasseOrdine)(array("apri",m_tabella,idord))

nord=ordine.campo("nord")
Doc.Title = firstup(tabella_singolare(m_tabella)&" "&nord)
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
end if

Bordo_sx=25
page.height=825
Set Param = Pdf.CreateParam
Set Param2 = Pdf.CreateParam

'TABELLA 1
Set Table1 = Doc.CreateTable("width="&int((page.width/2)-3)&"; height=600; Rows=1; Cols=3; Border=1; CellSpacing=-1; cellpadding=2 ")
Table1.Font = Doc.Fonts("Arial")
With Table1.Rows(1)
	'.BGColor = &H000000
	.Cells(1).Width = (table1.width/3-4)+15
	.Cells(1).AddText Doc.Title, ""
	.Cells(2).Width = (table1.width/3-4)-13
	.Cells(2).AddText "del "& formatDateTime(ordine.campo("data"), vbShortDate) , ""
	.Cells(3).Width = (table1.width/3-7)-3.5
	.Cells(3).AddText "Pagina 1 di 1", ""
	.Cells(3).border=1
	.height=20
End With

'TABELLA 2
Set Table2 = Doc.CreateTable("width="&(page.width/2)-2&"; height=600; Rows=4; Cols=4; Border=1; CellSpacing=-1; cellpadding=2 ")
Table2.Font = Doc.Fonts("Arial")
'table1.rows.add(15)
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
if ordine.campo("idbanca")="" or isnull(ordine.campo("idbanca")) or ordine.campo("idbanca")=0  then
	idbanca=1
else
	idbanca=ordine.campo("idbanca")
end if
set rsb=conn.execute("select * from banche where idbanca="&idbanca)
With Table2.Rows(3)
	.cells(1).colspan=4
	if metodo_pagamento(ordine.campo("tipopagamento"),"riba")=1 then
		testo=ordine.campo("banca_appoggio")
	else
		testo=rsb("denominazione")
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
			testo=rsb("iban")
		end if
	end if

	.Cells(1).AddText "IBAN<br><b>"&testo&"</b>", "size=8; html=true; Expand=true;"
	.height=28
End With
rsb.close
set rsb=nothing


'TABELLA 3
Set Table3 = Doc.CreateTable("width="&page.width-40&"; height=600; Rows=1; Cols=2; Border=1; CellSpacing=-1; cellpadding=2 ")
Table3.Font = Doc.Fonts("Arial")
'table1.rows.add(15)
With Table3.Rows(1)
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
	.Cells(2).AddText "SPETT.LE<br><b>"&ucase(denominazione(ordine.campo("nome"),ordine.campo("cognome"),ordine.campo("azienda"))&"<br>"&ordine.campo("indirizzo")&"<br>"&ordine.campo("cap")&" "&ordine.campo("citta")&" "&ordine.campo("provincia")&"</b>"&"<br>"&testo), "size=9; html=true; Expand=true;"

	.height=62 '60
End With
Set Row = Table3.Rows.Add(30) ' row height
with table3.rows(2)
	.Cells(1).AddText "NOTE SPEDIZIONE<br>"&ordine.campo("note_spedizione"), "size=7; html=true;"
	.Cells(2).AddText "NOTE SU ORDINE<br>"&ordine.campo("note"), "size=7; html=true;"
end with
if m_tabella="ordini" then
	if ordine.campo("cig")<>"" or ordine.campo("mepa")<>"" or ordine.campo("impegno_spesa")<>"" then
		Set Table3bis = Doc.CreateTable("width="&page.width-40&"; height=10; Rows=1; Cols=3; Border=1; CellSpacing=-1; cellpadding=2 ")
		Table3bis.Font = Doc.Fonts("Arial")
		With Table3bis.Rows(1)
			.Cells(1).Width = (page.width-40)/3
			.Cells(1).AddText "Codice CIG: "&ordine.campo("cig"), "size=8; "
			.Cells(2).Width = (page.width-40)/3
			.Cells(2).AddText "Ordine MEPA: "&ordine.campo("cig"), "size=8; "
			.Cells(3).Width = (page.width-40)/3
			.Cells(3).AddText "Impegno di spesa: "&ordine.campo("cig"), "size=8; "
			.height=20
			
		end with
		
		
	end if
	
end if



'TABELLA5
Set T_intestazione_articoli = Doc.CreateTable("width="&page.width-40&"; height=600; Rows=1; Cols=8; Border=1; CellSpacing=-1; cellpadding=2 ")
T_intestazione_articoli.Font = Doc.Fonts("Arial")
With T_intestazione_articoli.Rows(1)
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
T_intestazione_articoli.Rows.Add(450)

Set T_elenco_articoli = Doc.CreateTable("width="&page.width-40&"; height=600; Rows=1; Cols=8; Border=1; CellSpacing=-1; cellpadding=2 ")
T_elenco_articoli.Font = Doc.Fonts("Arial")
With T_elenco_articoli.Rows(1)
	.Cells(1).Width = T_intestazione_articoli.rows(1).cells(1).width
	.Cells(1).border =0
	.Cells(2).Width = T_intestazione_articoli.rows(1).cells(2).width
	.Cells(2).border =0
	.Cells(3).Width = T_intestazione_articoli.rows(1).cells(3).width
	.Cells(3).border =0
	.Cells(4).Width = T_intestazione_articoli.rows(1).cells(4).width
	.Cells(4).border =0
	.Cells(5).Width = T_intestazione_articoli.rows(1).cells(5).width
	.Cells(5).border =0
	.Cells(6).Width = T_intestazione_articoli.rows(1).cells(6).width
	.Cells(6).border =0
	.Cells(7).Width = T_intestazione_articoli.rows(1).cells(7).width
	.Cells(7).border =0
	.Cells(8).Width = T_intestazione_articoli.rows(1).cells(8).width
	.Cells(8).border =0
	.height=20
End With
T_elenco_articoli.border=0
sconto_totale=0
totord=0
n=2
		'Set Row = T_elenco_articoli.Rows.Add(10) ' row height
		'Row.Cells(2).AddText "TOTORD"&totord, "alignment=0; size=7; Expand=true;"
		'Set Row = T_elenco_articoli.Rows.Add(10) ' row height
		'Row.Cells(2).AddText "TOTORDA"&totord, "alignment=0; size=7; Expand=true;"

sql_dettaglio=replace(ordini_dett_esteso,"ordini",m_tabella)&"  where idord="&idord&" order by ordine,iddett;"
set rs1=conn.execute(sql_dettaglio)
Do While Not rs1.EOF
	testo=""
	prezzo=cdbl(rs1("prezzo"))
	sconto_prodotto=cdbl(rs1("sconto_prodotto"))
	Set Row = T_elenco_articoli.Rows.Add(10) ' row height
	Row.Cells(1).border =0
	if rs1("codice_ordine")<>"" then
		Row.Cells(1).AddText rs1("codice_ordine"), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	end if
	if rs1("varianti_ordine")<>"" then testo=testo&vbcrlf&rs1("varianti_ordine")
	if m_tabella="ordini" then
		if cint(rs1("numeri_di_serie"))>0 then
			
			testo=testo&vbcrlf&elenca_seriali(rs1("iddett"),rs1("quantita"))
		end if
	end if
	
	
	
	
	Row.Cells(2).AddText rs1("articolo_ordine")&testo, "alignment=0; size="&size_chr_pdf&"; Expand=true;"
	Row.Cells(2).border =0
	if rs1("um")<>"" then
		Row.Cells(3).AddText rs1("um"), "alignment=2; size="&size_chr_pdf&"; Expand=true;"
	end if
	Row.Cells(3).border =0
	if rs1("quantita")>0 then Row.Cells(4).AddText rs1("quantita"), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	Row.Cells(4).border =0
	nascondi=true
	if prezzo<>0 then
		Row.Cells(5).AddText formatcurrency(rs1("prezzo"),2), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
		nascondi=false
	end if

	Row.Cells(5).border =0
	testo=""
	if sconto_prodotto<>0 then
		Row.Cells(6).AddText formatnumber(rs1("sconto_prodotto"),2)&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	end if
	Row.Cells(6).border =0
	Row.Cells(7).border =0
	if isnull(prezzo) then prezzo=0
	subtot=(prezzo-(sconto_prodotto/100)*prezzo)*rs1("quantita")
	subtot=round(subtot,2)
	if nascondi=false then
		Row.Cells(7).AddText formatcurrency(subtot,2), "alignment=1; size="&size_chr_pdf&"; Expand=true;"
		Row.Cells(8).AddText iva(ordine.campo("data"))&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"

	end if
	Row.Cells(8).border =0
	totord=totord+subtot
		'Set Row = T_elenco_articoli.Rows.Add(10) ' row height
		'Row.Cells(2).AddText "TOTORDB"&totord, "alignment=0; size=7; Expand=true;"
		
	if spettanze=1 then
		'Cerco nelle spettanze
		txt=""
		conteggio_seriali=0
		sql="select "&m_tabella&"_dett_spettanze.quantita, "&m_tabella&"_dett_spettanze.taglia_misura, utenti_dipendenti.matricola, utenti_dipendenti.nominativo from ("&m_tabella&"_dett_spettanze inner join utenti_dipendenti on "&m_tabella&"_dett_spettanze.iddip=utenti_dipendenti.iddip) where iddett="&RS1("iddett")
        sql=replace(sql,"ordini",m_tabella)

		set rs_spettanze=conn.execute (sql)
		do while not rs_spettanze.eof 
			Set Row = T_elenco_articoli.Rows.Add(10) ' row height
			Row.Cells(1).border =0
			Row.Cells(2).border =0
			Row.Cells(3).border =0
			Row.Cells(4).border =0
			Row.Cells(5).border =0
			Row.Cells(6).border =0
			Row.Cells(7).border =0
			Row.Cells(8).border =0
			Row.Cells(2).AddText rs_spettanze("nominativo")&" Mat. "&rs_spettanze("matricola")&" ("&rs_spettanze("taglia_misura")&")", "alignment=0; size=7; html=true; Expand=true;"
			Row.Cells(4).AddText rs_spettanze("quantita"), "alignment=1; size=7; html=true; Expand=true;"
			
			rs_spettanze.movenext
			
			
			
		loop
		Row.Cells(1).SetBorderParams "Left=False, Right=False, Top=False"
		Row.Cells(2).SetBorderParams "Left=False, Right=False, Top=False"
		Row.Cells(3).SetBorderParams "Left=False, Right=False, Top=False"
		Row.Cells(4).SetBorderParams "Left=False, Right=False, Top=False"
		Row.Cells(5).SetBorderParams "Left=False, Right=False, Top=False"
		Row.Cells(6).SetBorderParams "Left=False, Right=False, Top=False"
		Row.Cells(7).SetBorderParams "Left=False, Right=False, Top=False"
		Row.Cells(8).SetBorderParams "Left=False, Right=False, Top=False"
		if txt<>"" then 
			response.write "<br>"&txt
		end if		
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
	Row.Cells(1).border =0
	Row.Cells(2).AddText "Trasporto", "alignment=0; size=7; Expand=true;"
	Row.Cells(2).border =0
	Row.Cells(3).border =0
	Row.Cells(4).border =0
	Row.Cells(5).AddText formatcurrency(ordine.campo("trasporto"),2), "alignment=1; size=7; Expand=true;"
	Row.Cells(5).border =0
	Row.Cells(6).border =0
	Row.Cells(7).border =0
	Row.Cells(7).AddText formatcurrency(ordine.campo("trasporto"),2), "alignment=1; size=7; Expand=true;"
	Row.Cells(8).border =0
	Row.Cells(8).AddText iva(ordine.campo("data"))&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	totord=totord+trasporto
end if






if sconto_ordine>0 then
	Set Row = T_elenco_articoli.Rows.Add(10) ' row height
	Row.Cells(1).border =0
	Row.Cells(2).AddText "Sconto ordine", "alignment=0; size=7; Expand=true;"
	Row.Cells(2).border =0
	Row.Cells(3).border =0
	Row.Cells(4).border =0
	Row.Cells(5).AddText formatcurrency(ordine.campo("sconto_ordine"),2), "alignment=1; size=7; Expand=true;"
	Row.Cells(5).border =0
	Row.Cells(6).border =0
	Row.Cells(7).border =0
	Row.Cells(7).AddText formatcurrency(ordine.campo("sconto_ordine"),2), "alignment=1; size=7; Expand=true;"
	Row.Cells(8).border =0
	Row.Cells(8).AddText iva(ordine.campo("data"))&"%", "alignment=1; size="&size_chr_pdf&"; Expand=true;"
	sconto_totale=sconto_totale+sconto_ordine
end if

'testo resi
if mod_larochelle and (m_tabella="ordini") then 
	
	Set Row = T_elenco_articoli.Rows.Add(20) ' row height
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





Set T_totali_articoli = Doc.CreateTable("width="&page.width-40&"; height=100; Rows=4; Cols=5; Border=1; CellSpacing=-1; cellpadding=1 ")
T_totali_articoli.Font = Doc.Fonts("Arial")
With T_totali_articoli.Rows(1)
	.Cells(1).Width = (page.width-40)/5+0.75
	'.Cells(1).AddText "SPESE BANCARIE", "alignment=2; size=7; Expand=true;"
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
	.Cells(5).AddText "TOTALE MERCE", "alignment=2; size=7; Expand=true;"
	.cells(5).SetBorderParams "bottom=False" 
	.height=14
End With
With T_totali_articoli.Rows(2)
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
With T_totali_articoli.Rows(3)
	.Cells(1).AddText "SCONTO SU "&ucase(cosa), "alignment=2; size=7; Expand=true;"
	.cells(1).SetBorderParams "bottom=False" 
	.cells(5).SetBorderParams "bottom=False" 
	.Cells(5).AddText "TOTALE "&ucase(cosa), "alignment=2; size=7; Expand=true;"
	.height=14
End With
With T_totali_articoli.Rows(4)
	'.Cells(1).AddText "", "alignment=2; size=7; Expand=true;"
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
Param("y") = page.height-240
Param("MaxHeight") = 455
FirstRow = 2
Do While True
	LastRow = Page.Canvas.DrawTable (T_elenco_articoli, param)
   if LastRow >= T_elenco_articoli.Rows.Count Then Exit Do ' entire table displayed

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
	if pagina=pagine then
		totord=round(totord,2)
		T_totali_articoli.Rows(2).Cells(2).AddText formatcurrency(totord,2), "alignment=1; size=8; Expand=true;" 'IMPONIBILE
		if tabella="ordini" then
			iva_ordine=ordine.campo("iva_ordine")
		else
			iva_ordine=iva(date())
		end if
		if isnull(iva_ordine) then iva_ordine=""
		
		T_totali_articoli.Rows(2).Cells(3).AddText iva_ordine, "alignment=1; size=8; Expand=true;" 'IVA %
		T_totali_articoli.Rows(2).Cells(5).AddText formatcurrency(totord,2), "alignment=1; size=10; Expand=true;" 'TOTALE MERCE

		'Calcolo IVA
		if tabella<>"ordini_fornitori" then
			testo_esenzione=" "&trattamento_iva(ordine.campo("trattamento_iva_ordine"))
			if esenzione_iva(ordine.campo("trattamento_iva_ordine")) then
				valore_iva=0
				if ordine.campo("trattamento_iva_ordine")=10 then
					valore_iva=roundup(totord*iva_ordine/100,2)
					testo_esenzione=""
				end if
			else
				valore_iva=roundup(totord*iva(ordine.campo("data"))/100,2)
				totord=totord+valore_iva
			end if	
			valore_iva=roundup(totord*iva(ordine.campo("data"))/100,2)
			totord=totord+valore_iva
		end if

		T_totali_articoli.Rows(2).Cells(4).AddText formatcurrency(ordine.campo("imposta_ordine"),2), "alignment=1; size=8; Expand=true;" 'IMPOSTA
		
		totord=totord+valore_iva
		T_totali_articoli.Rows(4).Cells(5).AddText formatcurrency(ordine.campo("totale"),2), "alignment=1; size=10; Expand=true;" 'TOTALE ORDINE


	end if
	
	'Table8.Rows(1).Cells(1).AddText "scadenze", "alignment=2; size=8; Expand=true; html=true;"
	'Page.Canvas.DrawText replace(Application("intestazione_ddt"),vbcrlf,"<br>"), "X=20; Y="&Page.Height-20&"; html=true" , doc.fonts("Arial")
	if esiste_logo then Page.Canvas.DrawImage logo, Param_logo

	Page.Canvas.DrawText intestazione_ddt_fat(ordine.campo("data")), "X=20; Y="&Page.Height-20&"; html=true" , doc.fonts("Arial")
	Page.Canvas.DrawTable Table1, "x="& (page.width/2)-1 &", y="&page.height-20
	Page.Canvas.DrawTable Table2, "x="& (page.width/2)-1 &", y="&page.height-39
	Y=147
	Page.Canvas.DrawTable Table3, "x=20, y="&page.height-Y
	'Page.Canvas.DrawTable T_intestazione_articoli, "x=20, y="&page.height-233-3
	'Page.Canvas.DrawTable T_totali_articoli, "x=20, y="&page.height-715-3
	Y=Y+Table3.height-1
	
	'if not ( rs_ordine is Nothing)
	if isobject(Table3bis) then
		Page.Canvas.DrawTable Table3bis, "x=20, y="&page.height-Y
		Y=Y+Table3bis.height
	end if
	'Page.Canvas.DrawTable Table8, "x=20, y="&page.height-217
	Page.Canvas.DrawTable T_intestazione_articoli, "x=20, y="&page.height-Y
	if pagina=pagine then
		Page.Canvas.DrawTable T_totali_articoli, "x=20, y="&page.height-715-3
	end if
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

%>


