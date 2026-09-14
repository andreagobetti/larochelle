<%
'Verifica chiusure 30_11_2015
'Migliorabile con getrows su rs
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<%
idddt=request("idddt")
if not isnumeric(idddt) or idddt="" then
	call esci(0,0)
end if
if session("iduser")="" then
	add2log "Tentativo di visualizare PDF DDT da utente non loggato"&vbcrlf&allCookies(),3
	call esci(7,idddt)
end if

sql="select ddt.*, ordini.idord, ordini.nord, ordini.data, ordini.iduser, ordini.idintestazione, ordini.idconsegna, ordini.cig, ordini.mepa_testo, ordini.mepa_tipo, ordini.mepa_data, ordini.impegno_spesa, utenti.cellulare, utenti.telefono from ordini inner join ddt on ordini.idord = ddt.idord inner join utenti on ordini.iduser = utenti.iduser where ordini.idord="&idddt
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,conn
if not rs.eof then
	if session("idadmin")="" and request("id2")<>chkDataUser(rs("data")) then
		add2log "Tentativo di visualizare PDF DDT, id2 incongruente:"&request("id2")&vbcrlf&allCookies(),3
		call esci(7,idddt)
	end if
	if session("idadmin")="" then
		denominazione_v=get_denominazione(sessioniduser)
		'conn.execute ("update fatture set download=now() where idfat="&idfat)
		add2log "[utente="&sessioniduser&"]"&denominazione_v&"[/utente] ha scaricato il [ddt="&rs("idddt")&"]"&rs("nddt")&"[/ddt]",2
	end if
	Set Pdf = Server.CreateObject("Persits.Pdf")
	Set Doc = Pdf.CreateDocument
	Set TextParam = PDF.CreateParam
	Set ImageParam = PDF.CreateParam
	nddt=rs("nord")
	Doc.Title = "D.d.T. "&nddt
	Doc.Creator = nomesito
	Bordo_sx=25
	Set Page = Doc.Pages.Add
	page.height=825
	Set Param = Pdf.CreateParam
	Set Param2 = Pdf.CreateParam

	'TABELLA 1
	Set Table1 = Doc.CreateTable("width=300; height=600; Rows=1; Cols=3; Border=1; CellSpacing=-1; cellpadding=2 ")
	Table1.Font = Doc.Fonts("arial")
	'table1.rows.add(15)
	With Table1.Rows(1)
		.BGColor = &HCCCCCC
		.Cells(1).Width = page.width/6-7
		.Cells(1).AddText "D.d.T. n° "&pre_fattura(rs("data"))& aggiungi_zeri(nddt), Param
		.Cells(2).Width = page.width/6-7
		.Cells(2).AddText "del "& formatDateTime(rs("data"), vbShortDate) , Param
		.Cells(3).Width = page.width/6-6
		.Cells(3).AddText "Pagina "&doc.pages.count' di 1", ""
		.height=20
	End With
	'Page.Canvas.DrawTable Table1, "x="& page.width/2 &", y="&page.height-20
	'TABELLA 2
	Set Table2 = Doc.CreateTable("width=300; height=600; Rows=4; Cols=2; Border=1; CellSpacing=-1; cellpadding=2 ")
	Table2.Font = Doc.Fonts("arial")
	'table1.rows.add(15)
	param.add "html=true"
	Set ccausale_ddt = new cl_causale_ddt
	With Table2.Rows(1)
		'.BGColor = &HCCCCCC
		.Cells(1).Width = page.width/3-15
		.Cells(1).AddText "CAUSALE<br><b>"&ucase( ccausale_ddt.elemento(rs("causale")))&"</b>", param
		.Cells(2).Width = page.width/6-6
		.Cells(2).AddText "COD.CLIENTE<br><b>"& rs("iduser")&"</b>", param
		.height=30
	End With
	set ccausale_ddt = Nothing
	With Table2.Rows(2)
		'.BGColor = &HCCCCCC
		.Cells(1).AddText "PORTO<br><b>"&ucase( porto_ddt(rs("porto")))&"</b>", "size=8; html=true;"
		.Cells(2).AddText "NUMERO COLLI<br><b>"& rs("colli")&"</b>", Param
		.height=30
	End With
	With Table2.Rows(3)
		'.BGColor = &HCCCCCC
		.Cells(1).AddText "IMBALLO<br><b>"&ucase( imballo_ddt(rs("imballo")))&"</b>", Param
		.Cells(2).AddText "PESO (Kg)<br><b>"& rs("peso")&"</b>", Param
		.height=30
	End With
	With Table2.Rows(4)
		'.BGColor = &HCCCCCC
		if rs("tipo_ddt")<>"no_ordine" then
			set rs_ordine=conn.Execute("select ordini.nord, ordini.data from ordini where idord="&rs("sub_idord"))
			.Cells(1).AddText "RIFERIMENTO ORDINE<br><b>"&rs_ordine("nord")&" del "&formatDateTime(rs_ordine("data"), vbShortDate) &"</b>", Param
			set rs_ordine = nothing
		end if
		.Cells(2).AddText "DIMENSIONE (cm)<br><b>"& rs("dimensione")&"</b>", Param
		.height=30
	End With
	'Page.Canvas.DrawTable Table2, "x="& page.width/2 &", y="&page.height-45
	'TABELLA 3
	Set Table3 = Doc.CreateTable("width="&page.width-40&"; height=600; Rows=1; Cols=2; Border=1; CellSpacing=-1; cellpadding=2 ")
	Table3.Font = Doc.Fonts("arial")
	'table1.rows.add(15)
	With Table3.Rows(1)
		'.BGColor = &HCCCCCC
		.Cells(1).Width = (page.width-40)/2
		telefono=""
		if rs("cellulare")<>"" then
			telefono=rs("cellulare")
		end if
		if rs("telefono")<>"" then
			call concatena_stringa( telefono,",",rs("telefono"))
		end if
		if telefono<>"" then telefono="<br>Tel: "&telefono
		.Cells(2).AddText "SPETT.LE "&dati_fatturazione_pdf(rs("idintestazione"))&telefono, "size=9; html=true; Expand=true;"
		
		.Cells(2).Width = (page.width-40)/2
		.Cells(1).AddText "LUOGO DI DESTINAZIONE<br>"&dati_consegna(rs("idconsegna"))&"", Param
		.height=70
	End With
	'Page.Canvas.DrawTable Table3, "x=20, y="&page.height-170

	Set Table4 = Doc.CreateTable("width="&page.width-40&"; height=300; Rows=1; Cols=4; Border=1; CellSpacing=-1; cellpadding=2; ")
	Table4.Font = Doc.Fonts("arial")
	With Table4.Rows(1)
		'.BGColor = &HCCCCCC
		.Cells(1).Width = 80
		.Cells(1).border =0
		'.Cells(1).AddText "COD. ARTICOLO", "alignment=2; size=8;"
		.Cells(2).Width = page.width-40-80-60-60+2
		.Cells(2).border =0
		'.Cells(2).AddText "DESCRIZIONE", "alignment=2; size=8;"
		.Cells(3).Width = 60
		.Cells(3).border =0
		'.Cells(3).AddText "U.M.", "alignment=2; size=8;"
		.Cells(4).Width = 60
		.Cells(4).border =0
		'.Cells(4).AddText "QUANTITA'", "alignment=2; size=8;"
		.height=20
	End With
	table4.border=0
	parziale=rs("parziale")

	if rs("tipo_ddt")="completo" then
		sql="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara where idord="&rs("sub_idord")&" order by ordine, iddett;"
	elseif rs("tipo_ddt")="parziale" then
		sql="select ordini_dett.*,ddt_dett_ordini.quantita, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara inner join ddt_dett_ordini on ordini_dett.iddett = ddt_dett_ordini.iddett  where ddt_dett_ordini.idddt="&rs("idord")&" order by ordine, iddett;"
		
		
	else
		sql="select ordini_dett.*, varianti_a.codicevara, varianti_a.variante_a, varianti_b.variante_b from (ordini_dett left join varianti_b on ordini_dett.idvarb = varianti_b.idvarb) left join varianti_a on ordini_dett.idvara = varianti_a.idvara where idord="&rs("idord")&" order by ordine, iddett;"
	end if






	Set rs1 = Server.CreateObject("adodb.recordset")
	rs1.open sql,conn
	

	n=2
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
		if not isnull(rs1("um")) then	
			Row.Cells(3).AddText rs1("um"), "alignment=2; size="&size_chr_pdf&"; Expand=true;"
		end if
		Row.Cells(3).border =0
		Row.Cells(4).AddText rs1("quantita"), "alignment=2; size="&size_chr_pdf&"; Expand=true;"
		Row.Cells(4).border =0
		rs1.movenext
		n=n+1
	Loop
	set rs1 = Nothing



	if mod_larochelle then 
		
		Set Row = Table4.Rows.Add(10) ' row height
		Row.Cells(1).border =0
		Row.Cells(2).AddText txt_restituzione, "alignment=0; html=true; size=8; Expand=true;"
		Row.Cells(2).border =0
		Row.Cells(3).border =0
		Row.Cells(4).border =0
	end if 





	'Page.Canvas.DrawTable Table4, "x=20, y="&page.height-235 & ", MaxHeight= 460"

	Set Table5 = Doc.CreateTable("width="&page.width-40&"; height=600; Rows=1; Cols=4; Border=1; CellSpacing=-1; cellpadding=2 ")
	Table5.Font = Doc.Fonts("arial")
	With Table5.Rows(1)
		.BGColor = &HCCCCCC
		.Cells(1).Width = 80
		.Cells(1).AddText "COD. ARTICOLO", "alignment=2; size=8;"
		.Cells(2).Width = page.width-40-80-60-60+2
		.Cells(2).AddText "DESCRIZIONE", "alignment=2; size=8;"
		.Cells(3).Width = 60
		.Cells(3).AddText "U.M.", "alignment=2; size=8;"
		.Cells(4).Width = 60
		.Cells(4).AddText "QUANTITA'", "alignment=2; size=8;"
		.height=20
	End With
	Table5.Rows.Add(440)
	Table5.Rows.Add(40)
	table5.rows(3).cells(1).colspan=4
	'Page.Canvas.DrawTable Table5, "x=20, y="&page.height-235
	'TABELLA 6
	Set Table6 = Doc.CreateTable("width="&page.width-40&"; height=30; Rows=1; Cols=4; Border=1; CellSpacing=-1; cellpadding=2 ")
	Table6.Font = Doc.Fonts("arial")
	Table6.Rows.Add(30)
	With Table6.Rows(2)
		.Cells(1).Width = (page.width-40)/2
		.cells(2).colspan=3
	End With




	param.clear
	Param("x") = 20
	Param("y") = page.height-235 
	Param("MaxHeight") = 460
	FirstRow = 2
	Do While True
		LastRow = Page.Canvas.DrawTable (Table4, param)
		
		
		' INIZIO BLOCCO PAGINA RIPETUTO
		'Page.Canvas.DrawText intestazione_ddt_fat(rs("data")), "X=20; Y="&Page.Height-20&"; html=true" , doc.fonts("arial")
		call intestazione_pdf_ddt_fat(doc, page,rs("data"))


		'FINE BLOCCO PAGINA RIPETUTO
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
		With Table6.Rows(1)
		'.BGColor = &HCCCCCC
		.Cells(1).Width = (page.width-40)/2
		.Cells(2).Width = 40
		.Cells(3).Width = 120
		testo=""
		.Cells(4).Width = (page.width-40)/2-40-120+2
		'.height=60
	End With
		if pagina=pagine then
			if rs("incaricato")=3 then
				incaricato=ucase(rs("vettore"))
			else
				incaricato=ucase(incaricato_ddt(rs("incaricato")))
			end if
			if isdate(rs("data_inizio")) then
				inizio= formatDateTime(rs("data_inizio"), vbShortDate)
			else
				inizio=""
			end if
			annotazioni=rs("annotazioni")&dati_mepa_pdf(rs)
			
		end if
		
		Table6.Rows(1).Cells(1).AddText "<center>INCARICATO DEL TRASPORTO<br><b>"&incaricato&"</b></center>", "alignment=2; size=8; Expand=true; html=true;"
		Table6.Rows(1).Cells(3).AddText "<center>DATA INIZIO TRASPORTO<br><b>"&inizio &"</b></center>", "alignment=2; size=8; Expand=true; html=true;"
		Table6.Rows(1).Cells(2).AddText "ORA", "alignment=2; size=8; Expand=true;"
		Table6.Rows(1).Cells(4).AddText "FIRMA CONDUCENTE", "alignment=2; size=8; Expand=true;"

		Table6.Rows(2).Cells(1).AddText "<b>DATA RICEVIMENTO MERCE", "alignment=0; size=10; html=true;"
		Table6.Rows(2).Cells(2).AddText "<b>FIRMA DESTINATARIO</b>", "alignment=0; size=10; html=true;"

		table5.rows(3).Cells(1).AddText "ANNOTAZIONI<br>"&annotazioni, "alignment=0; size=8; html=true;"

		
		
		
		
		

		table1.rows(1).Cells(3).cleartext
		table1.rows(1).Cells(3).AddText "Pagina "&doc.pages.count' di 1", ""

		
		Page.Canvas.DrawTable Table1, "x="& page.width/2 &", y="&page.height-20
		Page.Canvas.DrawTable Table2, "x="& page.width/2 &", y="&page.height-42
		Page.Canvas.DrawTable Table3, "x=20, y="&page.height-161
		Page.Canvas.DrawTable Table5, "x=20, y="&page.height-235
		Page.Canvas.DrawTable Table6, "x=20, y="&page.height-740

	next
	Doc.SaveHttp "attachment;filename=DDT_"&pre_fattura(rs("data"))&nddt&".pdf" 

	set rs = Nothing
	call connclose()
	Set Doc = Nothing
	Set TextParam = Nothing
	Set ImageParam = Nothing
	Set Param_logo = Nothing
	Set Pdf = Nothing
	call CheckConnChiusa()
else
	call esci(6,idddt)
end if

function dati_mepa_pdf(rs)
	dati_mepa_pdf=""
			
        if rs("cig")<>"" then
	        call concatena_stringa(dati_mepa_pdf," - ","Codice CIG: "&rs("cig"))
		end if
        if rs("mepa_tipo")>0 then
	        call concatena_stringa(dati_mepa_pdf," - ",mepa_tipo( rs("mepa_tipo"))&": "&rs("mepa_testo"))
		end if
        if rs("mepa_data")<>"" then
	        call concatena_stringa(dati_mepa_pdf," - ","Data ordine: "&rs("mepa_data"))
		end if
        

		if rs("impegno_spesa")<>"" then
	        call concatena_stringa(dati_mepa_pdf," - "," Impegno di spesa: "&rs("impegno_spesa"))
		end if

	
end function
%>


