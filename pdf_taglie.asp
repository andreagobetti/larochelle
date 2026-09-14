<%
'Verifica chiusure 30_11_2015
'Migliorabile con getrows su rs
%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include file="jsonObject.class.asp"-->

<%
iduser=request("iduser")
if not isnumeric(iduser) or iduser="" then
	call esci(0,0)
end if

	const margine_x=15
	Set Pdf = Server.CreateObject("Persits.Pdf")
	Set Doc = Pdf.CreateDocument
	Set TextParam = PDF.CreateParam
	Set ImageParam = PDF.CreateParam
	Doc.Title = "Taglie dipendenti"
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
		.Cells(1).AddText "ELENCO TAGLIE DIPENDENTI", "size=10; alignment=2;"
		.Cells(1).SetBorderParams "bottom=False"
		.height=25
	End With
	With tabella_titolo.Rows(2)
		'.BGColor = &H000000
		.Cells(1).AddText ucase(denominazioneV), "size=12; alignment=2;"
		.Cells(1).SetBorderParams "top=False"
		.height=25
	End With
	
	
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

	Set tabella__accessori = Doc.CreateTable("width="&int(page.width-(margine_x*2))&"; height=600; Rows=1; Cols=19; Border=0; CellSpacing=-1; cellpadding=0 ")
	tabella__accessori.Font = Doc.Fonts("Arial")
	parametri="size=8;"
	With tabella__accessori.Rows(1)
		'.BGColor = &H000000
		.Cells(1).AddText "<b>Velcro DX </b>",  parametri&" alignment=1; html=true;"
		.Cells(1).Width = 42
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
		
		.Cells(5).Width = 8
		.Cells(5).SetBorderParams "bottom=False; top=False;"


		.Cells(6).AddText "<b>Cintura e passanti </b>",  parametri&" alignment=1; html=true;"
		.Cells(6).Width = 72

		.Cells(7).AddText "SI",  parametri&" alignment=2;"
		.Cells(7).Width = 17
		.Cells(7).BGColor = verde(json("cintura"))

		.Cells(8).AddText "NO",  parametri&" alignment=2;"
		.Cells(8).Width = 17
		if json("cintura")=0 then
			.Cells(8).BGColor = &H99FF99
		else
			.Cells(8).BGColor = &Hffffff
		end if
		
		.Cells(9).Width = 8
		.Cells(9).SetBorderParams "bottom=False; top=False;"
		
		.Cells(10).AddText "<b>Camicia </b>",  parametri&" alignment=1; html=true;"
		.Cells(11).Width = 38
		
		.Cells(11).AddText "Militare",  parametri&" alignment=2;"
		.Cells(11).BGColor = verde(json("camiciam"))
		.Cells(11).Width = 40


		.Cells(12).AddText "Civile",  parametri&" alignment=2;"
		.Cells(12).BGColor = verde(json("camiciac"))
		.Cells(12).Width = 40
		
		
		.Cells(13).Width = 8
		.Cells(13).SetBorderParams "bottom=False; top=False;"
	
		
		
		.Cells(14).AddText "<b>Camiciotto </b>",  parametri&" alignment=1; html=true;"
		.Cells(14).Width = 60
		
		.Cells(15).AddText "Tutto bottoni",  parametri&" alignment=2;"
		.Cells(15).BGColor = verde(json("camiciottob"))
		.Cells(15).Width = 40

		.Cells(16).AddText "Polo",  parametri&" alignment=2;"
		.Cells(16).BGColor = verde(json("camiciottop"))
		.Cells(16).Width = 40
		
		.Cells(17).Width = 8
		.Cells(17).SetBorderParams "bottom=False; top=False;"
		
		.Cells(18).AddText "<b>Interasse fori berretto</b> ",  parametri&" alignment=1; html=true;"
		.Cells(18).Width = 90
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




	'Page.Canvas.DrawText "Heigt:"&page.height&" width:"&page.width, "X=350; Y="&Page.Height-18&";" , doc.fonts("Arial")
	tabella_taglie_Y=param_y-10
	Set tabella_taglie = Doc.CreateTable("width="&int(page.width-(margine_x*2))&"; height=600; Rows=1; Cols=26; Border=1; CellSpacing=-1; cellpadding=0 ")
	tabella_taglie.Font = Doc.Fonts("Arial")
	parametri_titolo="size=6; alignment=2;"
	With tabella_taglie.Rows(1)
		'.BGColor = &H000000
		.Cells(1).AddText "Dipendente", parametri_titolo
		
		.Cells(2).AddText "Giacca"&vbcrlf&"tecnica" , parametri_titolo
		.Cells(2).BGColor = &Ha8dcf0
		
		.Cells(3).AddText "Giacca", parametri_titolo
		.cells(3).colspan=7
		.Cells(3).BGColor = &H87CEEB
		
		.Cells(10).AddText "Cappotto", parametri_titolo
		.Cells(10).BGColor = &H66c1e5

		.Cells(11).AddText "Pantaloni"&vbcrlf&"tecnici", parametri_titolo
		.Cells(11).BGColor = &HF2F5A9
		
		.Cells(12).AddText "Pantaloni", parametri_titolo
		.cells(12).colspan=7
		.Cells(12).BGColor = &HFFFF00
		
		.Cells(19).AddText "Gonna", parametri_titolo
		.Cells(19).BGColor = &HF4FA58
		
		.Cells(20).AddText "Altro", parametri_titolo
		.cells(20).colspan=8
		
		.height=25
	End With
	
	parametri_titolo="size=5; alignment=2;"
	Set Row = tabella_taglie.Rows.Add(10) ' row height
	with row
		.Cells(1).Width = 90
		
		.Cells(2).AddText "Taglia", parametri_titolo
		.Cells(2).Width = 28

		.Cells(3).AddText "TG", parametri_titolo
		.Cells(3).Width = 21
		
		.Cells(4).AddText "Fondo", parametri_titolo
		.Cells(4).Width = 21
		
		.Cells(5).AddText "Manica", parametri_titolo
		.Cells(5).Width = 21
		
		.Cells(6).AddText "Torace", parametri_titolo
		.Cells(6).Width = 21
		
		.Cells(7).AddText "Vita", parametri_titolo
		.Cells(7).Width = 21
		
		.Cells(8).AddText "Bacino", parametri_titolo
		.Cells(8).Width = 21
		
		.Cells(9).AddText "Spalle", parametri_titolo
		.Cells(9).Width = 21
		
		.Cells(10).AddText "Lunghezza", parametri_titolo
		.Cells(18).Width = 28
		.Cells(11).AddText "Tecnici", parametri_titolo
		
		.Cells(12).AddText "TG", parametri_titolo
		.Cells(12).Width = 21
		
		.Cells(13).AddText "Lunghezza", parametri_titolo
		.Cells(13).Width = 26
		
		.Cells(14).AddText "Vita", parametri_titolo
		.Cells(14).Width = 21
		
		.Cells(15).AddText "Bacino", parametri_titolo
		.Cells(15).Width = 21
		
		.Cells(16).AddText "Cosce", parametri_titolo
		.Cells(16).Width = 21
		
		.Cells(17).AddText "Cavallo", parametri_titolo
		.Cells(17).Width = 24
		
		.Cells(18).AddText "Polpacci", parametri_titolo
		.Cells(18).Width = 26
		
		.Cells(19).AddText "Lunghezza", parametri_titolo
		.Cells(18).Width = 28
		
		.Cells(20).AddText "Scarpa", parametri_titolo
		.Cells(20).Width = 24
		
		.Cells(21).AddText "Berretto", parametri_titolo
		.Cells(21).Width = 24
		
		.Cells(22).AddText "Guanto", parametri_titolo
		.Cells(22).Width = 24
		
		.Cells(23).AddText "Cintura", parametri_titolo
		.Cells(23).Width = 24
		
		.Cells(24).AddText "Cinturone", parametri_titolo
		.Cells(24).Width = 26
		
		.Cells(25).AddText "Maglieria", parametri_titolo
		.Cells(25).Width = 26
		
		.Cells(26).AddText "Arma", parametri_titolo
		.Cells(26).Width = 120
		
	end with

	sql="select *, utenti_gradi.nome_grado,utenti_gradi.colore from  dipendenti left join utenti_gradi on dipendenti.grado = utenti_gradi.id inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.iduser="&iduser &" order by sesso desc,cognome, nome"
	set rs_dipendenti=conn.execute (sql)
	do while not rs_dipendenti.eof 
		iddip=rs_dipendenti("iddip")
		set rs_rilievi=conn.execute ("select dipendenti_misure.* from dipendenti_misure where iddip="&iddip&" order by data_rilievo desc limit 0,1")
		if not rs_rilievi.eof then
		
			parametri_titolo="size=7; alignment=2; html=false; color=black;"
			Set Row = tabella_taglie.Rows.Add(30) ' row height
			with row
				.Cells(1).AddText "<b>"&dipendente_colorato(rs_dipendenti("cognome")&" "&rs_dipendenti("nome"),rs_dipendenti("sesso")) &"</b><br>"&grado_colorato(rs_dipendenti("nome_grado"),rs_dipendenti("colore")), "size=7; html=true;"
				.Cells(2).AddText rs_rilievi("giacca_tecnici"), parametri_titolo
				.Cells(3).AddText rs_rilievi("giacca_TG")&vbcrlf&rs_rilievi("giacca_tg2"), parametri_titolo
				.Cells(4).AddText rs_rilievi("giacca_fondo")&vbcrlf&rs_rilievi("giacca_fondo2"), parametri_titolo
				.Cells(5).AddText rs_rilievi("giacca_manica")&vbcrlf&rs_rilievi("giacca_manica2"), parametri_titolo
				.Cells(6).AddText rs_rilievi("giacca_Torace")&vbcrlf&rs_rilievi("giacca_Torace2"), parametri_titolo
				.Cells(7).AddText rs_rilievi("giacca_vita")&vbcrlf&rs_rilievi("giacca_vita2"), parametri_titolo
				.Cells(8).AddText rs_rilievi("giacca_bacino")&vbcrlf&rs_rilievi("giacca_bacino2"), parametri_titolo
				.Cells(9).AddText rs_rilievi("giacca_spalle")&vbcrlf&rs_rilievi("giacca_spalle2"), parametri_titolo
				.Cells(10).AddText rs_rilievi("cappotto")&vbcrlf&rs_rilievi("cappotto2"), parametri_titolo
				.Cells(11).AddText rs_rilievi("pantaloni_tecnici"), parametri_titolo
				.Cells(12).AddText rs_rilievi("pantaloni_tg")&vbcrlf&rs_rilievi("pantaloni_tg2"), parametri_titolo
				.Cells(13).AddText rs_rilievi("pantaloni_lunghezza")&vbcrlf&rs_rilievi("pantaloni_lunghezza2"), parametri_titolo
				.Cells(14).AddText rs_rilievi("pantaloni_vita")&vbcrlf&rs_rilievi("pantaloni_vita2"), parametri_titolo
				.Cells(15).AddText rs_rilievi("pantaloni_bacino")&vbcrlf&rs_rilievi("pantaloni_bacino2"), parametri_titolo
				.Cells(16).AddText rs_rilievi("pantaloni_cosce")&vbcrlf&rs_rilievi("pantaloni_cosce2"), parametri_titolo
				.Cells(17).AddText rs_rilievi("pantaloni_cavallo")&vbcrlf&rs_rilievi("pantaloni_cavallo2"), parametri_titolo
				.Cells(18).AddText rs_rilievi("pantaloni_polpacci")&vbcrlf&rs_rilievi("pantaloni_polpacci2"), parametri_titolo
				.Cells(19).AddText rs_rilievi("gonna")&vbcrlf&rs_rilievi("gonna2"), parametri_titolo
				.Cells(20).AddText rs_rilievi("scarpa"), parametri_titolo
				.Cells(21).AddText rs_rilievi("berretto"), parametri_titolo
				.Cells(22).AddText rs_rilievi("guanto"), parametri_titolo
				.Cells(23).AddText rs_rilievi("cintura"), parametri_titolo
				.Cells(24).AddText rs_rilievi("cinturone"), parametri_titolo
				.Cells(25).AddText rs_rilievi("maglieria"), parametri_titolo
				.Cells(26).AddText rs_dipendenti("arma"), "size=5;"
			end with
		else
			Set Row = tabella_taglie.Rows.Add(30) ' row height
			with row
				.Cells(1).AddText "<b>"&dipendente_colorato(rs_dipendenti("cognome")&" "&rs_dipendenti("nome"),rs_dipendenti("sesso")) &"</b><br>"&grado_colorato(rs_dipendenti("nome_grado"),rs_dipendenti("colore")), "size=6; html=true;"
				.Cells(2).AddText "NESSUN RILIEVO", parametri_titolo
				.cells(2).colspan=25
			end with
		end if
		rs_dipendenti.MoveNext
	loop
	
	Set Param = Pdf.CreateParam
	
	param.clear
	Param("x") = margine_x
	Param("y") = page.height-75-tabella__accessori.height
	Param("MaxHeight") =page.height-75-20
	FirstRow = 3
	pagina=1
	Do While True
		if pagina=1 then
			Param("y") = page.height-75-tabella__accessori.height
		else
			Param("y") = page.height-75
		end if
		LastRow = Page.Canvas.DrawTable (tabella_taglie, param)
	   if LastRow >= tabella_taglie.Rows.Count Then Exit Do ' entire table displayed
	   ' Display remaining part of table on the next page
	   Set Page = Page.NextPage
	   Param.Add( "RowTo=2; RowFrom=1" ) ' Row 1 is header.
	   Param("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
	   FirstRow = LastRow + 1
	   pagina=pagina+1
	Loop



	pagine=doc.pages.count
	call add2log("pagine:"&pagine,0)

	for pagina=1 to pagine
		Set Page = Doc.Pages.item(pagina)
		Page.Canvas.DrawTable tabella_titolo, "x="&tabella_titolo_X&", y="&replace(tabella_titolo_Y,",",".")
		'Page.Canvas.DrawTable tabella_taglie, "x="&margine_x&", y="&replace(tabella_taglie_Y,",",".")
		if esiste_logo then
			Page.Canvas.DrawImage logo, parametri_logo
		
		end if
		if pagina=1 then
			Page.Canvas.DrawTable tabella__accessori, "x="&margine_x&", y="&replace(tabella_accessori_Y,",",".")
		end if
		
		Page.Canvas.DrawText "Pagina "&pagina&"/"&pagine, "X="&page.width-100&"; Y=20;" , doc.fonts("Arial")

		
	next
	set rs_rilievi = Nothing
	set rs_dipendenti = Nothing
	call connclose()



	Doc.SaveHttp "attachment;filename=Taglie_"&replace(denominazioneV," ","_")&".pdf" 

	Set Doc = Nothing
	Set TextParam = Nothing
	Set ImageParam = Nothing
	Set Param_logo = Nothing
	Set Pdf = Nothing
	
	
	function dipendente_colorato(nominativo,sesso)
		if sesso="F" then 
			dipendente_colorato="<font color=""#f74ca5"">"&nominativo&"</font>"
		else
			dipendente_colorato=nominativo
		end if
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
%>


