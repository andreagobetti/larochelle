<%
'Verifica chiusure 04_12_2015
%>
<!--#include virtual="/setup.asp" -->
<%
idpro=request("idpro")
Set fs = Server.CreateObject("Scripting.filesystemObject")

Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument

Set TextParam = PDF.CreateParam
Set ImageParam = PDF.CreateParam
sql = "select * FROM impostazioni"
set rs=conn.execute(sql)
rs.movefirst
fondo_listino=rs("fondo_listino")
rs.close


Doc.Title = "Scheda articolo"
Doc.Creator = nomesito
Bordo_sx=25
Set Page = Doc.Pages.Add
Set Param = Pdf.CreateParam
Set Param2 = Pdf.CreateParam
'testata
Set Param_testata = PDF.CreateParam
Param_testata("x") =Bordo_sx' (Page.Width - Image.Width / 3) / 2
on error resume next
if fs.FileExists(Server.MapPath( Application("img_email"))) then
	Set Image_testata = Doc.OpenImage( Server.MapPath( Application("img_email")) )
	'calcolo misure dell'immagine considerando la risoluzione DPI
	larghezza=Image_testata.width*72/Image_testata.resolutionx
	altezza=Image_testata.height*72/Image_testata.resolutiony
	'calcolo rapporto di riduzione X e Y
	scalaX=200/larghezza
	scalaY=40/altezza
	'Applico il rapporto predominante
	if scalaX<scalaY then
		scala=scalaX
	else
		scala=scalaY
	end if
	
	'dimemsione immagine testata
	Param_testata("ScaleX") = scala
	Param_testata("ScaleY") = scala

	Param_testata("y") = Page.Height  -20-altezza*scala
	if err.number=0 then

	immagine_testata=true
	end if
end if
on error goto 0
Set Table = Doc.CreateTable("width=562; height=5; Rows=1; Cols=2; Border=1; CellSpacing=-1; cellpadding=2 ")
Table.Font = Doc.Fonts("Arial")

Param.Set("alignment=center")



' Populate table with data
Set rs = Server.CreateObject("adodb.recordset")
sql="select * from prodotti where idpro="&idpro 
rs.Open sql, Conn
codice=rs("codice")
param.Set "expand=true; size=10" ' expand cell vertically
n=0
n=n+1


'Set Row = Table.Rows.Add(20) ' row height
'Row.Cells(1).AddText rs("articolo"), "alignment=0; size=15; expand=true;"




'Set Row = Table.Rows.Add(5) ' row height
'Prima riga
Set row = Table.Rows(1)
'Row.Cells(1).Width = 562
Row.Cells(1).AddText "<b>"&rs("articolo")&"</b>", "alignment=0; size=12; expand=true; html=true;"
row.cells(1).border=0
row.cells(1).colspan=2

'Seconda riga (immagine)
Set Row = Table.Rows.Add(20) ' row height
row.height=220
row.cells(1).colspan=2
row.cells(1).border=0
'row.Cells(1).BgColor = "brown"

'Immagine
fileimg=rs("fileimg")
if fileimg="nofoto" or isnull(fileimg) then
	txt="nofoto.gif"
end if
   	if fileimg<>"" then
		if fs.FileExists(Server.MapPath( upl_img_cat & fileimg)) then
			Set ImageParam = PDF.CreateParam
			set Image = Doc.OpenImage( Server.MapPath( upl_img_cat & fileimg) )
			
			'calcolo misure dell'immagine considerando la risoluzione DPI
			larghezza=image.width*72/image.resolutionx
			altezza=image.height*72/image.resolutiony
			
			margine=10
			
			scala=(row.height-margine)/altezza
			
			'calcolo nuova larghezza considerando la scala per centrare poi l'immagine
			nuova_larghezza=larghezza*scala
			
			'Trovo la x per centrare l'immagine
			ImageParam("x") =  (row.cells(1).width - nuova_larghezza ) / 2
			
			'Sotto laschio metà del margine
			ImageParam("y") = margine/2

			ImageParam("ScaleX") = scala
			ImageParam("ScaleY") = scala
			
			
			'Row.Cells(1).AddText "immagine: W"&image.width&"*H"&image.height&"DPI: "&image.resolutionx&"larghezz:"&larghezza&" altezza: "&altezza&" scala:"& scala & "larghezza immagine: "&(image.width*scalex)&" scala_txt:"& scalex_txt&"ImageParam.X:"&ImageParam("X")&"ImageParam.Y:"&ImageParam("Y")&"ImageParam.scaleX:"&ImageParam("scalex")&"ImageParam.scaleY:"&ImageParam("scaleY")&" altezza riga"&row.height&" larghezza riga:"&row.cells(1).width, "alignment=0; size=8; html=true; expand=true;"

			Row.Cells(1).canvas.drawimage image, ImageParam
			
			set ImageParam=nothing
			set image=nothing
	   end if
   end if
   set fs = Nothing	
	'ID articolo
	Set Row = Table.Rows.Add(20) ' row height
	Row.Cells(1).Width = 460
	Row.Cells(2).Width = 562-Row.Cells(1).Width-3
	Row.Cells(1).AddText "ID articolo: "&rs("idpro"), "alignment=0; size=10; expand=true;"
	
	'Codice
	Set Row = Table.Rows.Add(20) ' row height
	Row.Cells(1).AddText "Codice: "&rs("codice"), "alignment=0; size=10; expand=true;"
	'Prezzo
	Set Row = Table.Rows.Add(20) ' row height
	
	'Decisionale visualizza il prezzo
	prezzo_visibile=rs("prezzo_visibile")
	
	if prezzo_visibile=0  then
		'Visualizza il prezzo
		vedi_prezzo=1
		
		
	else 
		
		'Telefonare
		vedi_prezzo=0
		

	end if 

	
	
	
	
	if vedi_prezzo=0 then
	   Row.Cells(1).AddText "Prezzo: Contattaci","alignment=0;size=10;expand=true;"
	else
		
		prezzo_minimo= converti_typevar14(rs("prezzo_ve_min"))
		'articolo=server.HTMLEncode(articolo)
		if not isnull(prezzo_minimo) and prezzo_minimo>0 then 
			Row.Cells(1).AddText "Prezzo: da "&formatcurrency(prezzo_minimo,2,true), "alignment=0;size=10;expand=true; html=true;"
		else
		   Row.Cells(1).AddText "Prezzo: "&formatcurrency(rs("prezzo"),2,true) , "alignment=0;size=10;expand=true; html=true;"
		end if
   		if cdbl(rs("fascia_di_sconto"))>0 then
		   	
			Row.Cells(1).canvas.SetParams "FillColor =&H"&mid(colore_sconto(rs("fascia_di_sconto")),2)&""
			Row.Cells(1).canvas.FillRect 100,5,10,10
			fascia=mid(colore_sconto(rs("fascia_di_sconto")),2)
		end if

   end if

	
	
	
	
	
	
	'Disponibilità
	Set Row = Table.Rows.Add(20) ' row height
	Row.Cells(1).AddText "Disponibilità: "&disponibilita(RS("disponibilita"),lingua), "alignment=0; size=10; expand=true;"

	'Testo qr
	Set Row = Table.Rows.Add(20) ' row height
	Row.Cells(1).AddText "Utilizza il codice QR per visualizzare l'articolo sul nostro sito", "alignment=1; size=10; expand=true;"
	Row.Cells(1).SetBorderParams "right=false;"
	
	'QR
	Set row = Table.Rows(3)
	'Row.Cells(2).AddText "QR", "alignment=0; size=12; expand=true;"
	row.cells(2).rowspan=5
	testo_qr="http://"&nomesito&"/qr.asp?idpro="&rs("idpro")
	testo="<b>"&rs("codice")& "</b> "&rs("articolo")
	'table.rows(riga).cells(2).addtext testo,"size=10; html=true;"
	'table.rows(riga+1).cells(2).addtext formatdatetime(date(),2)&"  "&formatcurrency(rs("prezzo"),2)," Alignment =1; size=8;"
	row.cells(2).Canvas.DrawBarcode2D testo_qr, "x=5; y=5; type=3; BarWidth=3; " 
	'row.cells(2).Canvas.DrawBarcode2D testo_qr, "x=10; y=10; type=3;  " 
	Row.Cells(2).SetBorderParams "left=false;"
	
	
	
  	'descrizione
   param.Add "size=8"
   param.Add "html=true"
   'Set Row = Table.Rows.Add(20) ' row height
   'if len(rs("descrizione_listino"))>0 then
	'   Row.Cells(1).AddText "<b>"&rs("articolo")&"</b><br>"&replace(rs("descrizione_listino"),vbcrlf,"<br>"), "alignment=0; size=8; html=true; expand=true;"
	'else
	'   Row.Cells(1).AddText "<b>"&rs("articolo")&"</b><br>", "alignment=0; size=8; html=true; expand=true;"
	'end if
	
	'Descrizione
	Set Row = Table.Rows.Add(20) ' row height
	row.cells(1).colspan=2
	Row.Cells(1).AddText rs("descrizione"), "alignment=0; size=8; expand=true; html=true;"
	
	'Dettagli articolo
	sql_dettagli="select * from prodotti_dettagli where inschedaarticolo and idpro="&idpro &" order by ordine"
	set rs_dettagli=conn.execute(sql_dettagli)
	do while not rs_dettagli.eof
		Set Row = Table.Rows.Add(20) ' row height
		
		Row.Cells(1).AddText "<b>"&rs_dettagli("titolo")&"</b><br>"&replace(rs_dettagli("testo"),"<br />","<br>"), "alignment=0; size=8; html=true; expand=true;"
		row.cells(1).colspan=2
		
		rs_dettagli.movenext
	loop
	set rs_dettagli = Nothing
	'if row.height<image.height/2+10 then row.height=image.height/2+10 'corregge altezza riga in base all'immagine

' Render table on document
'Set Page = Doc.Pages.Add
set rs = nothing
call connclose()

Param.Clear 
Param("x") = Bordo_sx'(Page.Width - Table.Width) / 2 ' center table on page
Param("y") = Page.Height - 80
Param("MaxHeight") = 700

FirstRow = 2
Do While True
	LastRow = Page.Canvas.DrawTable( Table, Param )
	' Pagina
	Page.Canvas.DrawText ucase(nomesito), "X=350; Y="&Page.Height-18&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "Stampato il: "&date(), "X=350; Y="&Page.Height-42&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "I prezzi sono iva esclusa e possono variare senza preavviso secondo l'andamento del mercato", "X=25; Y="&Page.Height-65&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "Pagina "&doc.pages.count, "X=20; Y=25;" , doc.fonts("Arial")
	if fondo_listino<>"" then
		Page.Canvas.DrawText fondo_listino,  "x=150; y=25; width=350; alignment=center; size=8;html=true", doc.fonts("Arial")
	end if
	if immagine_testata=true then Page.Canvas.DrawImage Image_testata, Param_testata

   if LastRow >= Table.Rows.Count Then Exit Do ' entire table displayed

   ' Display remaining part of table on the next page
   Set Page = Page.NextPage
   Param.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
   Param("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
   FirstRow = LastRow + 1
Loop

Doc.SaveHttp "attachment;filename="&codice&".pdf" 
set Doc = Nothing	
Set Pdf = Nothing
Set Page = Nothing
Set row = Nothing
call CheckConnChiusa()

%>


<%
function RemoveTags(txt)
  'memorizza il testo in un buffer temporaneo
  dim tmptxt
  tmptxt = txt

  'esci se viene passata una stringa nulla (che è diverso da stringa di lunghezza 0)
  if IsNull(tmptxt) then
    exit function
  end if

  dim i, pos1, pos2
  'inzia il ciclo di ricerca...
  do
    'cerca il prossimo inizio di tag
    pos1 = Instr(tmptxt, "<")
    'se non lo trovi esci dal ciclo di ricerca (non ci sono più tag da eliminare)
    if pos1=0 then
      exit do
    else
      'se lo trovi, cerca il simbolo di chiusura del tag
      pos2 = Instr(pos1, tmptxt, ">")
      'se non lo trovi esci dal ciclo di ricerca
      if pos2=0 then
        exit do
      else
        'elimina il tag determinato da pos1 e pos2
        tmptxt = Left(tmptxt, pos1-1)&Mid(tmptxt, pos2+1)
      end if
    end if
  loop
  'restituisci il testo "depurato" dai tag HTML
  RemoveTags = replace(tmptxt,"&nbsp;"," ")
  RemoveTags = replace(RemoveTags,"font"," ")
  RemoveTags=txt
end function

%> 


