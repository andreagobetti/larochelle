<!--#include virtual="/setup.asp" -->
<%
	
if session("iduser")="" then
	call connclose()
	response.End
end if	
		
	
Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument
Set TextParam = PDF.CreateParam
Set ImageParam = PDF.CreateParam
Set fs = Server.CreateObject("Scripting.filesystemObject")





set rs=conn.execute("select * FROM impostazioni")
rs.movefirst
fondo_listino=rs("fondo_listino")
rs.close
Doc.Title = "Elenco prodotti "&nomesito
Doc.Creator = nomesito
Bordo_sx=25
Set Page = Doc.Pages.Add
Set Param = Pdf.CreateParam
Set Param2 = Pdf.CreateParam


		const file_logo="/public/files/logo_documenti.jpg"
		Set fs = Server.CreateObject("Scripting.filesystemObject")
		if fs.FileExists(Server.MapPath( file_logo)) then
			esiste_logo=true
			
			'Parte nuova
			dimensione_logo=90
			margine=3
			Set logo = Doc.OpenImage( Server.MapPath( file_logo) )
			'calcolo misure dell'immagine considerando la risoluzione DPI
			larghezza=logo.width*72/logo.resolutionx
			altezza=logo.height*72/logo.resolutiony
			
			
			scala=(dimensione_logo)/larghezza
			
			'calcolo nuova altezza considerando la scala per centrare poi l'immagine
			nuova_altezza=altezza*scala
			nuova_larghezza=larghezza*scala
			
			'Centra nella prima mezza pagina
			'param_x =((Page.Width/ 2)-nuova_larghezza)/2
			
			'Allinea a sx
			param_x=25
			
			'Sotto laschio metà del margine
			param_y = Page.Height - nuova_altezza - 10
			
			
			param_ScaleX = scala
			param_ScaleY = scala
			parametri="x="&param_x&"; y="&param_y&"; ScaleX="&param_ScaleX&"; ScaleY="&param_ScaleY&";"
			parametri=replace(parametri,",",".")
			
		else
			param_Y=Page.Height-20
		
		end if
		Set fs = Nothing

		'response.write parametri
   
Set Table = Doc.CreateTable("width=562; height=600; Rows=1; Cols=4; Border=1; CellSpacing=-1; cellpadding=2 ")
Table.Font = Doc.Fonts("Arial")

Set HeaderRow = Table.Rows(1)
Param.Set("alignment=center")
With HeaderRow
   .BGColor = &HCCCCCC
   .Cells(1).AddText "Codice", "alignment=2; size=8"
   .Cells(3).AddText "Articolo", "alignment=2; size=8"
   .Cells(4).AddText "U.m. / Prezzo", "alignment=2; size=8"
   .height=20
End With

' Set column widths
With Table.Rows(1)
   .Cells(1).Width = 60
   .Cells(2).Width = 80
   .Cells(3).Width = 360
   '.Cells(4).Width = 200
   .Cells(4).Width = 62
End With

' Populate table with data
sql="select prodotti.*, settori.nome_settore from prodotti inner join settori on prodotti.idcat = settori.idsettore"
sql=sql&" where visibilita=0 "
sql=sql&" order by nome_settore, articolo "

'response.write sql
set rs=conn.execute(sql)

param.Set "expand=true; size=10" ' expand cell vertically
n=0
settore_tmp=""
Do While Not rs.EOF




if rs("nome_settore")<>settore_tmp then
	Set Row = Table.Rows.Add(8) ' row height
	'Row.BGColor = &He6e6e6
	Row.BGColor = &HCCCCCC
	Row.height=20
	Row.Cells(1).AddText "Categoria: "&rs("nome_settore"), "alignment=0; size=10;"
	Row.Cells(1).colspan=4
	settore_tmp=rs("nome_settore")
end if 
n=n+1
   Set Row = Table.Rows.Add(80) ' row height
   'codice
   Row.Cells(1).AddText rs("codice"), "alignment=0; size=10; expand=true;"
	'Immagine
  fileimg=rs("fileimg")
	if fileimg="nofoto" or isnull(fileimg) then
		txt="nofoto.gif"
	elseif fileimg<>"" then
			txt=fileimg
   end if
   	if fileimg<>"" then
		Set fs = Server.CreateObject("Scripting.filesystemObject")
		if fs.FileExists(Server.MapPath( upl_img_cat & txt)) then
			
			Set ImageParam = PDF.CreateParam
			set Image = Doc.OpenImage( Server.MapPath( upl_img_cat & txt) )
			
			'Dimensioni del posto
			larghezza_posto=row.cells(2).width
			altezza_posto=row.height
			if image.resolutionx>0 and image.resolutiony>0 then
				'calcolo misure dell'immagine considerando la risoluzione DPI
				larghezza=image.width*72/image.resolutionx
				altezza=image.height*72/image.resolutiony
				
				margine=5
				
				scalay=(altezza_posto-margine)/altezza
				scalax=(larghezza_posto-margine)/larghezza
				
				if scalax<=scalay then scala=scalax else scala=scalay
				
				'calcolo nuova larghezza considerando la scala per centrare poi l'immagine
				nuova_larghezza=int(larghezza*scala)
				nuova_altezza=int(altezza*scala)
				
				'Trovo la x per centrare l'immagine
				pos_x=(larghezza_posto - nuova_larghezza ) / 2
				ImageParam("x") =  pos_x
				
				'Sotto laschio metà del margine
				pos_y=(altezza_posto - nuova_altezza ) / 2
				ImageParam("y") =  pos_y

				ImageParam("ScaleX") = scala
				ImageParam("ScaleY") = scala
			    'Row.Cells(3).AddText "larghezza immagine: "&image.width&" larghezza colonna: "&row.cells(2).width&" scala:"& scalex & "larghezza immagine: "&(image.width*scalex)&" scala_txt:"& scalex_txt&"ImageParam.scalex"&ImageParam("scalex"), "alignment=0; size=8; html=true; expand=true;"

				Row.Cells(2).canvas.drawimage image, ImageParam
				end if
				set ImageParam=nothing
			set image=nothing
			if session("iduser")=1 then conversione_immagine="larghezza cella:"&larghezza_posto&" altezza cella:"&altezza_posto&"larghezza:"&larghezza&" altezza:"&altezza& " scalax:"&scalax&" scalay:"&scalay&" scala:"&scala&" nuova larghezza:"&nuova_larghezza&" nuova altezza:"&nuova_altezza&" X:"&pos_x&" Y:"&pos_y

		else
			session("pdf_listino")=session("pdf_listino")&"<br>Immagine idpro:"&rs("codice")&" non trovata ("&Server.MapPath( upl_img_cat & txt)&")"
	   end if
   end if
	'descrizione
	param.Add "size=8"
	param.Add "html=true"
   
   'if len(rs("descrizione_listino"))>0 then
'	   Row.Cells(3).AddText conversione_immagine&"<b>"&rs("articolo")&"</b><br>"&replace(rs("descrizione_listino"),vbcrlf,"<br>"), "alignment=0; size=8; html=true; expand=true;"
'   else
	Row.Cells(3).AddText "<b>"&rs("articolo")&"</b><br>", "alignment=0; size=8; html=true; expand=true;"
'   end if



   if utente_admin or (session("vedi_prezzi")=1 and (cint(rs("prezzo_visibile")=2) or cint(rs("prezzo_visibile")=0)   )) then
	prezzo_minimo= converti_typevar14(rs("prezzo_ve_min"))
	'articolo=server.HTMLEncode(articolo)
		if not isnull(prezzo_minimo) and prezzo_minimo>0 then 
			Row.Cells(4).AddText rs("um")&vbcrlf&"Da "&formatcurrency(prezzo_minimo,2,true), "alignment=1;size=10;expand=true;"
		else
		   Row.Cells(4).AddText rs("um")&vbcrlf&formatcurrency(rs("prezzo"),2,true) , "alignment=1;size=10;expand=true;"
		end if
		if converti_typevar14(rs("fascia_di_sconto"))>0  then
			Row.Cells(4).canvas.SetParams "FillColor =&H"&mid(colore_sconto(rs("fascia_di_sconto")),2)&""
			Row.Cells(4).canvas.FillRect 5,5,10,10
		end if
	else
	   Row.Cells(4).AddText rs("um")&vbcrlf&"Telefonare","alignment=1;size=10;expand=true;"
   end if
   'if row.height<image.height/2+10 then row.height=image.height/2+10 'corregge altezza riga in base all'immagine
   rs.MoveNext
   set image= nothing
Loop

' Render table on document
'Set Page = Doc.Pages.Add
Param.Clear 
Param("x") = Bordo_sx'(Page.Width - Table.Width) / 2 ' center table on page
Param("y") = Page.Height - 70
Param("MaxHeight") = 700

FirstRow = 2
Do While True
	LastRow = Page.Canvas.DrawTable( Table, Param )
	' Pagina
	Page.Canvas.DrawText ucase(nomesito), "X=350; Y="&Page.Height-18&";" , doc.fonts("Arial")
	if mepa then nome_settore="MEPA"
	if utente_admin then testo="ADMIN (prezzi visibili)"
	Page.Canvas.DrawText "Elenco prodotti "&testo, "X=350; Y="&Page.Height-30&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "Generato il: "&date(), "X=350; Y="&Page.Height-42&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "I prezzi sono iva esclusa e possono variare senza preavviso secondo l'andamento del mercato", "X=25; Y="&Page.Height-55&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "Pagina "&doc.pages.count, "X=20; Y=25;" , doc.fonts("Arial")
	Page.Canvas.DrawText fondo_listino,  "x=150; y=25; width=350; alignment=center; size=8;html=true", doc.fonts("Arial")
	if esiste_logo then Page.Canvas.DrawImage logo, parametri

   if LastRow >= Table.Rows.Count Then Exit Do ' entire table displayed

   ' Display remaining part of table on the next page
   Set Page = Page.NextPage
   Param.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
   Param("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
   FirstRow = LastRow + 1
Loop
nome_settore=Replace(nome_settore,",", "")
nome_settore=Replace(nome_settore,".", "")
nome_documento="Elenco_prodotti_completo_"&date()
nome_documento=replace(nome_documento,"/","_")
Doc.SaveHttp "attachment;filename="&nome_documento&".pdf" 


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


