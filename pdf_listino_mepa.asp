<!--#include virtual="/setup.asp" -->
<%
	
idsettore=request("idsettore")
if idsettore <>"" and not isnumeric(idsettore) then
	call determina_sqlinjection(idsettore)
	idsettore=""
	response.end
end if
idtag=request("idtag")
if idtag <>"" and not isnumeric(idtag) then
	call determina_sqlinjection(idtag)
	idtag=""
	response.end
end if
Server.ScriptTimeout=200

if idtag="" then
	response.end
end if

Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument
Set Image = Doc.OpenImage( Server.MapPath( Application("img_email")) )
Set TextParam = PDF.CreateParam
Set ImageParam = PDF.CreateParam
Set fs = Server.CreateObject("Scripting.filesystemObject")
'sql = "select * FROM categorie WHERE idcat= " & request("cat") &";"
'rs.open sql,conn
'nome_categoria=rs("categoria")
'rs.close
sql = "select * FROM impostazioni"
set rs= conn.execute (sql)
rs.movefirst
fondo_listino=rs("fondo_listino")
rs.close
Doc.Title = "Listino "&nomesito
Doc.Creator = nomesito
Bordo_sx=25
Set Page = Doc.Pages.Add
Set Param = Pdf.CreateParam
Set Param2 = Pdf.CreateParam
	'testata
	Set Param_testata = PDF.CreateParam
	Param_testata("x") =Bordo_sx' (Page.Width - Image.Width / 3) / 2
	Param_testata("y") = Page.Height - Image.Height/3 -20
	Param_testata("ScaleX") = 1 / 3
	Param_testata("ScaleY") = 1 / 3
	Set Image_testata = Doc.OpenImage( Server.MapPath( Application("img_email")) )
   
Set Table = Doc.CreateTable("width=562; height=600; Rows=1; Cols=4; Border=1; CellSpacing=-1; cellpadding=2 ")
Table.Font = Doc.Fonts("Arial")

Set HeaderRow = Table.Rows(1)
Param.Set("alignment=center")
With HeaderRow
   .BGColor = &HCCCCCC
   .Cells(1).AddText "Codice", "alignment=2; size=8"
   .Cells(3).AddText "Articolo", "alignment=2; size=8"
   '.Cells(4).AddText "Descrizione", Param
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
Set rs = Server.CreateObject("adodb.recordset")
	sql_base="select prodotti.*, settori.Nome_Settore, prodotti_tags.IdTag FROM (prodotti LEFT JOIN settori ON prodotti.idcat = settori.idsettore) INNER JOIN prodotti_tags ON prodotti.IDpro = prodotti_tags.idpro "
	if  idsettore>0 then
		where= where & " idcat=" & idsettore
	end if
	if  idtag>0 then
		where= where & " idtag=" & idtag
	end if

if idsettore>0  then
	ordine= " ORDER BY prodotti.articolo;"
else
	ordine= " ORDER BY settori.nome_settore"
end if
if where<>"" then where=" where flag=1 and "& where
Sql = sql_base&where&ordine
session(questofile)=""
'response.write sql
rs.Open sql, Conn,3,3
param.Set "expand=true; size=10" ' expand cell vertically
n=0
Do While Not rs.EOF
n=n+1
   Set Row = Table.Rows.Add(20) ' row height
   'codice
   'row.BGColor = &HFA85F8
	if not isnull(rs("codice")) then
	   Row.Cells(1).AddText rs("codice"), "alignment=0; size=10; expand=true;"
	end if
	'descrizione
   param.Add "size=8"
   param.Add "html=true"
   if len(rs("descrizione_listino"))>0 then
	   Row.Cells(3).AddText "<b>"&rs("articolo")&"</b><br>"&replace(rs("descrizione_listino"),vbcrlf,"<br>"), "alignment=0; size=8; html=true; expand=true;"
   else
	   Row.Cells(3).AddText "<b>"&rs("articolo")&"</b><br>", "alignment=0; size=8; html=true; expand=true;"
   end if
	'Immagine
	fileimg=rs("fileimg")
	if fileimg="nofoto" or isnull(fileimg) then
		txt="nofoto.gif"
   end if
   	if fileimg<>"" then
	
		if fs.FileExists(Server.MapPath( upl_img_cat & txt)) then
		Set ImageParam = PDF.CreateParam
			set Image = Doc.OpenImage( Server.MapPath( upl_img_cat & txt) )
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
		    'Row.Cells(3).AddText "larghezza immagine: "&image.width&" larghezza colonna: "&row.cells(2).width&" scala:"& scalex & "larghezza immagine: "&(image.width*scalex)&" scala_txt:"& scalex_txt&"ImageParam.scalex"&ImageParam("scalex"), "alignment=0; size=8; html=true; expand=true;"

			Row.Cells(2).canvas.drawimage image, ImageParam
			set ImageParam=nothing
			set image=nothing
	   end if
   end if
   if rs("prezzo")<0 then
	   Row.Cells(4).AddText rs("um")&vbcrlf&"Telefonare","alignment=1;size=10;expand=true;"
   else
	   Row.Cells(4).AddText rs("um")&vbcrlf&formatcurrency(rs("prezzo"),2,true) , "alignment=1;size=10;expand=true;"
	   if rs("fascia_di_sconto")>0 then
			Row.Cells(4).canvas.SetParams "FillColor =&H"&mid(colore_sconto(rs("fascia_di_sconto")),2)&""
			Row.Cells(4).canvas.FillRect 5,5,10,10
		end if
   end if
   'if row.height<image.height/2+10 then row.height=image.height/2+10 'corregge altezza riga in base all'immagine
   'session(questofile)=session(questofile)&"idpro:"&rs("idpro")&"n:"&n
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
	Page.Canvas.DrawText "Listino soli articoli MEPA "&nome_settore, "X=350; Y="&Page.Height-30&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "Generato il: "&date(), "X=350; Y="&Page.Height-42&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "I prezzi sono iva esclusa e possono variare senza preavviso secondo l'andamento del mercato", "X=25; Y="&Page.Height-55&";" , doc.fonts("Arial")
	Page.Canvas.DrawText "Pagina "&doc.pages.count, "X=20; Y=25;" , doc.fonts("Arial")
	Page.Canvas.DrawText fondo_listino,  "x=150; y=25; width=350; alignment=center; size=8;html=true", doc.fonts("Arial")
	Page.Canvas.DrawImage Image_testata, Param_testata

   if LastRow >= Table.Rows.Count Then Exit Do ' entire table displayed

   ' Display remaining part of table on the next page
   Set Page = Page.NextPage
   Param.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
   Param("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
   FirstRow = LastRow + 1
Loop
nome_categoria=Replace(nome_categoria,",", "")
nome_categoria=Replace(nome_categoria,".", "")
nome_categoria="Listino_mepa"
Doc.SaveHttp "attachment;filename="&nome_categoria&".pdf" 


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


