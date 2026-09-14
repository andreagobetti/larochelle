<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CharSet = "UTF-8"%>
<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->
<%
id_tipo_allegato=request("id_tipo_allegato")
tipo_allegato=request("tipo_allegato")
if session("iduser")=1 then add2log "Sub_upload chiamato id_tipo_allegato:"&id_tipo_allegato&" tipo_allegato:"&tipo_allegato&"<br>"&vbcrlf&queryeform(),0
session("Json")=""

'1=fatture da fornitore
'2=Allegati articoli
'3=ordini
'4=Immagini articoli
'5=Manuali
'6=fatture cliente
'7=listini
'8=ebay
'9=Categorie mepa
'10=Articoli mepa
'11=Preventivi
'12=ordini fornitori

'Main Configiration of this page
maxWidth=800*2
maxHeight=1120*2

'savethumbsfolder="/public/files/aa/thumbnails/"
savethumbsfolder=""
thumbstype="fitin.square" ' or you can use "croped.square" 
thumbssize=80 'square edge
'Dont forget to give write permisions to uploading folder and database
session("test")=time()
'Connection String


'Internet Explorer and Opera do not yet support XMLHTTPRequest file uploads, so iframe based uploads require a Content-type of text/plain
strUserAgent = UCase(CStr(Request.ServerVariables("HTTP_USER_AGENT")))
if InStr(strUserAgent, "OPERA") or InStr(strUserAgent, "MSIE") or request("test")<>"" then
	response.ContentType="text/html"
else
	response.ContentType="application/json"
end if
Dim jsn : Set jsn = jsObject()
Set jsn("files") = jsArray()
if tipo_allegato="" or id_tipo_allegato="" then
	
	Set jsn("files")(Null) = jsObject()
	jsn("files")(Null)("error")="Tipo allegato o id mancanti"
	jsn("files")(Null)("name")=newname
	jsn("files")(Null)("size")=readfilesize
	jsn.Flush	
	set jsn = Nothing
	
	response.End
end if


'Persits ASPUpload
Set Upload = Server.CreateObject("Persits.Upload")

'Count = Upload.Save(Server.MapPath("/mdb-database/"))
Upload.Save '(Server.MapPath("/public/"))



'Set Upload = Server.CreateObject("Persits.Upload.1" )
'Upload.IgnoreNoPost = True ' It needs becouse, first loading the uploader page doesnt post anything to this page 
'Upload.CodePage = 65001 'Charset
'Upload.Overwritefiles = true 'It is not important, because always stytem generate a new image name. 
'Upload.SaveVirtual savefolder 'temprory save, can delete "virtual" but for goddady need virtual
'Upload.Save


'session("test")=session("test")&"Upload.save-"

'Database connection object
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "select SQL_CALC_FOUND_ROWS * FROM files where idcosa="&id_tipo_allegato&" and cosa="&tipo_allegato , conn, 3,2
n_file=clng(conn.Execute("Select Found_Rows();")(0).Value)
'session("test")=session("test")&"Qui4-"

'we must learn how many files uploaded (for opera browser)
totalfile=0
For Each File in Upload.files
	totalfile=totalfile+1
Next

i=1
For Each File in Upload.files

	'new file name
	select case tipo_allegato
	case "1" 	'Allegati fatture
		savefolder="/public/files/fatture/"
		inizio_nome="Fattura_"
		newname = inizio_nome&id_tipo_allegato&"_"&n_file&file.ext
	case "2"	'Allegati articoli
		savefolder="/public/files/allegati/"
		inizio_nome="All_"
		newname = inizio_nome&id_tipo_allegato&"_"&n_file&"_"&replace(file.filename," ","_")
	case "3"	'Allegati ordini
		savefolder="/public/files/allegati/"
		inizio_nome="All_O_"
		newname = inizio_nome&id_tipo_allegato&"_"&n_file&"_"&replace(file.filename," ","_")
		testo_add2log=" allegato a ordine idord:"&id_tipo_allegato&" "
	case "4" 	'Immagini articoli
		savefolder=upl_img_cat	'"/public/img_prod/"
		inizio_nome="IMG_"
		fine_nome=replace(time(),":","")
		newname = inizio_nome&id_tipo_allegato&"_"&n_file&"_"&fine_nome&file.ext
	case "5" 	'Manuali
		savefolder="/public/manuali/"
		inizio_nome="Manuale_"
		fine_nome=replace(time(),":","")
		newname = inizio_nome&replace(replace(file.filename," ","_"),file.ext,"")&"_"&fine_nome&file.ext
		testo_add2log=" manuale "&newname&" "
	case "6" 	'Manuali
		savefolder="/public/files/fatture/"
		inizio_nome="Fcli_"
		fine_nome=replace(time(),":","")
		newname = inizio_nome&replace(replace(file.filename," ","_"),file.ext,"")&"_"&fine_nome&file.ext
		testo_add2log=" fattura cliente "&newname&" "
	case "7" 	'Listini
		savefolder="/public/listini/"
		inizio_nome="Listino_"
		fine_nome=replace(time(),":","")
		newname = inizio_nome&replace(replace(file.filename," ","_"),file.ext,"")&"_"&fine_nome&file.ext
		testo_add2log=" listino "&newname&" "
	case "8" 	'ebay
		savefolder="/public/csv_ebay/"
		'inizio_nome=""
		'fine_nome=""
		newname=replace(file.filename,"-","_")
		newname = inizio_nome&replace(replace(newname," ","_"),file.ext,"")&"_"&fine_nome&file.ext
		testo_add2log=" csv "&newname&" "
	case "9" 	'Categorie mepa
		savefolder="/public/mepa/"
		'inizio_nome=""
		'fine_nome=""
		newname=left(replace(file.filename,"-","_"),15)
		newname = inizio_nome&replace(replace(newname," ","_"),file.ext,"")&"_"&fine_nome&file.ext
		testo_add2log=" csv "&newname&" "
	case "10" 	'Articoli mepa
		savefolder="/public/mepa_articoli/"
		'inizio_nome=""
		'fine_nome=""
		newname=left(replace(file.filename,"-","_"),15)
		newname = inizio_nome&replace(replace(newname," ","_"),file.ext,"")&"_"&fine_nome&file.ext
		testo_add2log=" csv "&newname&" "
	case "11" 	'Preventivi
		savefolder="/public/files/allegati/"
		inizio_nome="All_P_"
		newname = inizio_nome&id_tipo_allegato&"_"&n_file&"_"&replace(file.filename," ","_")
		testo_add2log=" allegato a preventivo idord:"&id_tipo_allegato&" "

	end select
	
	Set regEx = New RegExp

	regex.global = true
	regex.IgnoreCase = true
	regex.Pattern = "[^a-zA-Z0-9 _ .]"
	newname = regex.Replace(newname, "_")
	
	
	
	'Se manca creo la cartella
	set fsoobject = CreateObject("Scripting.filesystemObject" )
	if not fsoobject.FolderExists(Server.MapPath(savefolder)) then
		on error resume next
		fsoobject.CreateFolder(Server.MapPath(savefolder))
		if err.number<>0 then
			call add2log("Impossibile creare cartella: "&savefolder,-1)
			response.End
		end if
			
	end if

	
	on error resume next
	File.SaveAs Server.MapPath(savefolder & newname )
	if err.number<>0 then
			call add2log(Server.MapPath(savefolder & newname ),0)
			response.end
	end if



	if tipo_allegato="8" then
		'carico i dati csv
		Set csv = Server.CreateObject("ADODB.Connection")
		csv.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
       "Data Source=" & Server.MapPath("/public/csv_ebay/") & ";" & _ 
       "Extended Properties=""text;HDR=Yes;FMT=Delimited"""

		Set csvRS = csv.Execute("SELECT * FROM "&newname)
		txt=""
		if csvrs.eof then
			Set jsn("files")(Null) = jsObject()
			jsn("files")(Null)("error")="Il file CSV non contiene record"
			jsn("files")(Null)("name")=newname
			jsn("files")(Null)("size")=readfilesize
			jsn.Flush	
			set jsn = Nothing
			call connclose()
			response.End
		else
			on error resume Next
			code=csvrs("code")
			if err.number<>0 then
				Set jsn("files")(Null) = jsObject()
				jsn("files")(Null)("error")="Il file CSV non sembra un risultato di caricamento eBay"
				jsn("files")(Null)("name")=newname
				jsn("files")(Null)("size")=readfilesize
				jsn.Flush	
				set jsn = Nothing
				call connclose()
				response.End
				end if
			on error goto 0
		end if
		
		Do Until csvRS.EOF
			customlabel=split(csvrs("customlabel"),"-")
			idebay=customlabel(0)
			if txt<>"" then txt=txt&"| "
			txt=txt&idebay&":"&csvrs("status")
			'Verifico se è già una inserzione attiva
			endtime=conn.execute("select prodotti_ebay.endtime from prodotti_ebay where idebay="&idebay)(0)
			giorni=datediff("d",now(),endtime)
			if giorni<0 or not isnumeric(giorni) then
				'Inserzione scaduta
				txt=txt&" aggiorno"
				starttime=csvrs("starttime")
				if isnull(starttime) then
					starttime="NULL"
				else
					starttime="'"&starttime&"'"
				end if
				endtime=csvrs("endtime")
				if isnull(endtime) then
					endtime="NULL"
				else
					endtime="'"&endtime&"'"
				end if
				
				sql="update prodotti_ebay set itemid='"&csvrs("itemid")&"', starttime="&starttime&", endtime="&endtime&", status='"&lcase(csvrs("status"))&"', errormessage='errorcode:"&csvrs("errorcode")&"<br>"&csvrs("errormessage")&"' where idebay="&idebay
				conn.execute(sql)
				txt=txt&sql&"<br>"
			else
				'Inserzione attiva
								txt=txt&" attiva"		
			end if
		    csvRS.MoveNext
		Loop
		csvRS.Close
		csv.Close
		set csv = Nothing
		call add2log("Analizzato csv ebay "&txt,1)
	end if
	if tipo_allegato="9" then
		'Sostituisco ' - ' con ' '
		Const ForReading = 1
		Const ForWriting = 2
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		Set objFile = objFSO.OpenTextFile(Server.MapPath(savefolder&newname), ForReading)
		strText = objFile.ReadAll
		objFile.Close

		strNewText = Replace(strText, ", - ", ", ")

		Set objFile = objFSO.OpenTextFile(Server.MapPath(savefolder&newname), ForWriting)

		objFile.write( strNewText)

		objFile.Close
		
		set objFile = Nothing
		'''''conn.Execute("delete from mepa_categorie")
		'''''conn.Execute("delete from mepa_campi")
		
		
		
		Set csvconn = Server.CreateObject("ADODB.Connection")
		csvconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
       "Data Source=" & Server.MapPath(savefolder) & ";" & _ 
       "Extended Properties=""text;HDR=Yes;FMT=Delimited"""
		session(questofile)=""
		on error resume next
		Set csvRS = csvconn.Execute("SELECT * FROM "&newname)
		if err.number<>0 then
			call add2log ("SELECT * FROM "&newname,0)
			response.End
		end if	
			
		txt=""
		
		do while not csvrs.EOF
			'Aggiungo a categorie Mepa
			Set rsmepa = Server.CreateObject("ADODB.Recordset")
			sql="select mepa_categorie.* from mepa_categorie where descrizione='"&csvrs(0)&"'"
			
			rsmepa.Open  sql, conn, 3, 3
			
			if rsmepa.eof then
				rsmepa.addnew
				rsmepa("descrizione")=csvrs(0)
				rsmepa.update
				id=Get_last_id("mepa_categorie")
			else
				id=rsmepa("id")
			end if
			rsmepa.close
			
			'Aggiungo a mepa_campi
			on error resume next
			label=replace(csvrs(1),"'","''")
			sql="select mepa_campi.* from mepa_campi where idcategoria="&id&" and label='"&label&"'"
			rsmepa.Open  sql, conn, 3, 3

			if err.number<>0 then
				
				call add2log("Errore in query sql:"&sql,0)
				response.end
			end if
			
			if rsmepa.eof then
				rsmepa.addnew
				rsmepa("idcategoria")=id
			end if
			rsmepa("label")=csvrs(1)
			rsmepa("obbligatorio")=csvrs(2)
			rsmepa("digitabile")=csvrs(3)
			rsmepa("formato")=csvrs(4)
			rsmepa("dimensione")=csvrs(5)
			rsmepa("valori_default")=csvrs(6)
			rsmepa("range")=csvrs(7)
			rsmepa("algoritmo")=csvrs(8)
			rsmepa.update
			
			rsmepa.close
			txt=txt& "Aggiunto:"& csvrs(1)&" "&get_last_id("mepa_campi")&"<br>"
			
			
		csvrs.MoveNext
		loop
		call add2log(txt,0)
		csvrs.movefirst
		JSONstring= QueryToJSON(csvRS).jsString
		sql="Update mepa_categorie set campi_JSON='"&replace(JSONstring,"'","''")&"' where id="&id
		conn.execute (sql)
		set csvrs = Nothing
		set rsmepa = Nothing
	end if

	if tipo_allegato="10" then
		idcategoria=request.querystring("idcategoria")
		response.redirect "importa_articoli_mepa.php?idcategoria="&idcategoria&"&filename="&newname
		
	end if

	'control main config of the top of page
	if savethumbsfolder<>"" then
		if thumbstype="croped.square" then
'session("test")=session("test")&"Qui7A-"
			'create thumb of imgages with crop acording to squre size	
			Set JpegCropedSquare = Server.CreateObject("Persits.Jpeg")
			JpegCropedSquare.Open File.Path
			if JpegCropedSquare.OriginalHeight<JpegCropedSquare.OriginalWidth then
				JpegCropedSquare.Height = thumbssize
				JpegCropedSquare.Width = (JpegCropedSquare.OriginalWidth * thumbssize) / JpegCropedSquare.OriginalHeight
				point1=(JpegCropedSquare.Width-thumbssize)/2
				JpegCropedSquare.Crop point1, 0,(point1+thumbssize), thumbssize
				JpegCropedSquare.Save Server.Mappath(savethumbsfolder) & "\" & newname & "." & File.ImageType 
			else
				JpegCropedSquare.Width = thumbssize
				JpegCropedSquare.Height = (JpegCropedSquare.OriginalHeight * thumbssize) / JpegCropedSquare.OriginalWidth
				point2=(JpegCropedSquare.Height-thumbssize)/2
				JpegCropedSquare.Crop 0, point2, thumbssize,(point2+thumbssize)
				JpegCropedSquare.Save Server.Mappath(savethumbsfolder) & "\" & newname & "." & File.ImageType 
			end if
		else
			'create thumb of imgages with fit in square
			Set JpegFitSquare = Server.CreateObject("Persits.Jpeg")
			JpegFitSquare.Open File.Path
			if JpegFitSquare.OriginalHeight>JpegFitSquare.OriginalWidth then
				JpegFitSquare.Height = thumbssize
				JpegFitSquare.Width = (JpegFitSquare.OriginalWidth * thumbssize) / JpegFitSquare.OriginalHeight
				JpegFitSquare.Save Server.Mappath(savethumbsfolder) & "\" & newname & "." & File.ImageType 
			else
				JpegFitSquare.Width = thumbssize
				JpegFitSquare.Height = (JpegFitSquare.OriginalHeight * thumbssize) / JpegFitSquare.OriginalWidth
				JpegFitSquare.Save Server.Mappath(savethumbsfolder) & "\" & newname & "." & File.ImageType 
			end if
		end if	
	end if

	
	'set fsoget=fsoobject.GetFile(Server.MapPath(savefolder & newname & "." & File.ImageType))
	set fsoget=fsoobject.GetFile(Server.MapPath(savefolder & newname ))
	readfilesize=fsoget.size	
	
	txt_errore=""
	testo_log=""
	if tipo_allegato="4" then
		'Copio l'immagine in backup
		savefolderbk=savefolder&"bk/"
		
		
		'Se manca creo la cartella
		set fsoobject = CreateObject("Scripting.filesystemObject" )
		if not fsoobject.FolderExists(Server.MapPath(savefolderbk)) then
			fsoobject.CreateFolder(Server.MapPath(savefolderbk))
		end if




		
		
		'IMMAGINE L
			testo_log=testo_log&"<br><b>Immagine L</b>"
			Set Jpeg = Server.CreateObject("Persits.Jpeg")
			Jpeg.Open File.Path
			if Jpeg.OriginalWidth <800 or Jpeg.OriginalHeight<1120 then
				txt_errore="Pixel immagine troppo bassi: "&Jpeg.OriginalHeight&"x"&Jpeg.OriginalWidth
			end if	
			AspectRatio=jpeg.OriginalWidth/jpeg.OriginalHeight
			testo_log=testo_log&"Dimensioni immagine:"&jpeg.OriginalWidth&"x"&jpeg.OriginalHeight&":"&roundup(AspectRatio,2)
			if AspectRatio>0.73 then	'Altezza troppo bassa, immagine larga
				'Calcolo altezza ottimale
				optimalHeight=jpeg.OriginalWidth/0.7142
				
				testo_log=testo_log&"<br>Aumento altezza a:"&int(optimalHeight)
	
				'Calcolo delta
				deltaHeight=int((optimalheight-jpeg.OriginalHeight)/2)
	
				'Croppo in negativo solo in Y
				jpeg.Canvas.Brush.Color = &HFFFFFF
				jpeg.Crop 0, deltaHeight*-1, jpeg.OriginalWidth , jpeg.originalHeight + deltaHeight
				modificato=true
			elseif AspectRatio<0.70 then	'Larghezza troppo bassa, immagine alta
				'Calcolo larghezza ottimale
				optimalWidth=jpeg.OriginalHeight*0.7142
				testo_log=testo_log&"<br>Aumento larghezza a:"&int(optimalWidth)
				
				'Calcolo delta
				deltaWidth=int((optimalWidth-jpeg.OriginalWidth)/2)
	
				'Croppo in negativo solo in X
				jpeg.Canvas.Brush.Color = &HFFFFFF
				jpeg.Crop deltaWidth*-1, 0, jpeg.OriginalWidth+deltaWidth , jpeg.OriginalHeight
				modificato=true
			else
				testo_log=testo_log&"<br>Ratio corretto"
			end if
			jpeg.PreserveAspectRatio  = true

			fattore_riduzione_w=jpeg.width/maxWidth
			fattore_riduzione_h=jpeg.height/maxHeight
			testo_log=testo_log&"<br>Fattori di riduzione: w"&fattore_riduzione_w&" h"&fattore_riduzione_h
			if fattore_riduzione_w<=1 and fattore_riduzione_h<=1 then
				'nessuna riduzione
				testo_log=testo_log&"<br>Nessuna riduzione richiesta"
			elseif fattore_riduzione_w>=fattore_riduzione_h then
				jpeg.Width= maxWidth
				'Image.Height= image.OriginalHeight/fattore_riduzione_w
				testo_log=testo_log&"<br>Regolato larghezza w="&jpeg.Width&" h="&jpeg.Height
				modificato=true
			else
				jpeg.Height= maxHeight
				'Image.Width=image.OriginalWidth/fattore_riduzione_h
				testo_log=testo_log&"<br>Regolato altezza w="&jpeg.Width&" h="&jpeg.Height
				modificato=true
			end if
			if modificato then
				testo_log=vbcrlf&testo_log&"<br>Dimensione finale:"&jpeg.width&"x"&jpeg.height&":"&roundup(jpeg.width/jpeg.height,2)
				txt_chk=chk_immagine(jpeg.width,jpeg.height)
				Jpeg.Save Server.MapPath(savefolder&newname)
			end if
			
			'Dopo aver manipolato l'immagine grande effettuo la copia
			fsoobject.CopyFile Server.MapPath(savefolder&newname),Server.MapPath(savefolderbk&newname),true

			if Application("filigrana")="1" then
				'Applico la filigrana su immagine grande
				rapporto=Application("width_filigrana")/18

				' Font path
				FontPath = Jpeg.WindowsDirectory & "\fonts\courbd.ttf"
				y=Jpeg.Height/3*2
				x=Jpeg.width
				'Calcolato in base al margine
				'font_size=(x-(margine*2))/rapporto
				
				'Calcolato in base all'ingombro in larghezza
				font_size=(x/2)/rapporto
				margine=(x/2/2)
				
				
				Jpeg.Canvas.Font.Color = &HFAC83C 
				jpeg.Canvas.Font.Size = font_size
				jpeg.Canvas.Font.Opacity = 0.5
				txt=txt&vbcrlf&larg
				on error resume next
				jpeg.Canvas.PrintTextEx Application("testo_filigrana"), margine, y, "c:\Windows\Fonts\Arial.ttf"
				if err.number>0 then
					testo_log=testo_log&vbcrlf&"<b>Filigrana NON applicata, immagine non RGB</b>"
					txt_errore=txt_errore&"<b>Filigrana NON applicata, immagine non RGB</b>"
				else
					Jpeg.Save Server.MapPath(savefolder&newname)
					testo_log=testo_log&vbcrlf&"Applicato filigrana"
				end if
				
				
				
			else
				testo_log=testo_log&"Filigrana non applicata"
			end if
			
			
			
			
		'IMMAGINE S
			savefolders=savefolder&"s/"
			Jpeg.Open Server.MapPath(savefolderbk&newname)
			
			maxWidth=228
			maxHeight=319
			testo_log=testo_log&"<br><b>Immagine S</b>"
			AspectRatio=jpeg.OriginalWidth/jpeg.OriginalHeight
			testo_log=testo_log&"<br>Dimensioni immagine:"&jpeg.OriginalWidth&"x"&jpeg.OriginalHeight&":"&roundup(AspectRatio,2)
			if AspectRatio>0.73 then	'Altezza troppo bassa, immagine larga
				'Calcolo altezza ottimale
				optimalHeight=jpeg.OriginalWidth/0.7142
				
				testo_log=testo_log&"<br>Aumento altezza a:"&int(optimalHeight)
	
				'Calcolo delta
				deltaHeight=int((optimalheight-jpeg.OriginalHeight)/2)
	
				'Croppo in negativo solo in Y
				jpeg.Canvas.Brush.Color = &HFFFFFF
				jpeg.Crop 0, deltaHeight*-1, jpeg.OriginalWidth , jpeg.originalHeight + deltaHeight
				modificato=true
			elseif AspectRatio<0.70 then	'Larghezza troppo bassa, immagine alta
				'Calcolo larghezza ottimale
				optimalWidth=jpeg.OriginalHeight*0.7142
				testo_log=testo_log&"<br>Aumento larghezza a:"&int(optimalWidth)
				
				'Calcolo delta
				deltaWidth=int((optimalWidth-jpeg.OriginalWidth)/2)
	
				'Croppo in negativo solo in X
				jpeg.Canvas.Brush.Color = &HFFFFFF
				jpeg.Crop deltaWidth*-1, 0, jpeg.OriginalWidth+deltaWidth , jpeg.OriginalHeight
				modificato=true
			else
				testo_log=testo_log&"<br>Immagine conforme"
			end if
			
			jpeg.PreserveAspectRatio  = true

			fattore_riduzione_w=jpeg.width/maxWidth
			fattore_riduzione_h=jpeg.height/maxHeight
			testo_log=testo_log&" Fattori di riduzione: w"&fattore_riduzione_w&" h"&fattore_riduzione_h
			if fattore_riduzione_w<=1 and fattore_riduzione_h<=1 then
				'nessuna riduzione
				testo_log=testo_log&" Nessuna riduzione"
			elseif fattore_riduzione_w>=fattore_riduzione_h then
				jpeg.Width= maxWidth
				'Image.Height= image.OriginalHeight/fattore_riduzione_w
				testo_log=testo_log&" Regolato larghezza w="&jpeg.Width&" h="&jpeg.Height
				modificato=true
			else
				jpeg.Height= maxHeight
				'Image.Width=image.OriginalWidth/fattore_riduzione_h
				testo_log=testo_log&" Regolato altezza w="&jpeg.Width&" h="&jpeg.Height
				modificato=true
			end if
			testo_log=testo_log&"<br>Dimensione finale:"&jpeg.width&"x"&jpeg.height&":"&roundup(jpeg.width/jpeg.height,2)
			Jpeg.Quality=80
			Jpeg.Save Server.MapPath(savefolders&newname)

		
		
		
		
		
		
		
		
		
		
		
	end if
	
	set fs=Server.CreateObject("Scripting.filesystemObject")
	set f=fs.GetFile(Server.MapPath(savefolder&newname))
	fsize=f.Size
	set f=Nothing
	set fs=Nothing
								
	'if tipo_allegato<>"8" then					
														
	'add database to new image	
	rs.addnew
	picid=rs("idfiles")	'for deleting	we use id					 
	rs("uploaddate")=nowtime
	rs("filesize")=fsize
	rs("filename")=newname
	rs("idcosa")=id_tipo_allegato
	rs("cosa")=tipo_allegato
	rs("uploaddate")=now()
	rs("path")=savefolder
	rs("ordine")=n_file*2
	rs("chk")=txt_chk
	if tipo_allegato=5 then rs("descrizione")=newname
	rs.update
	idfile=Get_last_id("files")
	'end if
	if tipo_allegato="4" and n_file<2 then
			'Imposto immagine prodotto predefinita 
			set rs_pro=Server.CreateObject("ADODB.Recordset")
			rs_pro.Open "select prodotti.* from prodotti where idpro="&id_tipo_allegato,conn,1,3
			rs_pro("fileimg")=newname
			rs_pro.update
			rs_pro.close
			set rs_pro=nothing
	end if			
	if tipo_allegato="4" then
			set rs_pro=conn.execute ("select prodotti.* from prodotti where idpro="&id_tipo_allegato)
			txt=" immagine [articolo="& rs_pro("idpro")&"]"&rs_pro("codice")&"[/articolo] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]"&testo_log
			set rs_pro=nothing
	else
		txt="completato"
	end if			
						 
	Set jsn("files")(Null) = jsObject()
	if tipo_allegato="4" and txt_errore="" then
		newname=" "&newname&", immagine conforme"
		txt=txt&newname
	end if
	jsn("files")(Null)("name")=newname
	jsn("files")(Null)("size")=readfilesize
	jsn("files")(Null)("url")=savefolder &  newname 
	jsn("files")(Null)("thumbnail_url")=savethumbsfolder & "/" & newname & "." & File.ImageType
	jsn("files")(Null)("delete_url")="sub_upload.asp?islem=delete&id="&picid
	jsn("files")(Null)("delete_type")="POST"
	jsn("files")(Null)("idfile")=idfile
	if txt_errore<>"" then
	 	jsn("files")(Null)("error")=txt_errore
	 	txt_errore="<br>"&txt_errore
	 end if

									
	i=i+1	
	'remove the temp file which is the orjinal file
	'File.Delete
	
Next
session("Json")=time()&" "&jsn.jsString
response.write jsn.jsString	
add2log "Upload "&testo_add2log&txt&txt_errore&vbcrlf&"From:"&request.querystring("from"),2

set jsn=Nothing
rs.Close
set rs=Nothing
call ResponseEnd()

response.End
'For delete pictures small code
'dont forget the control who accessing of this page or anybody can delete any picture of your system
islem=request.QueryString("islem")
if islem="delete" then
	imagefileid=request.QueryString("id")
	if not isnumeric(imagefileid) then
		response.End()
	end if
	'Database connection to learn filename from id
	SQLStrFileName = "select * FROM files WHERE idfiles="&cint(imagefileid)
	set ConnectionObjFileName = Server.CreateObject("ADODB.Connection")
	ConnectionObjFileName.Open ConnectionStr
	set objRsFileName = Server.CreateObject("ADODB.Recordset")
	objRsFileName.CursorType = 2
	objRsFileName.CursorLocation = 2
	objRsFileName.LockType = 3
	objRsFileName.Open SQLStrFileName, ConnectionObjFileName, , , &H0001
		 if objRsFileName.eof then
			 response.End()
		 End if
	filename=objRsFileName("filename")
	'delete the file
 	Set fso = CreateObject("Scripting.filesystemObject" )  
	if fso.FileExists (Server.MapPath(savefolder & filename)) Then
		fso.DeleteFile(Server.MapPath(savefolder & filename)) 
		'delete from database
		delsql="delete FROM pictures WHERE pictureid="&cint(imagefileid)
		set objConnDel = Server.CreateObject("ADODB.Connection")
		objConnDel.Open ConnectionStr
		objConnDel.Execute delsql
		objConnDel.Close
		SET objConnDel = Nothing
		'delete thumb image
		Set fsothumb = CreateObject("Scripting.filesystemObject" )  
		if fsothumb.FileExists (Server.MapPath(savethumbsfolder & filename)) Then
			fsothumb.DeleteFile(Server.MapPath(savethumbsfolder & filename)) 
		end if
	end if
response.End()
end if

'Create new filenames function
Function newfilename( nNoChars, sValidChars )
	Const szDefault = "abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
	Dim nCount
	Dim sRet
	Dim nNumber
	Dim nLength
	Randomize
	If sValidChars = "" Then
		sValidChars = szDefault		
	End If
	nLength = Len( sValidChars )
	For nCount = 1 To nNoChars
		nNumber = Int((nLength * Rnd) + 1)
		sRet = sRet & Mid( sValidChars, nNumber, 1 )
	Next
	newfilename = sRet	
End Function
'usage ->   newname = "YourSite.com_"+Now()+"_" + newfilename(5,"")

%>