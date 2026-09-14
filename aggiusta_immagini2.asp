<!--#include virtual="/setup.asp" -->

<%
response.buffer=false
Server.ScriptTimeout=60*30




'Creazione miniatura 228x319 (S)
savefolder=upl_img_cat&"s/"
set fsoobject = CreateObject("Scripting.filesystemObject" )
if not fsoobject.FolderExists(Server.MapPath(savefolder)) then
	fsoobject.CreateFolder(Server.MapPath(savefolder))
end if



		maxWidth=228
		maxHeight=319
	
	'on error resume next
		dim testo_log
		modificato=false
		sql="select * from files where cosa=4" 
		Set rs_files = Server.CreateObject("ADODB.Recordset")
		rs_files.Open sql , conn, 3,2
		max=rs_files.recordcount
		
		Set Jpeg = Server.CreateObject("Persits.Jpeg")
		do while  not rs_files.eof 
			testo_log=""
			idpro=rs_files("idcosa")
			fileName=rs_files("filename")
			openpath=rs_files("path")
			' Open source image
			imgurl=openpath&rs_files("filename")
			'on error resume Next

			Jpeg.Open Server.MapPath(imgurl)
			
			if err.number=0 then

	
				AspectRatio=jpeg.OriginalWidth/jpeg.OriginalHeight
				testo_log="Dimensioni immagine:"&jpeg.OriginalWidth&"x"&jpeg.OriginalHeight&":"&roundup(AspectRatio,2)
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
				Jpeg.Save Server.MapPath(savefolder&fileName)
				conn.execute ("update prodotti set fileimg='"&filename&"' where idpro="&idpro)
			else
				add2log "immagine mancante:"& openpath&filename&" idfiles:"&rs_files("idfiles"),0
			end if
			rs_files.MoveNext
			response.write testo_log

		loop


%>
