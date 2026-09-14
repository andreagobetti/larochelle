<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
Elimina tutti gli articoli senza settore<br>
<form action="<%=questofile%>" method="post">
	
<input type="submit" name="esegui" name="Esegui"/>
</form> 
<%
if request.form("esegui")<>"" then
	
	strSql="select prodotti.*, settori_prodotti.ID FROM prodotti LEFT JOIN settori_prodotti ON prodotti.IDpro = settori_prodotti.IDpro WHERE (((settori_prodotti.ID) Is Null))"

	set rs_articoli=conn.execute(strsql)
	
	
	do while not rs_articoli.eof 
	
	
	
		idpro=rs_articoli("idpro")
		sql="select ordini_dett.* from ordini_dett where idpro="&idpro
		set rs=conn.execute (sql)
		if rs.eof then in_ordini=false else in_ordini=true
		if in_ordini then str_log=str_log&"Presente in ordini: "
		do while not rs.eof
		str_log=str_log&rs("idord")& " "
		rs.movenext
		loop
		rs.close
		if in_ordini then
			str_log=str_log& "<br> "
		else
			str_log=str_log& "Presente in nessun ordine<br> "
		end if
			
		'elimina da tabella preferiti
		sql="delete from preferiti where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in preferiti<br>"
		
		'elimina da tabella carrello
		sql="delete from carrello where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in carrello<br>"
		
		'elimina da tabella magazzino
		sql="delete from magazzino where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in magazzino<br>"
		
		'elimina da tabella prodotti_consigliati
		sql="delete from prodotti_consigliati where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in articoli_consigliati<br>"
		
		'elimina da tabella settori_prodotti
		sql="delete from settori_prodotti where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in settori<br>"
		'elimina da tabella varianti_a
		sql="delete from varianti_a where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in varianti_a<br>"
		
		'elimina da tabella varianti_b
		sql="delete from varianti_b where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in varianti_b<br>"
		
		'elimina da tabella prodotti_brand
		sql="delete from prodotti_brand where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in prodotti_brand<br>"
		
		'elimina da tabella prodotti_dettagli
		sql="delete from prodotti_dettagli where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in prodotti_brand<br>"
		
		'elimina da tabella prodotti_gruppihome
		sql="delete from prodotti_gruppihome where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in prodotti_gruppihome<br>"
		
		'elimina da tabella prodotti_marca_modello
		sql="delete from prodotti_marca_modello where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in prodotti_marca_modello<br>"
		'elimina da tabella prodotti_tags
		sql="delete from prodotti_tags where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in prodotti_tags<br>"
		
		'elimina da tabella prodotti_mepa_campi
		sql="delete from prodotti_mepa_campi where idpro=" & idpro
		conn.execute sql,num
		str_log=str_log&"Eliminati "&num&" record in prodotti_mepa_campi<br>"
		
		
		'elimina da tabella files
		
		sql="select * from files where cosa=4 and idcosa="&idpro
		rs.Open sql, conn, 1, 3
		do while not rs.eof
			call elimina_file(rs("path")&rs("filename"))
			call elimina_file(rs("path")&"s/"&rs("filename"))
			
			
			rs.delete
			rs.movenext
		loop
		rs.close
		
		
		'elimina da tabella prodotti
		Set fsoMyFile = CreateObject("Scripting.filesystemObject")
		sql="select * from prodotti where idpro=" & idpro
		rs.Open sql, conn, 3, 3
		if not rs.eof then
			call elimina_file(upl_img_cat&idpro&"_b.jpg")
			call elimina_file(upl_img_cat&idpro&"_s.jpg")
			
			add2log "Eliminato [articolo="& idpro&"]"&rs("codice")& " idpro "&idpro&"[/articolo] eseguito da "& session("nominativo")&vbcrlf&"Articolo:"&rs("articolo")&"<br>"&str_log, 3
	
			rs.delete
		end if
		set rs = nothing
	
	
	
	 rs_articoli.MoveNext
	 loop	
	
	
	response.write "ESEGUITO"
end if
call connclose()
sub elimina_file(pathAndFilename)
	dim MapPathPathAndFilename
	Set fsoMyFile = CreateObject("Scripting.filesystemObject")
	MapPathPathAndFilename=Server.MapPath(pathAndFilename)
			if fsoMyFile.FileExists(MapPathPathAndFilename) then
				testo=testo&VbCrLf&"File eliminato"
				fsoMyFile.DeleteFile(MapPathPathAndFilename)
				str_log=str_log&"Eliminato file "&pathAndFilename&"<br>"
			else
				str_log=str_log&VbCrLf&"File "&pathAndFilename&" non trovato"
			end if
	Set fsoMyFile = Nothing
end sub

%>