<%

function box_rs_articoli(rs, colonne)
	dim testo_sconto, stringa, n_articoli, idpro
	if iRecordsShown="" then
		iRecordsShown=0
		iPageSize=1000
	end if
	'if utente_andrea then box_rs_articoli="[soloio]iRecordsShown:"&iRecordsShown&"iPageSize:"&iPageSize&"rs.EOF:"&rs.EOF&"IsObject(rs):"&IsObject(rs)&"rs is Nothing:"&(rs is Nothing)&"[/soloio]"
	Do While iRecordsShown < iPageSize And not rs.EOF
		stringa=""

		novita=cbool(rs("novita"))
		prezzo=cdbl(rs("prezzo"))
		'if idsettore="" then idsettore=clng(rs("idcat"))
		'idsettore=clng(rs("idcat"))
		prezzo_minimo= converti_typevar14(rs("prezzo_ve_min"))
		promozione=rs("promozione")
		sconto=rs("sconto")
		idpro=rs("idpro")
		prezzo_visibile=rs("prezzo_visibile")
		if colonne=4 then
			stringa=stringa&"<div class=""col-md-3 col-sm-6 col-xs-12"">"
		elseif colonne=3 then
			stringa=stringa&"<div class=""col-md-4 col-sm-6 col-xs-12"" id=""id_"&rs("idpro")&""">"
		end if
		
		
		
		
		if prezzo>0 and Application("prezzi_con_iva")>0 then
			prezzo=prezzo*(1+Application("prezzi_con_iva")/100)
		end if
		if idsettore>0 then href_settore="&idsettore="&idsettore
		if Application("prezzi_con_iva")>0 then
			prezzo_minimo=prezzo_minimo*(1+Application("prezzi_con_iva")/100)
		end if
		'articolo=server.HTMLEncode(articolo)
		if not isnull(prezzo_minimo) and prezzo_minimo>0 then 
			if Application("prezzi_con_iva")>0 then
				prezzo_minimo=prezzo_minimo*(1+Application("prezzi_con_iva")/100)
			end if
			prezzoar="Da "&simbolo_valuta&formatnumber(prezzo_minimo,2)
			
		elseif promozione=1  and sconto>0 then
		    prezzoold=simbolo_valuta&formatnumber(prezzo,2)
		    prezzo=prezzo*(100-sconto)/100
		    prezzoar=simbolo_valuta&FormatNumber(prezzo,2)
			testo_sconto=sconto&"%"
		
		else
		    prezzoar=simbolo_valuta&FormatNumber(prezzo,2)
		end if
		
		fileimg=upl_img_cat&"s/"&rs("fileimg")
		'pos=instrrev(fileimg,"/")
		'newfileimg=left(fileimg,pos)&"s/"&mid(fileimg,pos+1)
		
		'response.write stringa&vbcrlf&stringanuova
		
		
		stringa=stringa& 	"<div class=""item item-hover"" id=""idpro_"&rs("idpro")&""" itemscope itemtype=""http://schema.org/Product"">"&_
					"<div class=""item-image-container"" style=""height:330px;"">" &_
					"<figure>"&_
					"<a href=""product.asp?idpro="&rs("idpro")&href_settore&""">"&_
					"<img itemprop=""image"" src="""&fileimg&""" alt="""&rs("articolo")&""" class=""item-image"">"&_
					"<img src="""&fileimg&""" alt="""&rs("articolo")&""" class=""item-image-hover""></a></figure>"
		if rs("flag")=1 then
			stringa=stringa &"<div class=""item-bollino-verde""></div>"
		end if 
		if not mod_larochelle and Application("scopri_prezzi")=1 then
			
			stringa=stringa & "<div class=""item-price-container"">"
			if promozione=1  and sconto>0 then
			    prezzoarold=split(prezzoold,",")
			    
				stringa=stringa & "<span class=""old-price"">"&prezzoarold(0)&"<span class=""sub-price"">,"&prezzoarold(1)&"</span></span>"
			end if

				'Decisionale visualizza il prezzo
				if prezzo_visibile=0 or (prezzo_visibile=2 and session("vedi_prezzi")=1) then
					'Visualizza il prezzo
					prezzoar=split(prezzoar,",")   
					stringa=stringa & "<span class=""item-price"" >"&prezzoar(0)&"<span class=""sub-price"">,"&prezzoar(1)&"</span></span>"
				else 
					'Telefonare
						stringa=stringa & "<span class=""item-price""><span class=""sub-price"">"&traduci("noprezzo")&"</span></span>"
				end if
			stringa=stringa & "</div>"
		end if 'if not mod_larochelle
		if novita=1 then
			stringa=stringa & "<span class=""new-rect"">"&traduci("new")&"</span>"
		end if
		
		if promozione=1 and sconto>0 and rs("telefonare")=0 and Application("scopri_prezzi")=1 then
			stringa=stringa & "<span class=""discount-rect"">-"&testo_sconto&"</span>"
		end if
		if promozione=1 and sconto=0 then
			stringa=stringa & "<span class=""discount-trasv"">Promozione</span>"
		end if
		'if utente_andrea then stringa=stringa&"prezzo_nascosto:"&prezzo_nascosto
		stringa=stringa & "</div><!-- End .item-image -->"&_
		 "<div class=""item-meta-container"" style=""height:100px;"">"&_
		 "<h3 class=""item-name""><a href=""product.asp?idpro="&rs("idpro")&href_settore&"""><span itemprop=""name"">"&rs("articolo")&"</span></a></h3>"&_
		 "</div><!-- End .item-meta-container --></div><!-- End .item -->"
		if colonne=4 then
			 stringa=stringa&"</div><!-- End .col-md-3 -->"
		elseif colonne=3 then
			 stringa=stringa&"</div><!-- End .col-md-4 -->"
		end if
		box_rs_articoli=box_rs_articoli&stringa
		iRecordsShown = iRecordsShown + 1

		rs.MoveNext
	loop
	
end function




%>