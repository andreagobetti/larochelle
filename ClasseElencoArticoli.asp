<%
'Sintassi

	'set modificheRS= (new ClasseModificheRS)(oper)
	'modifichers.leggi(rs)	
	'val=modificheRS.confronta(rs)	





Class ClasseElencoArticoli

	Private m_tabella
	Private m_id
	Private m_tipo
	
	
    Public Default Function Init(parameters)
	    Set Init = Me
	    
	    n_parametri=UBound(parameters)+1
        m_tabella = parameters(0)
        m_id = parameters(1)
        
        
        
        
	    if n_parametri=3 then
	    	m_tipo=parameters(2)
		end if 
		    

	    
	    
	    
	    
    End Function
    
    Private Sub Class_Terminate
	    if m_log then call add2log(m_log_classe,0)
	End Sub
	
	public sub elenco_articoli_head()
		%>
		<script>
			var aggiungi_a='<%=m_tabella%>';
			var aggiungi_a_id=<%=m_id%>;
		</script>
		<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1" id="tab_articoli">
		<thead id="tab_articoli_testa">
			<%if utente_andrea then %>
			<tr>
				<td colspan="6" >m_tabella:<%=m_tabella%> m_id:<%=m_id%> m_tipo:<%=m_tipo%> </td>
			</tr>
			<%end if %>
			<tr>
				<td colspan="6" class="ui-widget-header">Elenco articoli:<input type="button" id="mieidettagli" value="Dettagli"></td>
			</tr>
			<tr>
				<td align="center" class=testotabella style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;"><b>Codice</b></td>
				<td style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Articolo</b></td>
				<td align="center" class=testotabella style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;"><b>Prezzo<br>unitario</b></td>
				<td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Quantit&agrave;</b></td>
				<td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Sconto %</b></td>
				<td align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; font-size:11px;" class=testotabella><b>Prezzo<br>totale</b></td>
			</tr>
		</thead>
		<%
	end sub

	Public sub elenco_articoli()
		on error goto 0
		spettanze=0
		m_manca_numeri_di_serie=false
		m_numeri_di_serie=false
		mostra_giacenze=true
		
		'spettanze=0

	    m_log_txt=m_log_txt&",elenco_articoli"
	    if m_tabella="ordini" then
		    if ordine.campo("stato")>3 then
				mostra_giacenze=false
			end if
			spettanze=ordine.campo("spettanze")
			mostra_spostamento=ordine.campo("stato")<=3
			totale_merce_ordine=ordine.campo("totale_merce_ordine")
		elseif m_tabella="preventivi" then
			m_modifica_articoli=true
		elseif m_tabella="ordini_fornitori" then
			'spettanze=ordine_Dictionary.item("spettanze")
		end if
		
		
		
		
		%>
		
		
		
		
		<tbody id="tab_articoli_body">
		<%'CICLO SU ARTICOLI
			
        sql_ordine="select ordini_dett.*, magazzino.quantita_magazzino,magazzino.idmag, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, magazzino.db_ven, varianti_a.codicevara,varianti_a.variante_a, varianti_b.variante_b, ordini_dett_note.nota FRom (((magazzino RIGHT JoIN (ordini_dett left join ordini_dett_note on ordini_dett.iddett = ordini_dett_note.iddett) oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JoIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara ) left join prodotti on  magazzino.idpro = prodotti.idpro  where idord="&m_id &" order by ordine,iddett"
	
			
		if m_tabella="ddt" then
			if m_tipo="" then
				m_tipo=conn.execute("select tipo from ddt where idord="&m_id)
			end if
			if m_tipo="0" then
		        sql=sql_ordine
			elseif m_tipo="1" then
		        sql="select ordini_dett.*, magazzino.quantita_magazzino,magazzino.idmag, magazzino.Data_arrivo, magazzino.quantita_riordino, magazzino.quantita_ordinata, magazzino.db_ven, varianti_a.codicevara,varianti_a.variante_a, varianti_b.variante_b, ordini_dett_note.nota FRom (((magazzino RIGHT JoIN (ordini_dett left join ordini_dett_note on ordini_dett.iddett = ordini_dett_note.iddett) oN (magazzino.idvarb = ordini_dett.idvarb) AND (magazzino.idvara = ordini_dett.idvara) AND (magazzino.idpro = ordini_dett.idpro)) LEFT JoIN varianti_b oN ordini_dett.idvarb = varianti_b.IDvarb) LEFT JoIN varianti_a oN ordini_dett.idvara = varianti_a.IDvara ) left join prodotti on  magazzino.idpro = prodotti.idpro  inner join ddt_dett_ordini on ordini_dett.iddett = ddt_dett_ordini.iddett where idddt="&idddt&" order by ordine, iddett;"			
			elseif m_tipo="2" then

		        sql=sql_ordine
			end if

			
			
			
			
		else
			
        
        
	        sql=replace(sql_ordine,"ordini",m_tabella)
        
		end if        
        
        
        
		set rs=conn.execute(sql)
		
		tot_costo=0
		tot_no_costo=0
		tot_vendita=0
		
		do while not rs.eof
			selcs=""
			txt_ordinato=""
			txt_fornitori=""
			txt_ordinato2=""
			tr_ordini=false
			tr_fornitori=""
			rowspan=""
			quantita_ordinato=0
			stile_cella=""
			txt_spettanze=""
			if rs("codice_ordine")<>"" then
				riga_libera=false
			else
				riga_libera=true
			end if
			
			if m_tabella<>"ordini_fornitori" then
				costo_ordine=converti_typevar14(rs("costo_ordine"))
				if rs("codice_ordine")<>"" and costo_ordine>0 then
					tot_costo=tot_costo+costo_ordine*rs("quantita")
					tot_vendita=tot_vendita+cdbl(rs("totale_riga"))
				else
					anomalia_costo=anomalia_costo&"<br><span class='red'><b>"&rs("codice_ordine")&"</b> costo zero</span>"

				end if
			end if
		
		
		
			if m_tabella="ordini" then
				quantita_ordinato=clng(rs("quantita_ordinato"))
				if rs("flag1_ordine") then selcs="1"
				if rs("flag1_ordine") then txt_ordinato2="Ordinato"
				
				if clng(rs("quantita_ordinato"))>0 then
					set rs_ordinato=conn.execute("select ordini_fornitori.idord, ordini_fornitori.nord, ordini_fornitori.stato, ordini_dett_ordini_fornitori_dett.quantita from ordini_dett_ordini_fornitori_dett inner join ordini_fornitori_dett on ordini_dett_ordini_fornitori_dett.iddett_fornitori = ordini_fornitori_dett.iddett inner join ordini_fornitori on ordini_fornitori_dett.idord = ordini_fornitori.idord where iddett_ordini="&rs("iddett"))
					do while not rs_ordinato.eof
						if txt_ordinato<>"" then txt_ordinato=txt_ordinato&", "
						txt_ordinato=txt_ordinato& "Ordine fornitore "&rs_ordinato("nord")
						rs_ordinato.movenext
					loop
					txt_ordinato=rs("quantita_ordinato")&" in ordine: "&txt_ordinato
					tr_ordini=true
				end if
			end if

		
			if m_tabella="ordini_fornitori" then
			
				sql="select ordini.idord, ordini.nord, ordini.stato, ordini_dett_ordini_fornitori_dett.quantita from ordini_dett_ordini_fornitori_dett inner join ordini_dett on ordini_dett_ordini_fornitori_dett.iddett_ordini = ordini_dett.iddett inner join ordini on ordini_dett.idord = ordini.idord where iddett_fornitori="&rs("iddett")
				set rs_ordinato=conn.execute(sql)
				do while not rs_ordinato.eof
					if tr_fornitori<>"" then tr_fornitori=tr_fornitori&", "
					tr_fornitori=tr_fornitori&rs_ordinato("quantita")& " per Ordine "&rs_ordinato("nord")
					rs_ordinato.movenext
				loop
			end if
		
		
		
		
			if txt_ordinato<>"" then
				rowspan="2"
				tr_ordini=true
			end if 
			if tr_fornitori<>"" then
				rowspan="2"
			end if 
			'Anticipo spettanze per determinare stile cella
			'Elenco spettanze
			if mod_larochelle and m_tabella="ordini" and cint(rs("quantita"))>0 then
				'Cerco nelle spettanze
				txt_spettanze = dettaglio_spettanze( m_tabella, RS("iddett"), rs("quantita"),false,stile_cella)
			end if

			%>
	          <tr valign="middle" <%'if rs("modificato")=true then response.write ("bgcolor='#ffff99'")%> id="id_<%=rs("iddett")%>" style="border-top: 1px solid gray; padding-bottom:10px; padding-top:10px;" class="<%=stile_cella%>">
			      <td align="center" rowspan="<%=rowspan%>">
	              <%if m_modifica_articoli then
	              %>
	              
	              <input type="hidden" name="selcs<%=rs("iddett")%>" id="selcs_<%=rs("iddett")%>" class="selezione" value="<%=selcs%>">
	              <span class="div_icona"><span class="ui-icon ui-icon-triangle-1-s ui-corner-all left_menu bg-gray" id="<%=rs("iddett")%>" data-ordine="<%=rs("ordine")%>"></span></span>
	              <%end if %>
	              <%if mostra_spostamento then %>
	              <span class="div_icona"><span class="ui-icon ui-icon-arrowthick-2-n-s ui-corner-all handle bg-gray" ></span></span>
	              <%end if %>
	              <%if m_modifica_articoli then %>
	              <span class="div_icona" id="sel_<%=rs("iddett")%>"><%=txt_ordinato2%></span>
	              
	              <%end if%>
	              <%
		              if m_tabella="ordini" then
		           if cint(rs("numeri_di_serie"))>0 then %>
	              <span class="div_icona"><span class="ui-icon ui-icon-grip-dotted-horizontal ui-corner-all handle bg-gray" ></span></span>
	              <%
		              end if
		              end if %>
              <%
			txt="product.asp?idpro=" & RS("idpro")
		%>
              <b><a href="<%=txt%>" title="Clicca per le opzioni" class="menu_idpro" data-idmag="<%=RS("idmag")%>" data-idpro="<%=RS("idpro")%>"><%=codice_articolo_e_variante(  rs("codice_ordine"),rs("codicevara"))%></a></b>
            </td>
            <td >
            <%if rs("codice_ordine")<>"" or  m_modifica_articoli=false then%>
				<%=rs("articolo_ordine")%>
			<%else %>
				<textarea type="text" name="articolo<%=rs("iddett")%>" rows="1" class="chkchge autogrow"><%=rs("articolo_ordine")%></textarea>
			<%end if%>
            <%
	            	if rs("varianti_ordine")<>"" then response.write "<br>"&rs("varianti_ordine")
					if m_tabella="ordini" then
						if cint(rs("numeri_di_serie"))>0 then
						
							'Cerco nei seriali
							txt_seriale=""
							conteggio_seriali=0
							seriali_occorrenti=clng(rs("quantita")*rs("numeri_di_serie"))
							set seriali=conn.execute ("select numeri_seriali.* from numeri_seriali where iddett="&RS("iddett"))
							do while not seriali.eof 
								if txt_seriale<>"" then txt_seriale=txt_seriale&", "
								txt_seriale=txt_seriale&seriali("seriale")
								conteggio_seriali=conteggio_seriali+1				
								seriali.movenext
							loop
							if txt_seriale<>"" then 
								response.write "<br>Numeri di serie: "&txt_seriale
								m_numeri_di_serie=true
							end if
							if conteggio_seriali < seriali_occorrenti then
								response.write "<br><span class=""rosso"">"&seriali_occorrenti-conteggio_seriali &" numeri di serie mancanti</span>"
								m_manca_numeri_di_serie=true						
							end if
						
						end if
					end if

					response.write txt_spettanze
					'if m_tabella="ordini" then
						if rs("nota")<>"" then
							response.write "<br>Nota:<span id=""notad"&RS("iddett")&""" class=""notaarticolo"">"&rs("nota")&"</span>"
						end if
					'end if

					%></td>
            <td align="right" ><%
                if not m_modifica_articoli then %>
              <%=formatcurrency(rs("prezzo"),2)%>
              <%else %>
              <input type="text" name="prezzo<%=rs("iddett")%>" size="8" value="<%=FormatNumber(rs("prezzo"),2,,,0)%>" style="text-align:right;" class="chkchge">
              <%end if%></td>
            <%
				stile=""
				if mostra_giacenze and  rs("codice_ordine")<>"" then
					if m_tabella="ordini" then
						quantita=clng(rs("quantita"))
						
						if isnull(rs("quantita_magazzino") ) then
							stile=" color:blue;"
						else
							quantita_magazzino=clng(rs("quantita_magazzino"))
							
							if quantita>(quantita_magazzino+quantita_ordinato) then
								stile=" color:#ff0000;"
							elseif quantita>quantita_magazzino then
								stile=" color:#ff9900;"	'Arancione (Parte in ordine)
							else
								stile=" color:#66c466;" 'Verde (tutto a magazzino)
							end if
						end if
					end if
					giacenza="/"&CancellettoSeNull(rs("quantita_magazzino"))
					if quantita_ordinato>0 then giacenza=giacenza&"+"&quantita_ordinato
				else
					giacenza=""
				end if
				%>
            <td align="center" valign="middle" style="<%=stile%>">
	            
              <%if not m_modifica_articoli then %>
              <%=rs("quantita")%>
              <%else 
              if (spettanze=1 and cint(rs("quantita"))>0) then 
              %>
				<%=rs("quantita")%>
              <%
              else
              %>
              <input type="text" name="quantita<%=rs("iddett")%>" size="3" value="<%=rs("quantita")%>" style="text-align:right;">
			  <%end if %>
  			  <input type="hidden" name="iddett<%=rs("iddett")%>" value="SI">
              <%end if%><%=giacenza%>
              </td>
            <td align="center" valign="middle" >
              <%if not m_modifica_articoli then %>
              <%=FormatNumber(rs("sconto_prodotto"),2)%>
              <%else %>
              <input type="text" name="sconto_prodotto<%=rs("iddett")%>" size="4" value="<%=FormatNumber(rs("sconto_prodotto"),2)%>" style="text-align:right;">
              <%end if%></td>
            <td align="right" valign="middle" nowrap >&#8364; <span class="prezzo_totale"><%=formatnumber(totale_riga(rs("prezzo"),rs("sconto_prodotto"),rs("quantita")),2)%></span></td>
          </tr>
		  <%if tr_ordini then %>
		  <tr <%if rs("modificato")=true then response.write ("bgcolor='#ffff99'")%>>
			<td colspan="5">
			<%=soloio("idddt:"&rs("iddett")&", idmag:"&rs("idmag")&" , quantita_scalato_magazzino:"&rs("quantita_scalato_magazzino")&", quantita_ordinato:"&rs("quantita_ordinato"))%>
			<%=txt_ordinato %>			
			</td>
		</tr>
		  <%end if%>
		  <%if tr_fornitori<>"" then %>
		  <tr <%if rs("modificato")=true then response.write ("bgcolor='#ffff99'")%>>
			<td colspan="5">
			<%
			response.write tr_fornitori
			%>
			</td>
		</tr>
		  <%end if%>
		  
		  
		  <tr class="mieidettagli" >
		  <td colspan="6">Riga <%=rs("iddett")%>, <%if m_tabella<>"ordini_fornitori" then %>Prezzo di acquisto: <%=formatnumber(rs("costo_ordine"),2)%><%end if %><%if utente_andrea then %><span class="soloio"> idmag:<%=rs("idmag")%> idpro e var: <%=rs("idpro")%>|<%=rs("idvara")%>|<%=rs("idvarb")%><%if m_tabella="ordini" then %> db_ven:<%=rs("db_ven")%>numeri_di_serie:<%=rs("numeri_di_serie")%> quantita_scalato_magazzino: <%=rs("quantita_scalato_magazzino")%><br>tot_costo: <%=tot_costo%> tot_no_costo: <%=tot_no_costo%> tot_vendita: <%=tot_vendita%> in_ordine: <%=rs("in_ordine")%> pronto: <%=rs("pronto")%> consegnato: <%=rs("consegnato")%></span><%end if %>
		   <%end if %>
		  </td>
		  
		  </tr>
		  
		  
		 
          <%
            if vecchio_conteggio then
	            prezzo=cdbl(rs("prezzo"))
	            
				totord=totord+round((prezzo-(cdbl(rs("sconto_prodotto"))/100)*prezzo)*rs("quantita"),2)
	        else
				totord=totord+totale_riga(rs("prezzo"),rs("sconto_prodotto"),rs("quantita"))
			end if
            rs.movenext
			loop
			set rs = nothing
			%>
          </tbody>
          
          
          
		<%
			
			'http://it.wikihow.com/Calcolare-il-Margine-di-Contribuzione
			if tot_costo>0 then
				if tot_vendita=totale_merce_ordine  then
					m_txt_margine="Margine = &#8364;"&formatnumber(tot_vendita-tot_costo,2)&" Ricarico:"&formatnumber(calcola_ricarico(tot_vendita,tot_costo),0)&"%"
				else
					m_txt_margine="Margine calcolabile su &#8364;"&formatnumber(tot_vendita,2)&" = &#8364;"&formatnumber(tot_vendita-tot_costo,2)&" Ricarico:"&formatnumber(calcola_ricarico(tot_vendita,tot_costo),0)&"%"
				end if
			else
				m_txt_margine="Margine e ricarico non calcolabili, mancano prezzi di acquisto"
			end if
			
			if m_manca_numeri_di_serie then %>
			<script>
				$(function () {
					$(".stato2").addClass("ui-state-disabled");
					$("#aggiungi_seriali").show();
				});
			</script>
			
			
			
			<%end if 
		
	end sub	
		
		
	Public sub riga_pulsanti(modifica_articoli,sposta,ordina_fornitore)
		%>
		
		<tr valign="middle" >
            <td colspan="6" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
			<%if ordina_fornitore then %>
				<input type="submit" value="Ordina a fornitore" id="ordina_fornitore" class="oper_articoli">
			<%end if %>
			<%if modifica_articoli then %><input type="submit" value="Sposta o copia" id="sposta_articoli" class="oper_articoli"><%end if %>
			<%if modifica_articoli then%><span style="float: right;"><input type="button" value="Aggiungi articolo" id="puls_aggiungi_articolo">
			<input type="submit" name="Aggiorna_ordine" id="Aggiorna_ordine" value="Aggiorna ordine"></span>
			<%end if%>
            </td>
        </tr>

		
		
		<%
		
		
	end sub

	
End Class

%>