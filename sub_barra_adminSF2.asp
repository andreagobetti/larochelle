<%
	if request.querystring("soloio")="x" then
		if session("soloio")="" then
			session("soloio")="x"
		else
			session("soloio")=""
		end if
	end if
%>
<div class="ui-widget ui-state-default ui-corner-all nostampa"  style="text-align:center;">
  <table style="width: 100%; border-spacing: 0px;"   id="nav_links">
    <tr >
      <td ><ul class="sf-menu" >
          <li class="current"> <a href="/"><span class="<%if barra=1 then%>rosso<%end if%>">Home</span></a> </li>
        </ul></td>
	  <%if ha_il_permesso("Z1") then%>

      <td >
	  <ul class="sf-menu" >
          <li class="current"> <a href="pag_adm_main.asp">Amministra</a>
            <ul>
              <li> <a href="pag_adm_accessi.asp">Accessi utenti</a> </li>
              <li><a href="pag_adm_critici.asp">Registro</a></li>
              <li><a href="pag_adm_chktabelle.asp">Controllo integrit&agrave; tabelle</a></li>
              <li><a href="pag_adm_rep22.asp">Riavii server</a></li>
              <li><a href="pag_adm_elenco_report.asp">Elenco report</a></li>
			  <li><a href="pag_adm_setup.asp" target="_blank">Configurazione</a>
				<ul>
					<li class="current"><a href="pag_adm_testi_email.asp">testi nelle email</a></li>
					<li><a href="pag_adm_dizionario.asp">Gestione dizionario</a></li>
					<li><a href="pag_adm_banche.asp">Banche</a></li>
					<li><a href="pag_adm_ordini_risposte.asp">Risposte predefinite ordini</a></li>
					<li><a href="pag_adm_metodi_pagamento.asp">Metodi pagamento</a></li>
                </ul>
			  </li>
              <%if utente_andrea then%>
              <li><a href="pag_adm_database.asp" target="_blank">Database</a></li>
              <li><a href="/te8/" target="_blank">TE8</a></li>
              <li><a href="pag_adm_showtable.asp" target="_blank">showtable</a></li>
              <li><a href="pag_adm_cambiauser.asp" target="_blank">Cambia utente</a></li>
              <li><a href="pag_adm_copia_traduzioni.asp" target="_blank">Copia traduzioni</a></li>
              <li><a href="<%=questofile%>?soloio=x">Mostra Solo io</a></li>
			  <li><a href="#" target="_blank">Info...</a>
				<ul>
				  <li><a href="pag_adm_variabili.asp?oper=cookie">Cookie</a></li>
				  <li><a href="pag_adm_variabili.asp?oper=session">Session</a></li>
				  <li><a href="pag_adm_variabili.asp?oper=application">Application</a></li>
					<li ><a href="pag_server_info.asp">Server Information</a></li>
					<li><a href=" pag_server_info2.asp">ASPinfo</a></li>
					<li><a href="infophp.php">InfoPHP</a></li>
                </ul>
			  </li>
              <%end if%>
            </ul>
          </li>
        </ul>
		</td>
		<%end if %>
	  <%if ha_il_permesso("A6") then%>
      <td ><ul class="sf-menu" >
          <li class="current"> <a href="category.asp">Catalogo</a>
            <ul>
              <li> <a href="pag_adm_artic.asp">Elenco articoli</a> </li>
              <li> <a href="pag_adm_artic_ebay.asp">Articoli Ebay</a> </li>
              <li> <a href="pag_adm_articoli_mepa.asp">Articoli MEPA</a> </li>
              <li><a href="pag_adm_magazzino.asp">Magazzino</a>
                <ul>
                  <li class="current"><a href="pag_adm_movimenti.asp">Movimenti</a></li>
				<li class="current"><a href="pag_adm_rep09.asp">Inventario</a></li>
                </ul>
              </li>
              <li><a href="pag_adm_settori.asp">Settori</a></li>
              <li><a href="pag_adm_correlati.asp">Liste articoli correlati</a></li>
              <li ><a href="pag_adm_marca_modello.asp">Marca modello</a></li>
              <li ><a href="pag_adm_elenco_brand.asp">Elenco brand</a></li>
              <li ><a href="pag_adm_tags.asp">Tags</a></li>
              <li><a href="pag_adm_distinta_base.asp">Distinte basi</a></li>
              <li><a href="listini.asp">Listini</a></li>
              <li><a href="pag_adm_numeri_seriali.asp">Numeri di serie</a></li>
              <li class="current"> <a href="#">Altro</a>
                <ul>
                  <li ><a href="pag_adm_categorie_mepa.asp">Categorie MEPA</a></li>
				  <li><a href="pag_adm_upc_ean.asp">Codici UPC-EAN</a></li>
                  <li ><a href="pag_adm_varianti.asp">Liste varianti</a></li>
                  <li ><a href="pag_adm_aggiungi_a_ordine.asp">Aggiunta voci a ordine</a></li>
                  <li><a href="pag_adm_unitamisura.asp">Unit&agrave; di misura</a></li>
                  <li><a href="pag_adm_chk_img.asp">Controllo immagini</a></li>
                  <li><a href="pag_adm_attrezzature.asp">Attrezzature</a></li>
                  <li><a href="pag_adm_artic_ordina.asp">Riordino articoli</a></li>
                </ul>
              </li>
              <li> <a href="#">Report</a>
                <ul>
	                
                  <li><a href="pag_adm_excel.asp?oper=articoli">Esporta Excel</a></li>
                  <li><a href="pag_adm_rep09_excel.asp">Esporta inventario</a></li>
                  <li><a href="pag_adm_rep14.asp">Etichette articoli</a></li>
                  <li><a href="pag_adm_rep25.asp">Ricarico articoli</a></li>
                </ul>
              </li>
            </ul>
          </li>
        </ul></td>
		<%end if %>
      <td ><ul class="sf-menu" >
          <li class="current"> <a href="pag_adm_ordini.asp"><span class="<%if barra=4 then%>rosso<%end if%>">Ordini</span></a>
          
            <ul>
	            <%if mod_larochelle then %>
              <li> <a href="pag_adm_sostituzioni.asp">Sostituzioni</a></li>
              <%end if %>
              <li> <a href="pag_adm_preventivi.asp">Preventivi</a></li>
              <%if ha_il_permesso("D2") then%>
              <li><a href="pag_adm_ordini_fornitori.asp">Ordini a fornitore</a></li>
			  <%end if %>
              <%if ha_il_permesso("Z1") then%>
              <li> <a href="#">Report</a>
                <ul>
                  <li><a href="pag_adm_rep10.asp">Ordini mese</a></li>
                  <li><a href="pag_adm_rep24.asp">Incassi su ordini</a></li>
                  <li><a href="pag_adm_rep04.asp">Elenco incassi da ordini</a></li>
                  <li><a href="pag_adm_rep12.asp">Ordinato per cliente</a></li>
                  <li><a href="pag_adm_rep01.asp">Elenco spedizioni</a></li>
                </ul>
              </li>
              <li><a href="#">Altro</a>
                <ul>
                  <li><a href="pag_adm_rep19.asp">Transazioni PayPal</a></li>
                  <li><a href="pag_adm_ordini.asp?cercain=eliminati">ordini eliminati</a></li>
                  <li><a href="pag_adm_ordine_sposta.asp">Sposta ordini e fatture</a></li>
                </ul>
              </li>
			  <%end if %>
            </ul>
          </li>
        </ul></td>
	  <%if ha_il_permesso("B2") then%>

      <td ><ul class="sf-menu" >
          <li class="current"> <a href="pag_adm_ddt.asp"><span class="<%if barra=5 then%>rosso<%end if%>">DDT</span></a> </li>
        </ul></td>
      <td ><ul class="sf-menu" >
          <li class="current"> <a href="pag_adm_fatture.asp"><span class="<%if barra=6 then%>rosso<%end if%>">Fatture</span></a>
            <ul>
              <li><a href="pag_adm_fatture_for.asp">Fatture da fornitore</a> </li>
              <li><a href="pag_adm_note_credito.asp">Note di credito</a> </li>
              <li> <a href="#">Report</a>
                <ul>
                  <li><a href="pag_adm_rep07.asp">Fatture clienti</a></li>
                  <li><a href="pag_adm_rep18.asp">Fatture da fornitori</a></li>
                  <li><a href="pag_adm_rep13.asp">Scadenze riba clienti</a></li>
                  <li><a href="pag_adm_rep06.asp">Scadenze avere</a></li>
                  <li><a href="pag_adm_rep15.asp">Scadenze dare</a></li>
                  <li><a href="pag_adm_rep16.asp">Fatture cliente da saldare</a></li>
                  <li><a href="pag_adm_rep20.asp">Grafico entrate e uscite</a></li>
                  <li><a href="pag_adm_rep24.asp">Grafico bilancio</a></li>
                </ul>
              </li>
            </ul>
          </li>
        </ul></td>
	  <%end if%>
      <td ><ul class="sf-menu" >
          <li class="current"> <a href="pag_adm_user.asp"><span class="<%if barra=7 then%>rosso<%end if%>">Clienti</span></a>
            <ul>
			<%if ha_il_permesso("A5") then %>
              <li><a href="pag_adm_user.asp?cercain=fornitori">Fornitori</a> </li>
			<%end if %>
              <%if ha_il_permesso("Z1") then%>
			  <li><a href="pag_adm_email_newsl.asp">Email newsletter</a></li>
              <li> <a href="pag_adm_agenti.asp">Agenti</a></li>
              <li><a href="pag_adm_adm.asp">Gestione admin</a> </li>
              <%end if%>
              <%if ha_il_permesso("E5") then%>
              <li><a href="#">Altro</a>
                <ul>
	                <li><a href="pag_adm_rep11.asp">Esporta Excel</a></li>
	                <li><a href="pag_adm_rep17.asp">Elenco comuni clienti</a></li>
					<li><a href="pag_adm_tipologia.asp">Tabella tipolgia utenti</a></li>
					<li><a href="pag_adm_user_duplicati.asp">Ricerca duplicati</a></li>
					<li><a href="pag_adm_user_chkcfpiva.asp">Controllo cf e Piva</a></li>
					<li><a href="pag_adm_ordine_sposta.asp">Sposta ordini e fatture</a></li>
					<li><a href="pag_adm_ipbannati.asp">IP bannati</a></li>
					<%if mod_dip then %>
					<li><a href="pag_adm_gradi.asp">Gradi</a></li>
					<li><a href="pag_adm_dipendenti.asp">Dipendenti dimessi</a></li>
					<%end if %>
                </ul>
              </li>
              <%end if %>
            </ul>
          </li>
        </ul></td>
            <% 
	        if false then
		        if session("chk_note")="" then
			        check_msg=true
		    	elseif  DateDiff("s", session("chk_note"), application("ultima_nota")) >= 0 Then
	
			        check_msg=true
			    end if
		        if check_msg then
		        	set rs_msg=conn.execute ("SELECT Count(messaggi_utenti.id) AS ConteggioDiid FROM messaggi_utenti GROUP BY messaggi_utenti.data_lettura, messaggi_utenti.iduser HAVING (((messaggi_utenti.data_lettura) Is Null) AND ((messaggi_utenti.iduser)="&sessioniduser&"));")
		        	if not rs_msg.eof then
			        	session("n_msg")=rs_msg("ConteggioDiid")
		        	else
			        	session("n_msg")=""
		        	end if
		        	session("chk_note")=now()
		        	call add2log("Conteggio note utente "&sessioniduser,0)
	
		        end if
	        	if session("n_msg")<>"" then
	        		n_msg="<span class=""badge"">"&session("n_msg")&"</span>"
	        	end if
	
	        %>
	         <td><ul class="sf-menu" >
		         <li class="current">
	         	<a href="pag_adm_messaggistica.asp?oper=utenti">Da fare</a><%=n_msg%>
	         	
	         	<%if utente_andrea then response.write "<ul><li>refresh:"&check_msg&"<br>session:"&session("chk_note")&"<br>app:"&application("ultima_nota")&"</li></ul>" end if%>
		         </li>
	         </ul>
	         </td>
	         <% end if %>
			 
			 
			 <%
				if session("chache_appunti_time")="" then
			        check_msg=true
		    	elseif  DateDiff("s", session("chache_appunti_time"), application("time_ultimo_appunto")) >= 0 Then
			        check_msg=true
			    end if
		        if check_msg then
					
					
					N=conta_appunti()
					
	
		        end if
	        	if session("chache_appunti")<>"" then
	        		n_msg="<span class=""badge"">"&session("chache_appunti")&"</span>"
	        	end if

			 
			 
			 %>
			 
	         <td><ul class="sf-menu" >
		         <li class="current">
	         	<a href="pag_adm_appunti.asp">Appunti<%if session("chache_appunti")<>"" then response.write "<span class=""badge"" style=""vertical-align:top;"" id=""appuntidaleggere"">"&session("chache_appunti")&"</span>"%></a>
	         	
		         </li>
	         </ul>
	         </td>
	         
        
        
        
              <td ><input id='cercatuttoplaceholder' style="width:75px;" value="Cerca.." />
			  <input id='cercatutto' style="width:300px; display: none;" value="Cerca.."  />
              </td>
    </tr>
  </table>
<script src="Jquery/js/select2/select2.js"></script><script src="Jquery/js/select2/select2_locale_it.js"></script>
<script type="text/javascript">
$(document).ready(function () {
	var form = $('#cercatutto'), w = form.width();
		$("#cercatuttoplaceholder").click(function(){
		$("#cercatuttoplaceholder").animate({'width': '300px'}, 'fast',function(){attiva();});
	});
	$(form).on("select2-selecting", function(e) {
		var id = e.object.id;
		var t=e.object.t;
		//if(data!=null){t=data.t}
		if(t=="O")
		{
			location.href = "pag_adm_ordini.asp?idord="+id;
		}
		if(t=="F")
		{
				location.href = "pag_adm_fatture.asp?idfat="+id;
		}
		if(t=="U")
		{
				location.href = "pag_adm_user.asp?iduser="+id;
		}
		if(t=="D")
		{
				location.href = "pag_adm_ddt.asp?idddt="+id;
		}
		if(t=="P")
		{
				location.href = "product.asp?idpro="+id;
		}
		if(t=="D")
		{
				location.href = "pag_adm_user.asp?iduser="+id+"&iddip="+e.object.iddip;
		}
		if(t==""){
			//form.select2("destroy");
			form.animate({'width': w}, 'slow');
		}
	});
	
	$(form).on("select2-close", function() {
			//form.select2("destroy");
			//form.select2("container").hide();
			form.select2("container").hide();
			$("#cercatuttoplaceholder").show();
			$("#cercatuttoplaceholder").animate({'width': 75}, 'fast');
	});
	function attiva(){
	$("#cercatuttoplaceholder").hide();
		form.show();
		form.select2({
			placeholder: 'Cerca...',
			minimumInputLength: 1,
			//allowClear: true,
			ajax: {
				quietMillis: 150,
				url: "ajax_function.asp?select2=tutto",
				dataType: 'json',
				data: function (term, page) {
					return {
						term: term
					};
				},
				results: function (data) {
					
					return {
						results: data
					};
				}
			}
		});
		form.select2("container").find("div.select2-drop").append("<span>Inserisci un numero per cercare ordini, fatture e ddt, testo per cercare in articoli e clienti</span>");
		form.select2("open");
	}
});
  
  
</script>
</div>
