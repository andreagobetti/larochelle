<!--#include virtual="/setup.asp" -->
<!--#include file="jsonObject.class.asp"-->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<%
	iduser=request("iduser")
	
	
	
	
%>
<!--#include virtual="/sub_head_adm.asp" -->
<style>
/* landscape styles */
 th {
  position: relative;
  padding: 10px;
}
th span {
  transform-origin: 0 50%;
  transform: rotate(-90deg); 
  white-space: nowrap; 
  display: block;
  position: absolute;
  bottom: 0;
  left: 50%;
}
tr {
	border-bottom:1px solid;
}
.celladipendente{
	font-weight: bold;
	white-space: nowrap;
}
.denominazione{
	font-weight: bold;
	font-size: large;
}
.grado{
font-size: xx-small;
font-weight: normal;
white-space: normal;
}
#accessori{
	border: 1px solid black;
	padding: 10px;
}
#accessori label {
	float: left;
	width: 150px;
	margin-bottom: 5px;
	font-weight: bold;
	
}
#accessori input[type="text"] {
	/*width: 500px;*/
	margin-bottom: 5px;
}
/*tabella rilievo taglie */
#accessori input.taglie[type="text"] {
	width: 60px;
	margin-bottom: 5px;
}
#accessori input[type="radio"] {
	width: 30px;
	margin-bottom: 5px;
	margin-right: 10px;
}
#accessori input[type="checkbox"] {
	margin-bottom: 5px;
	margin-left: 20px;
}
#accessori textarea {
	width: 550px;
	height: 50px;
	margin-bottom: 5px;
}
#accessori br {
	clear: left;
}
.option{
	padding-left: 20px;
	padding-right: 20px;
	padding-top: 3px;
	padding-bottom: 3px;
}
.green{
	background-color: #99FF99;
	margin-left: 2px;
}
.articolo_ordine{
	border: 1px solid black;
}
.quantita{
	border: 1px solid black;
	text-align: center;
}
</style>
</head>
<body> 
<div id="wrap">
	<div id="header">
		<%=titolo_top%>
		<!-- Div Content INIZIO-->
		<%barra=0%>
		<%
			if idord<>"" then
				set rs_ordine=conn.execute ("select ordini.nord, ordini.iduser from ordini where idord="&idord)	 
				iduser=rs_ordine("iduser")
				nord=rs_ordine("nord")
			    set rs_ordine = Nothing
			 end if
		%>
		<!--#include virtual="/sub_barra_adminsf2.asp" -->
	    <div class="ui-widget-header ui-corner-all titolo_admin"><a href="pag_adm_ordini.asp?idord=<%=idord%>">Elenco taglie <%=nord%></a></div>
	    <%
		    
	    %>
	    <div class="ui-widget-header denominazione" style="padding: 5px;">
		    <%
			    response.write get_denominazione(iduser)
			    %>
	    </div>
	    <%
		    set rs_dipendenti=conn.execute ("select dipendenti.iddip, dipendenti.nominativo from dipendenti inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.iduser="&iduser&" order by sesso desc, cognome, nome")
		    arrRS=rs_dipendenti.getrows()
		    ubound_arrRS=ubound(arrRS,2)
	    %>
	    <form method="POST" action="<%=questofile%>">
			<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1" id="spettanze">
				<thead >
			            <tr>
				            <td colspan="1"></td>
				            <td colspan="1" align="center" bgcolor="#a8dcf0">GIACCA<br>Tecnica</td>
				            <td colspan="7" align="center" bgcolor="#87CEEB">GIACCA</td>
				            <td colspan="1" align="center" bgcolor="#66c1e5">Cappotto</td>
				            
				            <td  colspan="1" align="center" bgcolor="#F2F5A9">PANTALONI<br>Tecnici</td>
				            <td  colspan="7" align="center" bgcolor="#FFFF00">PANTALONI</td>
				            <td  colspan="1" align="center" bgcolor="#F4FA58">GONNA</td>
				            <td  colspan="1" align="center" bgcolor=""></td>
				            <td  colspan="1" align="center" bgcolor=""></td>
				            <td  colspan="1" align="center" bgcolor=""></td>
				            <td  colspan="1" align="center" bgcolor=""></td>
				            <td  colspan="1" align="center" bgcolor=""></td>
				            <td  colspan="1" align="center" bgcolor=""></td>
				            <td  colspan="1" align="center" bgcolor=""></td>
			            </tr>
			            <tr>
				            <td colspan="1"></td>
				            <td class="cella-misura-head">Taglia</td>
				            <td class="cella-misura-head">TG</td>
				            <td class="cella-misura-head">Fondo</td>
				            <td class="cella-misura-head">Manica</td>
				            <td class="cella-misura-head">Torace</td>
				            <td class="cella-misura-head">Vita</td>
				            <td class="cella-misura-head">Bacino</td>
				            <td class="cella-misura-head">Spalle</td>
				            <td class="cella-misura-head">Lunghezza</td>
				            
				            <td class="cella-misura-head">Tecnici</td>
				            <td class="cella-misura-head">TG</td>
				            <td class="cella-misura-head">Lunghezza</td>
				            <td class="cella-misura-head">Vita</td>
				            <td class="cella-misura-head">Bacino</td>
				            <td class="cella-misura-head">Cosce</td>
				            <td class="cella-misura-head">Cavallo</td>
				            <td class="cella-misura-head">Polpacci</td>
				            <td class="cella-misura-head">Lunghezza</td>
				            <td class="cella-misura-head">Scarpa</td>
				            <td class="cella-misura-head">Berretto</td>
				            <td class="cella-misura-head">Guanto</td>
				            <td class="cella-misura-head">Cintura</td>
				            <td class="cella-misura-head">Cinturone</td>
				            <td class="cella-misura-head">Maglieria</td>
				            <td class="cella-misura-head">Arma</td>
			            </tr>
				</thead>
				<tbody>
					<%
					
						sql="select *, utenti_gradi.nome_grado, utenti_gradi.colore, utenti_gradi.colore_nominativo from  dipendenti left join utenti_gradi on dipendenti.grado = utenti_gradi.id left join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.iduser="&iduser&" order by sesso desc, cognome, nome"
						set rs_dipendenti=conn.execute (sql)
						do while not rs_dipendenti.eof 
							iddip=rs_dipendenti("iddip")
							set rs_rilievi=conn.execute ("select dipendenti_misure.* from dipendenti_misure where iddip="&iddip&" order by data_rilievo desc limit 0,1")
						%>
						<tr>
							<td colspan="1" class="celladipendente">
								<a href="#" class="dipendente" id="dipendente<%=rs_dipendenti("iddip")%>"><%=dipendente_colorato(rs_dipendenti("cognome")&" "&rs_dipendenti("nome"),rs_dipendenti("sesso"),rs_dipendenti("colore_nominativo"))%></a><%=" "&rs_dipendenti("matricola")%><br><span class="grado"><%=grado_colorato(rs_dipendenti("nome_grado"),rs_dipendenti("colore"))%></span>
							</td>
							<%
								
								if not rs_rilievi.eof then

								
								%>
							
							
				            <td class="cella-misura"><%=rs_rilievi("giacca_tecnici")%></td>
				            <td class="cella-misura"><%=rs_rilievi("giacca_TG")%><br><%=rs_rilievi("giacca_tg2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("giacca_fondo")%><br><%=rs_rilievi("giacca_fondo2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("giacca_manica")%><br><%=rs_rilievi("giacca_manica2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("giacca_Torace")%><br><%=rs_rilievi("giacca_Torace2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("giacca_vita")%><br><%=rs_rilievi("giacca_vita2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("giacca_bacino")%><br><%=rs_rilievi("giacca_bacino2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("giacca_spalle")%><br><%=rs_rilievi("giacca_spalle2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("cappotto")%><br><%=rs_rilievi("cappotto2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("pantaloni_tecnici")%></td>
				            <td class="cella-misura"><%=rs_rilievi("pantaloni_tg")%><br><%=rs_rilievi("pantaloni_tg2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("pantaloni_lunghezza")%><br><%=rs_rilievi("pantaloni_lunghezza2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("pantaloni_vita")%><br><%=rs_rilievi("pantaloni_vita2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("pantaloni_bacino")%><br><%=rs_rilievi("pantaloni_bacino2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("pantaloni_cosce")%><br><%=rs_rilievi("pantaloni_cosce2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("pantaloni_cavallo")%><br><%=rs_rilievi("pantaloni_cavallo2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("pantaloni_polpacci")%><br><%=rs_rilievi("pantaloni_polpacci2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("gonna")%><br><%=rs_rilievi("gonna2")%></td>
				            <td class="cella-misura"><%=rs_rilievi("scarpa")%></td>
				            <td class="cella-misura"><%=rs_rilievi("berretto")%></td>
				            <td class="cella-misura"><%=rs_rilievi("guanto")%></td>
				            <td class="cella-misura"><%=rs_rilievi("cintura")%></td>
				            <td class="cella-misura"><%=rs_rilievi("cinturone")%></td>
				            <td class="cella-misura"><%=rs_rilievi("maglieria")%></td>
				            <td class="cella-misura"><%=rs_dipendenti("arma")%>
				            <%if rs_dipendenti("nota_dipendente")<>"" then %>
				            <br><%=rs_dipendenti("nota_dipendente")%>
				            <%end if %>
				            
				            </td>
				            <%
        						else
							%>
								<td colspan="25" class="cella-misura">Nessun rilievo misurre</td>
								
							
							
							<%
							
						end if	

				            
				 %>           
				            
				            
				            
						</tr>
						<%
							
							rs_dipendenti.MoveNext
							loop
							set rs_dipendenti = Nothing
							set rs_rilievi = Nothing
							
							%>
				</tbody>
			</table>
<%
	call tabella_accessori(iduser)
	call connclose()
	%>
			
			
	    </form>
        <!-- Div Content FINE-->
	</div><!-- End #header -->
<div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" -->
</div><!-- End #wrap -->
<script>
	$(function() {
    var header_height = 0;
    $('table th span').each(function() {
        if ($(this).outerWidth() > header_height) header_height = $(this).outerWidth();
    });
    $('table th').height(header_height);
});
	</script>
</body>
</html>
<%
	function green(val)
		green=""
		if val=1 then green="green"
	end function
%>