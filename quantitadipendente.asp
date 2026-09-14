<%
'Verifica chiusure 30_11_2015
%>
<!--#include virtual="/setup.asp" -->
<%
	
iduser=session("iduser")
idfoglio=request("foglio")
iddip=request("iddip")






if request.form("salva")<>"" then
	
	
	Call aggiorna()
	
	end if


Set rs_dipendenti=conn.Execute("select * from dipendenti where iddip="&iddip)

foglio_spettanze=true




nominativodip=rs_dipendenti("cognome")&" "&rs_dipendenti("nome")


ordinaper=Request.Cookies("/foglio-spettanze.asp")("ordina")
if ordinaper="codice" then
	orderby=" codice"
	Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
	Response.Cookies(questofile)("ordina")=ordinaper
elseif ordinaper="articolo" then
	orderby="  articolo"
	Response.Cookies(questofile).Expires =dateadd("yyyy",5,Date())
	Response.Cookies(questofile)("ordina")=ordinaper
else
	orderby="  codice"
end if
	sql="select utenti_foglio_spettanze.* from utenti_foglio_spettanze where id="&idfoglio
	set rs_foglio=conn.Execute(sql)
	
	


sql="select * from prodotti where idpro in (  SELECT idpro FROM `preferiti` WHERE iduser="&iduser&" union distinct select idpro from utenti_spettanze where idfoglio="&idfoglio&") order by "&orderby

set rs_prodotti=conn.execute(sql)



%>
<!--#include virtual="/config/header_inc.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<section id="content">
	<div id="breadcrumb-container">
		<div class="container">
			<ul class="breadcrumb">
						<li><a href="/">Home</a></li>
						<li class=""><a href="/fogli-spettanze.asp">Fogli spettanze</a></li>
						<li class=""><a href="/foglio-spettanze.asp?idfoglio=<%=idfoglio%>">Foglio spettanze del <%=rs_foglio("creato")%></a></li>
				<li class="">Imposta quantità</li>
			</ul>
		</div>
	</div>
	<div class="container">
	    <form method="POST" action="<%=questofile%>" >
			<input type="hidden" name="iddip" value="<%=iddip%>">
			<input type="hidden" name="idfoglio" value="<%=idfoglio%>">
			<div class="row">
				<div class="col-md-12">
						<strong>Imposta quantità dipendente <%=nominativodip%></strong>
	
					<div class="table-responsive">									
						<table class="table checkout-table" >
							<thead>
								<tr>
									<th class="table-title">Prodotto</th>
									<th class="table-title">Quantità</th>
								</tr>
							</thead>
								
							<tbody>
							<%
							do while not rs_prodotti.EOF
								articolo="<a href=""/product.asp?idpro="&rs_prodotti("idpro")&""" target=""_blank"" tabindex=""-1""  >"&rs_prodotti("codice")&" "&rs_prodotti("articolo")&"</a>"
								
								
								sql="select utenti_spettanze.* from utenti_spettanze where idfoglio="&idfoglio&" and idpro= "&rs_prodotti("idpro")&" and iddip="&iddip
															set rs_spettanze=conn.execute(sql)
															if not rs_spettanze.eof then
																quantita=rs_spettanze("quantita")
															else
																quantita=0
															end if
															if quantita=0 then quantita=""

								
								
								testocella="<input type=""text"" name=""quantita_"&rs_prodotti("idpro")&"_"&iddip&""" class=""spett"" "&disabled&"value="""&quantita&""">"
							%>
								<tr>
									<td style="text-align: left;">
										<%=articolo%> <%if session("vedi_prezzi")=1 then %><span class="pull-right prezzo"><%=simbolo_valuta&rs_prodotti("prezzo")%></span><%end if %>
										</td>
									<td ><%=testocella%></td>
									
									
									
								</tr>
								<%
								rs_prodotti.MoveNext
								loop
								set rs_prodotti = Nothing
								%>
							</tbody>
						</table>
							
					</div><!-- End .table-reponsive -->
	
				</div><!-- End .col-md-12 -->
			</div><!-- End .row -->
			<div class="row" style="margin-top:10px;">
				<div class="col-md-2 col-sm-2 col-xs-12 " style="text-align: center;">
					<input type="submit" class="btn btn-info" name="salva" value="SALVA MODIFICHE"/>
				</div><!-- col-md-2 -->
			</div><!-- row -->
	    </form>

	</div><!-- End .container -->

</section><!-- End #content -->
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
    <!-- AGGIUNTE -->
<script>


</script>

    </body>
</html>

<%
	
	
	function aggiorna()
		iduser=session("iduser")
		idfoglio=request("idfoglio")
		iddip=request("iddip")

		sql="select * from prodotti where idpro in (  SELECT idpro FROM `preferiti` WHERE iduser="&iduser&" union distinct select idpro from utenti_spettanze where idfoglio="&idfoglio&") order by prodotti.codice"
				Set rs_spettanza = Server.CreateObject("ADODB.Recordset")

		set rs_prodotti=conn.execute(sql)
		cisonoquantita=false
		do while not rs_prodotti.EOF
			modifiche_row=""
			articolo="<b>"&rs_prodotti("codice")&"</b>"
			idpro=rs_prodotti("idpro")
								stringa=idpro&"_"&iddip


			        quantita=trim(request.form("quantita_"&stringa))
				    'm_log_txt=m_log_txt&", analizzo quantita_"&stringa&" trovato valore:"&quantita
					rs_spettanza.open "select * from  utenti_spettanze where idfoglio="&idfoglio&" and idpro="&idpro&" and iddip="&iddip,conn,3 ,3
					if quantita<>"" and quantita<>"0" and isnumeric(quantita) then
						cisonoquantita=true
						if rs_spettanza.eof then
							modifiche_row=modifiche_row&" aggiungo "&quantita&" per "&dipendente
							rs_spettanza.addnew
							rs_spettanza("idpro")=idpro
							rs_spettanza("iddip")=iddip
							rs_spettanza("iduser")=iduser
							rs_spettanza("quantita")=quantita
							rs_spettanza("idfoglio")=idfoglio
						else
							if cint(quantita)<>rs_spettanza("quantita") then
								modifiche_row=modifiche_row&" modifico "&quantita&" per "&dipendente
								rs_spettanza("quantita")=quantita
							end if
						end if
						rs_spettanza.update
					elseif quantita="0" or quantita="" then
						if not rs_spettanza.eof then
							modifiche_row=modifiche_row&" elimino per "&dipendente
							rs_spettanza.delete
						end if
					end if
					rs_spettanza.close
			if modifiche_row<>"" then
				m_log_txt=m_log_txt&"<br>"&articolo&":"&modifiche_row
			end if
 
		rs_prodotti.MoveNext
		Loop
		set rs_prodotti = Nothing


		
		response.write "aggiornato"
		response.redirect("foglio-spettanze.asp?idfoglio="&idfoglio&"#tabella")
		end function
	
	%>