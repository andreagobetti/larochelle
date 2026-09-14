<%
'Verifica chiusure 30_11_2015
%>
<!--#include virtual="/setup.asp" -->
<%
if session("iduser") = "" then call login()

idord=request("idord")
%>
<!--#include virtual="/config/header_inc.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<section id="content">
	<div id="breadcrumb-container">
		<div class="container">
			<ul class="breadcrumb">
				<li><a href="/">Home</a></li>
				<li class="active"><%=traduci("tit10")%></li>
			</ul>
		</div>
	</div>
	<div class="container">
		<div class="row">
			<div class="col-md-12">
				<header class="content-title">
					<h1 class="title"><%=traduci("tit10")%>&nbsp;<%=idord%></h1>
					<p class="title-desc"><%=traduci("tit10txt")%></p>
				</header>
				<div class="table-responsive">									
					<table class="table checkout-table" >
						<thead>
							<tr>
								<th class="table-title"><%=traduci("nord")%></th>
								<th class="table-title"><%=traduci("data")%></th>
								<th class="table-title"><%=traduci("impo")%></th>
								<th class="table-title"><%=traduci("stato")%></th>
								<th class="table-title"><%=traduci("pagato")%></th>
								<th class="table-title"><%=traduci("download")%></th>
							</tr>
						</thead>
							
						<tbody>
						<%
						'sql="select * FROM ordini where iduser="&session("iduser")& " and not eliminato"
						sql="select ordini.idord, ordini.nord, ordini.iduser, ordini.data, ordini.stato, ordini.tipopagamento, ordini.totale, sommadiincassi.SommaDiimporto FROM ordini LEFT JOIN (select incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idord  )  AS sommadiincassi ON ordini.idord = sommadiincassi.idord where iduser="&SessionIDUser& " and not eliminato order by data desc"
						
						sql="select ordini.idord, ordini.nord, ordini.iduser, ordini.data, ordini.stato, ordini.tipopagamento, ordini.totale, sommadiincassi.SommaDiimporto FROM ordini LEFT JOIN (select incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idord  )  AS sommadiincassi ON ordini.idord = sommadiincassi.idord where iduser="&SessionIDUser& " and not eliminato order by data desc"
						
sql="select ordini.idord, ordini.nord, ordini.iduser, ordini.data, ordini.stato, ordini.nome, ordini.cognome, ordini.azienda, ordini.regione, ordini.provincia, ordini.verde, utenti.email, utenti.cellulare, utenti.telefono, ordini.totale, ordini.indirizzo, ordini.citta, ordini_fatture.idfat, ordini.eliminato, ddt.IDddt, ddt.Nddt, fatture.Nfat, SUM(incassi.importo) AS SommaDiimporto, utenti.voto, utenti.tipologia FROM ((((ordini LEFT JOIN incassi ON ordini.idord = incassi.idord) LEFT JOIN utenti ON ordini.iduser = utenti.iduser) LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDFat) LEFT JOIN ddt ON ordini.idord = ddt.idord GROUP BY ordini.idord having ordini.iduser="&SessionIDUser& " and ordini.eliminato=0 order by data desc"
						sql="SELECT ordini.idord, ordini.nord, ordini.iduser, ordini.data, ordini.stato, ordini.tipopagamento, ordini.totale, sommadiincassi.SommaDiimporto, ddt.IDddt, ddt.Nddt, ddt.data as ddt_data, fatture.IDFat, fatture.Nfat, fatture.data as fatture_data FROM (((ordini LEFT JOIN (select incassi.idord, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idord  )  AS sommadiincassi ON ordini.idord = sommadiincassi.idord) LEFT JOIN ddt ON ordini.idord = ddt.idord) LEFT JOIN ordini_fatture ON ordini.idord = ordini_fatture.idord) LEFT JOIN fatture ON ordini_fatture.idfat = fatture.IDFat where ordini.iduser="&SessionIDUser& " and ordini.eliminato=0 order by data desc"

'response.write sql
						
						
						Set rs_ordini=conn.Execute(sql)
						do while not rs_ordini.EOF
						SommaDiimporto=rs_ordini("SommaDiimporto")
						if isnull(SommaDiimporto) then
							SommaDiimporto=0
						else
							SommaDiimporto=CDbl(SommaDiimporto)
						end if
						if cdbl(rs_ordini("totale"))>SommaDiimporto then
							pagato=false
						else
							pagato=true
						end if
						%>
							<tr>
								<td class=""><span class="order-number"><a href="order.asp?idord=<%=rs_ordini("idord")%>"><%=rs_ordini("nord")%></a></span></td>
								<td class="item-price-col"><span class="item-price-special"><%=FormatDateTime(rs_ordini("data"),2)%></span></td>
								<td class="item-price-col" style="white-space: nowrap;"><span class="item-price-special"><%=simbolo_valuta&FormatNumber(rs_ordini("totale"),2)%></span></td>
								<td class="item-name-col"><%=stato_ordine(rs_ordini("stato"))%></td>
								<td class="item-name-col"><%if pagato=false then%>
								In attesa di pagamento
								<%
								if Application("codice_paypal")<>"" then
								if rs_ordini("tipopagamento")=99 then
									call add2log("tipopagamento=99 idord:"&rs_ordini("idord"),0)
								end if
								if rs_ordini("stato")>=3 and   instr(lcase(metodo_pagamento(rs_ordini("tipopagamento"),"descrizione")),"paypal")>0 and  rs_ordini("stato")<>7  then%>
  
	  <form name="_xclick" action="https://www.paypal.com/it/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="<%=Application("codice_paypal")%>">
<input type="hidden" name="currency_code" value="EUR">
<input type="hidden" name="lc" value="IT">
<input type="hidden" name="item_name" value="Ordine <%=rs_ordini("nord")%>">
<input type="hidden" name="item_number" value="<%=rs_ordini("idord")%>">
<input type="hidden" name="amount" value="<%=replace(rs_ordini("totale"),",",".")%>">
<input name="notify_url" type="hidden" value="http://<%= lcase(nomesito)%>/pag_paypal_ipn.asp" />
<input type="image" src="https://www.paypalobjects.com/it_IT/i/btn/x-click-but6.gif" border="0" name="submit" alt="Effettua il pagamento con PayPal. È un sistema rapido, gratuito e sicuro.">
</form>
  <%end if
  end if%>
								
								
								
								<%else%>Pagato<%end if%></td>
							<td class="item-name-col"><%if not isnull(rs_ordini("idfat")) then%> <a href="pdf_fattura.asp?idfat=<%=rs_ordini("idfat")%>&id2=<%=chkDataUser(rs_ordini("fatture_data"))%>" class="btn btn-custom-2 btn-sm" ><%=traduci("fattura")&" "&rs_ordini("nfat")&" del "&FormatDateTime(rs_ordini("fatture_data"))%> </a> <%end if%><%if not isnull(rs_ordini("idddt")) then %> <a href="pdf_ddt.asp?idddt=<%=rs_ordini("idddt")%>&id2=<%=chkDataUser(rs_ordini("ddt_data"))%>" class="btn btn-custom-2 btn-sm"><%=traduci("ddt")&" "&rs_ordini("nddt")&" del "&FormatDateTime(rs_ordini("ddt_data"),2)%> </a> <%end if%></td>

								
								
								
							</tr>
							<%
							rs_ordini.MoveNext
							loop
							set rs_ordini = Nothing
							%>
						</tbody>
					</table>
						
				</div><!-- End .table-reponsive -->

			</div><!-- End .col-md-12 -->
		</div><!-- End .row -->
	</div><!-- End .container -->

</section><!-- End #content -->
	<!--#include virtual="/footer_inc.asp" -->
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->
    <!-- AGGIUNTE -->


    </body>
</html>