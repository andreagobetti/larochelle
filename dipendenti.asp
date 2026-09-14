<%
'Verifica chiusure 30_11_2015
%>
<!--#include virtual="/setup.asp" -->
<%
if session("iduser") = "" then call login()
if request.querystring("dimetti")<>"" then

	iddip=request.querystring("dimetti")
	sql="update utenti_dipendenti set dimesso=1 where iduser="&session("iduserDipendenti")&" and iddip="&iddip
	conn.execute sql,n
	dipendente=ucase(conn.execute ("select concat_ws(' ',cognome,nome) from dipendenti where iddip="&iddip)(0))
	denominazione_v=get_denominazione(sessioniduser)
	call add2log("[utente="&sessioniduser&"]"&denominazione_v&"[/utente] ha dimesso il dipendente "&dipendente&" iddip:"&iddip,2)
	call inserisci_appunto("<a href=""pag_adm_user.asp?iduser="&sessioniduser&"""><b>"&denominazione_v&"</b></a> ha dimesso il dipendente "&dipendente,array(1,4))
end if
%>
<!--#include virtual="/config/header_inc.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<section id="content">
	<div id="breadcrumb-container">
		<div class="container">
			<ul class="breadcrumb">
				<li><a href="/">Home</a></li>
				<li class="active">Elenco dipendenti</li>
			</ul>
		</div>
	</div>
	<div class="container">
		<div class="row">
			<div class="col-md-12">
				<header class="content-title">
					<h1 class="title">Elenco dipendenti</h1>
					<p class="">E' consentito aggiungere nuovi dipendenti ed effettuare modifiche sui dipendenti aggiunti, non è consentito modificare i dati dei dipendenti esistenti o di quelli presi in carico nella gestione di ordini in corso.</p>
				</header>
				
				<div class="col-md-12 text-right">
											<a  class="btn btn-custom-2" href="dipendente.asp">Aggiungi dipendente</a>
				</div>
				<div class="xlg-margin"></div><!-- space -->
				<div class="table-responsive">									
					<table class="table checkout-table" >
						<thead>
							<tr>
								<th class="table-title">Cognome</th>
								<th class="table-title">Nome</th>
								<th class="table-title">Matricola</th>
								<th class="table-title"></th>
							</tr>
						</thead>
							
						<tbody>
						<%
						sql="select dipendenti.*, misure.data_rilievo, utenti_gradi.nome_grado, utenti_gradi.colore, utenti_gradi.colore_nominativo FROM (dipendenti left JOIN utenti_gradi ON dipendenti.grado = utenti_gradi.id) left join (SELECT iddip,max(data_rilievo) as data_rilievo FROM dipendenti_misure group by iddip  ) as misure ON dipendenti.iddip = misure.iddip inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.dimesso=0 and utenti_dipendenti.iduser="&session("iduserDipendenti")

						Set rs_dipendenti=conn.Execute(sql)
						do while not rs_dipendenti.EOF
						%>
							<tr>
								<td ><%=rs_dipendenti("cognome")%></td>
								<td ><%=rs_dipendenti("nome")%></td>
								<td ><%=rs_dipendenti("matricola")%></td>
								<td><span style="float: right;"><%if rs_dipendenti("nuovo")=1 then %><a class="btn btn-info" href="dipendente.asp?iddip=<%=rs_dipendenti("iddip")%>">Modifica</a>&nbsp; <%end if %></span></td>
								
								
								
							</tr>
							<%
							rs_dipendenti.MoveNext
							loop
							set rs_dipendenti = Nothing
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
<script>
function dimetti(iddip){
if (confirm('Vuoi dimettere questo dipendente?')) {
    location.href="dipendenti.asp?dimetti="+iddip;
} else {
    // Do nothing!
}

}

</script>

    </body>
</html>