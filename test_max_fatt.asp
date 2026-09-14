<%t_inizio=timer()%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/chk_piva_cf.asp" -->
<!--#include virtual="/paginazione.asp" -->
<%
	response.write "Normale:" &max_fatt(0,year(date()))&"<br>"
	response.write "Pa:" &max_fatt(1,year(date()))&"<br>"
	%>