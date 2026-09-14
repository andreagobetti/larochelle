<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc2.asp" -->

<%
response.write "totale_merce: 100 "&importo_trasporto(100)&"<br>"
response.write "totale_merce: 500 "&importo_trasporto(500)&"<br>"
response.write "totale_merce: 1000 "&importo_trasporto(1000)&"<br>"
response.write "totale_merce: 10000 "&importo_trasporto(10000)&"<br>"
%>
