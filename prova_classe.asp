<!--#include virtual="/pag_adm_ordini_inc_cl.asp" -->


<%


dim causale_ddt
set causale_ddt=new cl_Causale_Ddt
response.write "causale_ddt.elemento 4: "&causale_ddt.elemento(4)&"<br><br>"
response.write "<br>Nelementi:"&causale_ddt.n_elementi&"<br>"

n=causale_ddt.stampa_option

set causale_ddt=nothing
%>