<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
'Elimina le scadenze per le fatture saldate
sql="select sommadiincassi.SommaDiimporto, fatture.IDFat, fatture.Nfat, fatture.totale_fattura,fatture.data FROM fatture LEFT JOIN (select incassi.idfat, Sum(incassi.importo) AS SommaDiimporto FROM incassi GROUP BY incassi.idfat )  AS sommadiincassi ON fatture.IDFat = sommadiincassi.idfat  order by year(fatture.data) desc, nfat desc"
set rs=conn.execute (sql)
do while not rs.eof
	if not isnull(rs("sommadiimporto")) then
		if  cdbl(rs("totale_fattura"))>=cdbl(rs("sommadiimporto")) then
			sql = "delete FROM scadenze WHERE idfat="&rs("idfat")
			conn.execute (sql),n
			if n>0 then response.write "Eliminate "&n& " scadenze da idfat "&rs("idfat")&"<br>"
			n=0
		end if
	end if
	rs.movenext
loop
%>
