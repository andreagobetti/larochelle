<!--#include virtual="/setup.asp" -->
<%   

sql="select prodotti_tags.IDtag, prodotti_tags.IDpro, Count(prodotti_tags.ID) AS ConteggioDiID, First(prodotti_tags.Id) AS PrimoDiId FROM prodotti_tags GROUP BY prodotti_tags.IDtag, prodotti_tags.IDpro HAVING (((Count(prodotti_tags.ID))>1));"
set rs_tag=conn.execute (sql)


		do while not rs_tag.eof
			conn.execute ("delete prodotti_tag.* from prodotti_tags where id="&rs_tag("primodiid"))




			rs_tag.movenext
		loop







%>
