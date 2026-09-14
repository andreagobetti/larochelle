<!--#include virtual="/setup.asp" -->

<%   
'Imposta campo riba nella tabella incassi

sql="UPDATE utenti SET utenti.idloged = Null;"

conn.execute (sql)

rs.Open "impostazioni", conn, 3, 3

rs("id_loged")=1
rs.update
 rs.Close








%>
