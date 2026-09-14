<%
	strConn = "server=62.149.150.165;uid=Sql582810;pwd=e6bebde5;database=Sql582810_2;driver=MySQL ODBC 3.51 Driver"

Set Conn = Server.CreateObject("ADODB.Connection")
conn.Open strConn

if err.Number = 0 then
  Response.write "<P>Connessione riuscita!"
else
  Response.write "<P>Errore: " & err.Description
end if
	%>